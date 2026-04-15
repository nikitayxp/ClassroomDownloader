#!/usr/bin/env python3
"""
Google Classroom Downloader
Script para baixar todos os documentos do Google Classroom organizados por disciplina.

Requisitos:
    pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client

Uso:
    python classroom_downloader.py
"""
import os
os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'

import re
import json
import pickle
import sys
from datetime import datetime
from pathlib import Path

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

try:
    import winreg
except ImportError:
    winreg = None

# Scopes necessários para acessar o Classroom e Drive
SCOPES = [
    'https://www.googleapis.com/auth/classroom.courses.readonly',
    'https://www.googleapis.com/auth/classroom.coursework.me.readonly',
    'https://www.googleapis.com/auth/classroom.courseworkmaterials.readonly',
    'https://www.googleapis.com/auth/classroom.announcements.readonly',
    'https://www.googleapis.com/auth/drive.readonly'
]

# Pasta onde os arquivos serão salvos
DOWNLOAD_DIR = Path.home() / "Classroom_Downloads"

# Arquivos de credenciais
CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.pickle"
CONFIG_FILE = Path(__file__).resolve().parent / "config.json"

# Configuração de autoarranque no Windows
WINDOWS_STARTUP_RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
WINDOWS_STARTUP_VALUE_NAME = "ClassroomDownloader"

# Status de operações de escrita/download
RESULT_DOWNLOADED = "downloaded"
RESULT_SKIPPED = "skipped"
RESULT_ERROR = "error"


def sanitize_filename(filename):
    """Normaliza um nome para uso seguro como ficheiro no Windows.

    Remove caracteres inválidos, substitui quebras de linha por espaço,
    remove pontos/espaços no final e limita o tamanho total a 150 caracteres.
    """
    filename = str(filename or "")
    # Substitui quebras de linha para evitar erro 22 no Windows.
    filename = filename.replace('\n', ' ').replace('\r', ' ')
    # Remove caracteres que não são permitidos em nomes de arquivo
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    # Normaliza múltiplos espaços e remove espaços/pontos no fim
    filename = re.sub(r'\s+', ' ', filename).strip().rstrip(' .')
    # Limita o tamanho do nome
    if len(filename) > 150:
        name, ext = os.path.splitext(filename)
        if ext and len(ext) < 150:
            filename = name[:150 - len(ext)].rstrip(' .') + ext
        else:
            filename = filename[:150].rstrip(' .')
    return filename or "arquivo"


def get_publication_name(title, text, date_str, default_prefix):
    """Gera um nome base determinístico para uma publicação.

    Prioriza título, depois excerto do texto e por fim data/prefixo de fallback.
    """
    if title and title.strip():
        return sanitize_filename(title)
    if text and text.strip():
        first_line = text.strip().splitlines()[0]
        return sanitize_filename(first_line[:60])
    if date_str:
        return sanitize_filename(f"{default_prefix}_{date_str[:10]}")
    return sanitize_filename(default_prefix)


def load_config():
    """Lê o ficheiro de configuração e devolve um dicionário válido.

    Retorna None quando o ficheiro não existe ou está inválido para que o
    fluxo de inicialização recrie a configuração com defaults seguros.
    """
    if not os.path.exists(str(CONFIG_FILE)):
        return None

    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)

        if not isinstance(config, dict):
            raise ValueError("Estrutura inválida para config.json")

        return config
    except Exception as e:
        print(f"⚠️ Erro ao ler config.json: {str(e)}")
        print("⚠️ Um novo arquivo de configuração será recriado.")
        return None


def save_config(config):
    """Persiste a configuração do script em config.json."""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"⚠️ Erro ao salvar config.json: {str(e)}")
        return False


def get_startup_python_executable():
    """Resolve o executável Python mais adequado para arranque automático.

    Usa pythonw.exe quando disponível para execução em background e faz
    fallback para o executável atual quando necessário.
    """
    python_executable = Path(sys.executable)
    pythonw_executable = python_executable.with_name('pythonw.exe')

    if os.path.exists(str(pythonw_executable)):
        return str(pythonw_executable)

    return str(python_executable)


def register_startup_windows(show_success=True):
    """Regista o script no arranque do Windows via HKCU\\...\\Run.

    Retorna True quando o valor é gravado com sucesso no Registo e False em
    qualquer falha, sem interromper o fluxo principal da aplicação.
    """
    if os.name != 'nt' or winreg is None:
        print("⚠️ Autoarranque via Registro só é suportado no Windows.")
        return False

    try:
        script_path = str(Path(__file__).resolve())
        python_executable = get_startup_python_executable()
        startup_command = f'"{python_executable}" "{script_path}"'

        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, WINDOWS_STARTUP_RUN_KEY) as key:
            winreg.SetValueEx(key, WINDOWS_STARTUP_VALUE_NAME, 0, winreg.REG_SZ, startup_command)

        if show_success:
            print("✅ Autoarranque no Windows ativado.")

        return True
    except Exception as e:
        print(f"⚠️ Não foi possível configurar autoarranque: {str(e)}")
        return False


def ask_startup_preference():
    """Pergunta ao utilizador se deve ativar autoarranque no Windows.

    Aceita apenas S ou N e repete até obter uma resposta válida.
    """
    prompt = "Deseja que este script procure novos ficheiros automaticamente sempre que iniciar o Windows? (S/N) "

    while True:
        answer = input(prompt).strip().upper()

        if answer in ('S', 'N'):
            return answer == 'S'

        print("Resposta inválida. Digite apenas S ou N.")


def initialize_config():
    """Inicializa a configuração da aplicação e estado de autoarranque.

    Em primeira execução solicita preferência de autoarranque, persiste a
    decisão em config.json e mantém metadados de versão e última execução.
    """
    config = load_config()
    now_iso = datetime.now().isoformat(timespec='seconds')

    if config is None:
        enable_auto_run = ask_startup_preference()
        startup_registered = False

        if enable_auto_run:
            startup_registered = register_startup_windows(show_success=True)

        config = {
            'version': 1,
            'first_run_completed': True,
            'auto_run_on_startup': enable_auto_run,
            'startup_registered': startup_registered,
            'created_at': now_iso,
            'last_run': now_iso,
        }

        save_config(config)
        return config

    if config.get('auto_run_on_startup'):
        config['startup_registered'] = register_startup_windows(show_success=False)

    config['version'] = config.get('version', 1)
    config['first_run_completed'] = True
    config['last_run'] = now_iso

    save_config(config)
    return config


def save_publication_text(destination_folder, publication_name, body_text):
    """Guarda o corpo textual da publicação num ficheiro .txt local.

    Se o ficheiro já existir, devolve status de skip para suportar execução
    incremental sem reescrita desnecessária.
    """
    try:
        safe_publication_name = sanitize_filename(publication_name or "Publicacao")
        txt_file_name = sanitize_filename(f"{safe_publication_name} - Texto.txt")
        txt_path = destination_folder / txt_file_name

        if os.path.exists(str(txt_path)):
            print(f"    ⏭️ Ficheiro {txt_file_name} já existe. A saltar...")
            return RESULT_SKIPPED

        content = (body_text or "").strip()
        if not content:
            content = "Sem texto/descricao disponivel."

        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(content + "\n")

        print(f"    📝 Texto salvo: {txt_file_name}")
        return RESULT_DOWNLOADED
    except Exception as e:
        print(f"    ❌ Erro ao salvar texto da publicacao: {str(e)}")
        return RESULT_ERROR


def get_credentials():
    """Obtém credenciais OAuth válidas para acesso ao Google Classroom/Drive.

    Reutiliza token local quando possível, tenta refresh automático e inicia
    fluxo de autenticação interativo quando necessário.
    """
    creds = None
    
    # Carrega token existente
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    
    # Se não houver credenciais válidas, faz login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"❌ Arquivo '{CREDENTIALS_FILE}' não encontrado!")
                print("\n📋 Para usar este script, você precisa:")
                print("1. Acesse: https://console.cloud.google.com/")
                print("2. Crie um projeto ou use um existente")
                print("3. Ative a API do Google Classroom e Google Drive")
                print("4. Crie credenciais OAuth 2.0 (tipo 'Desktop app')")
                print("5. Baixe o arquivo JSON e renomeie para 'credentials.json'")
                print("6. Coloque o arquivo na mesma pasta deste script")
                return None
            
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Salva o token para próximas execuções
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
    
    return creds


def download_file_from_drive(service_drive, file_id, file_name, destination_folder):
    """Transfere um ficheiro do Drive para a pasta de destino.

    Suporta exportação de ficheiros Google Workspace, aplica naming seguro e
    evita redownload quando o destino já existe localmente.
    """
    try:
        # Obtém informações do arquivo
        file_metadata = service_drive.files().get(fileId=file_id, fields='name,mimeType').execute()
        
        # Sanitiza o nome do arquivo
        safe_name = sanitize_filename(file_name or file_metadata['name'])
        
        # Verifica se é um Google Workspace file (Docs, Sheets, etc.)
        mime_type = file_metadata['mimeType']
        export_mime_types = {
            'application/vnd.google-apps.document': ('application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.docx'),
            'application/vnd.google-apps.spreadsheet': ('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.xlsx'),
            'application/vnd.google-apps.presentation': ('application/vnd.openxmlformats-officedocument.presentationml.presentation', '.pptx'),
            'application/vnd.google-apps.drawing': ('image/png', '.png'),
        }
        
        destination_path = destination_folder / safe_name
        export_mime = None
        
        if mime_type in export_mime_types:
            # Exporta Google Workspace files para formato Office
            export_mime, extension = export_mime_types[mime_type]
            if not destination_path.suffix:
                safe_name = sanitize_filename(f"{safe_name}{extension}")
                destination_path = destination_folder / safe_name

        if os.path.exists(str(destination_path)):
            print(f"    ⏭️ Ficheiro {safe_name} já existe. A saltar...")
            return RESULT_SKIPPED
            
        if export_mime:
            request = service_drive.files().export_media(fileId=file_id, mimeType=export_mime)
        else:
            # Download direto para outros arquivos
            request = service_drive.files().get_media(fileId=file_id)
        
        # Faz o download em chunks
        from googleapiclient.http import MediaIoBaseDownload
        import io
        
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        
        while not done:
            status, done = downloader.next_chunk()
        
        # Salva o arquivo
        fh.seek(0)
        with open(destination_path, 'wb') as f:
            f.write(fh.read())
        
        print(f"    ✅ {safe_name}")
        return RESULT_DOWNLOADED
        
    except Exception as e:
        print(f"    ❌ Erro ao baixar {file_name}: {str(e)}")
        return RESULT_ERROR


def download_link_content(url, file_name, destination_folder):
    """Guarda um link externo como atalho .url no sistema local.

    Em modo incremental, ignora a criação quando o ficheiro de destino já
    existe e retorna status de skip.
    """
    try:
        safe_name = sanitize_filename(f"{file_name}.url")
        destination_path = destination_folder / safe_name

        if os.path.exists(str(destination_path)):
            print(f"    ⏭️ Ficheiro {safe_name} já existe. A saltar...")
            return RESULT_SKIPPED
        
        # Cria um arquivo .url (atalho do Windows)
        with open(destination_path, 'w', encoding='utf-8') as f:
            f.write(f"[InternetShortcut]\n")
            f.write(f"URL={url}\n")
        
        print(f"    🔗 Link salvo: {safe_name}")
        return RESULT_DOWNLOADED
        
    except Exception as e:
        print(f"    ❌ Erro ao salvar link {file_name}: {str(e)}")
        return RESULT_ERROR


def process_materials(service_classroom, service_drive, course_id, course_name, course_folder):
    """Processa anúncios, trabalhos e materiais de um curso.

    Guarda textos das publicações, descarrega anexos e devolve contadores de
    itens novos e itens já existentes (modo incremental).
    """
    downloaded_count = 0
    skipped_count = 0

    def register_result(result):
        nonlocal downloaded_count, skipped_count
        if result == RESULT_DOWNLOADED:
            downloaded_count += 1
        elif result == RESULT_SKIPPED:
            skipped_count += 1
    
    try:
        # Cria subpastas
        anuncios_folder = course_folder / "Anuncios"
        trabalhos_folder = course_folder / "Trabalhos"
        materiais_folder = course_folder / "Materiais"
        anuncios_folder.mkdir(exist_ok=True)
        trabalhos_folder.mkdir(exist_ok=True)
        materiais_folder.mkdir(exist_ok=True)
        
        # ===== BUSCA ANÚNCIOS =====
        print(f"  📢 Buscando anúncios...")
        page_token = None
        
        while True:
            announcements = service_classroom.courses().announcements().list(
                courseId=course_id,
                pageToken=page_token,
                pageSize=100
            ).execute()
            
            for announcement in announcements.get('announcements', []):
                materials = announcement.get('materials', [])
                announcement_text = announcement.get('text', '')
                announcement_name = get_publication_name(
                    title='',
                    text=announcement_text,
                    date_str=announcement.get('updateTime') or announcement.get('creationTime'),
                    default_prefix='Anuncio'
                )

                register_result(save_publication_text(anuncios_folder, announcement_name, announcement_text))

                announcement_prefix = sanitize_filename(announcement_name[:30])
                
                for i, material in enumerate(materials):
                    # Drive file
                    if 'driveFile' in material:
                        drive_file = material['driveFile']['driveFile']
                        file_id = drive_file['id']
                        file_name = drive_file.get('title', f'anuncio_{i}')
                        
                        # Adiciona prefixo do anúncio para evitar conflitos
                        file_name = f"[{announcement_prefix}]_{file_name}"
                        
                        register_result(download_file_from_drive(service_drive, file_id, file_name, anuncios_folder))
                    
                    # Link externo
                    elif 'link' in material:
                        link = material['link']
                        url = link['url']
                        title = link.get('title', 'Link')
                        
                        file_name = f"[{announcement_prefix}]_{title}"
                        
                        register_result(download_link_content(url, file_name, anuncios_folder))
                    
                    # YouTube video
                    elif 'youtubeVideo' in material:
                        video = material['youtubeVideo']
                        video_id = video['id']
                        title = video.get('title', 'Video')
                        url = f"https://youtube.com/watch?v={video_id}"
                        
                        file_name = f"[{announcement_prefix}]_{title}"
                        
                        register_result(download_link_content(url, file_name, anuncios_folder))
            
            page_token = announcements.get('nextPageToken')
            if not page_token:
                break
        
        # ===== BUSCA TRABALHOS/COURSEWORK =====
        print(f"  📝 Buscando trabalhos...")
        page_token = None
        
        while True:
            coursework = service_classroom.courses().courseWork().list(
                courseId=course_id,
                pageToken=page_token,
                pageSize=100
            ).execute()
            
            for work in coursework.get('courseWork', []):
                materials = work.get('materials', [])
                work_description = work.get('description', '')
                work_name = get_publication_name(
                    title=work.get('title', ''),
                    text=work_description,
                    date_str=work.get('updateTime') or work.get('creationTime'),
                    default_prefix='Trabalho'
                )
                
                work_folder_name = sanitize_filename(work_name[:50])
                work_folder = trabalhos_folder / work_folder_name
                work_folder.mkdir(exist_ok=True)

                register_result(save_publication_text(work_folder, work_name, work_description))
                
                for i, material in enumerate(materials):
                    # Drive file
                    if 'driveFile' in material:
                        drive_file = material['driveFile']['driveFile']
                        file_id = drive_file['id']
                        file_name = drive_file.get('title', f'trabalho_{i}')
                        
                        register_result(download_file_from_drive(service_drive, file_id, file_name, work_folder))
                    
                    # Link externo
                    elif 'link' in material:
                        link = material['link']
                        url = link['url']
                        title = link.get('title', 'Link')
                        
                        register_result(download_link_content(url, title, work_folder))
                    
                    # YouTube video
                    elif 'youtubeVideo' in material:
                        video = material['youtubeVideo']
                        video_id = video['id']
                        title = video.get('title', 'Video')
                        url = f"https://youtube.com/watch?v={video_id}"
                        
                        register_result(download_link_content(url, title, work_folder))
                    
                    # Formulário
                    elif 'form' in material:
                        form = material['form']
                        form_url = form.get('formUrl', '')
                        title = form.get('title', 'Formulario')
                        
                        register_result(download_link_content(form_url, title, work_folder))
            
            page_token = coursework.get('nextPageToken')
            if not page_token:
                break
        
        # ===== BUSCA MATERIAIS DO CURSO =====
        print(f"  📚 Buscando materiais do curso...")
        page_token = None
        
        while True:
            course_work_materials = service_classroom.courses().courseWorkMaterials().list(
                courseId=course_id,
                pageToken=page_token,
                pageSize=100
            ).execute()
            
            for material_item in course_work_materials.get('courseWorkMaterial', []):
                materials = material_item.get('materials', [])
                material_description = material_item.get('description', '')
                material_name = get_publication_name(
                    title=material_item.get('title', ''),
                    text=material_description,
                    date_str=material_item.get('updateTime') or material_item.get('creationTime'),
                    default_prefix='Material'
                )

                material_folder_name = sanitize_filename(material_name[:50])
                material_folder = materiais_folder / material_folder_name
                material_folder.mkdir(parents=True, exist_ok=True)

                register_result(save_publication_text(material_folder, material_name, material_description))
                
                for i, material in enumerate(materials):
                    if 'driveFile' in material:
                        drive_file = material['driveFile']['driveFile']
                        file_id = drive_file['id']
                        file_name = drive_file.get('title', f'material_{i}')
                        
                        register_result(download_file_from_drive(service_drive, file_id, file_name, material_folder))
                    
                    elif 'link' in material:
                        link = material['link']
                        url = link['url']
                        title = link.get('title', 'Link')
                        
                        register_result(download_link_content(url, title, material_folder))
            
            page_token = course_work_materials.get('nextPageToken')
            if not page_token:
                break
        
    except HttpError as e:
        print(f"  ⚠️ Erro ao processar materiais: {str(e)}")
    
    return downloaded_count, skipped_count


def main():
    """Executa o fluxo completo de configuração, autenticação e download.

    Inicializa preferências locais, percorre disciplinas do utilizador,
    descarrega conteúdo incrementalmente e grava relatório final em JSON.
    """
    print("=" * 60)
    print("📚 Google Classroom Downloader")
    print("=" * 60)
    print()

    # Inicializa configuração (inclui setup opcional de autoarranque)
    initialize_config()
    
    # Obtém credenciais
    print("🔐 Autenticando...")
    creds = get_credentials()
    if not creds:
        return
    
    # Cria serviços
    service_classroom = build('classroom', 'v1', credentials=creds)
    service_drive = build('drive', 'v3', credentials=creds)
    
    # Cria pasta de downloads
    DOWNLOAD_DIR.mkdir(exist_ok=True)
    print(f"📁 Pasta de downloads: {DOWNLOAD_DIR}")
    print()
    
    # Lista cursos do usuário
    print("📋 Buscando suas disciplinas...")
    print()
    
    try:
        courses = []
        page_token = None
        
        while True:
            response = service_classroom.courses().list(
                studentId='me',
                pageToken=page_token,
                pageSize=100
            ).execute()
            
            courses.extend(response.get('courses', []))
            page_token = response.get('nextPageToken')
            
            if not page_token:
                break
        
        if not courses:
            print("❌ Nenhuma disciplina encontrada!")
            return
        
        print(f"✅ Encontradas {len(courses)} disciplinas")
        print()
        
        # Estatísticas
        total_downloaded = 0
        total_skipped = 0
        course_stats = []
        
        # Processa cada curso
        for course in courses:
            course_id = course['id']
            course_name = course.get('name', 'Sem nome')
            course_section = course.get('section', '')
            
            # Cria nome da pasta
            folder_name = sanitize_filename(course_name)
            if course_section:
                folder_name += f"_{sanitize_filename(course_section)}"
            
            course_folder = DOWNLOAD_DIR / folder_name
            course_folder.mkdir(exist_ok=True)
            
            print(f"📖 {course_name}")
            if course_section:
                print(f"   Seção: {course_section}")
            
            # Processa materiais
            count_downloaded, count_skipped = process_materials(
                service_classroom,
                service_drive,
                course_id,
                course_name,
                course_folder
            )
            
            course_stats.append({
                'name': course_name,
                'files': count_downloaded,
                'skipped': count_skipped
            })
            total_downloaded += count_downloaded
            total_skipped += count_skipped
            
            print(f"   📥 {count_downloaded} arquivo(s) baixado(s)")
            print(f"   ⏭️ {count_skipped} arquivo(s) já existente(s)")
            print()
        
        # Resumo final
        print("=" * 60)
        print("📊 RESUMO")
        print("=" * 60)
        print(f"\n📁 Pasta de downloads: {DOWNLOAD_DIR}")
        print(f"📚 Disciplinas processadas: {len(courses)}")
        print(f"📥 Total de arquivos baixados: {total_downloaded}")
        print(f"⏭️ Total de arquivos já existentes: {total_skipped}")
        print()
        
        for stat in course_stats:
            print(f"  • {stat['name']}: {stat['files']} baixado(s), {stat['skipped']} já existente(s)")
        
        print()
        print("✅ Download concluído!")
        
        # Salva relatório
        report_timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        report_file = DOWNLOAD_DIR / f"_download_report_{report_timestamp}.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump({
                'total_courses': len(courses),
                'total_files': total_downloaded,
                'total_skipped': total_skipped,
                'courses': course_stats
            }, f, indent=2, ensure_ascii=False)
        
        print(f"📄 Relatório salvo em: {report_file}")
        
    except HttpError as e:
        print(f"❌ Erro ao acessar o Classroom: {str(e)}")
    except Exception as e:
        print(f"❌ Erro inesperado: {str(e)}")


if __name__ == '__main__':
    main()
