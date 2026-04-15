# Classroom Downloader

Automação em Python para descarregar e organizar localmente conteúdos do Google Classroom.

Este projeto foi desenvolvido para simplificar o backup de materiais académicos, mantendo uma estrutura previsível por disciplina e categoria de publicação, com foco em robustez para Windows e experiência de uso prática.

## Título e Descrição

O Classroom Downloader é um script de automação Python que se liga às APIs do Google Classroom e Google Drive para:

- recolher anúncios, trabalhos e materiais do curso;
- descarregar anexos e referências (incluindo links e vídeos como atalhos);
- guardar o texto original das publicações em ficheiros .txt;
- organizar tudo em diretórios locais de forma automática.

## Funcionalidades Principais (Features)

- Download incremental: evita downloads e gravações redundantes, ignorando ficheiros que já existem localmente.
- Arranque automático com o Windows: primeira execução com prompt S/N e persistência da escolha em config.json.
- Histórico de relatórios dinâmicos: gera um relatório por execução com o formato _download_report_YYYY-MM-DD_HH-MM-SS.json.
- Conversão de ficheiros Google Workspace para Microsoft Office:
  - Google Docs -> .docx
  - Google Sheets -> .xlsx
  - Google Slides -> .pptx
  - Google Drawings -> .png
- Organização automática por disciplina e por categoria (Anuncios, Trabalhos, Materiais).
- Exportação de links externos e vídeos YouTube como ficheiros .url para referência local.
- Sanitização de nomes para compatibilidade com Windows (incluindo limite de tamanho e limpeza de caracteres inválidos).

## Para Utilizadores Finais (Como Usar o .exe)

### Pré-requisitos

- O executável do projeto.
- Um ficheiro credentials.json válido da Google Cloud API na mesma pasta do .exe.

### Passos

1. Copie classroom_downloader.exe para uma pasta local à sua escolha.
2. Coloque credentials.json exatamente na mesma pasta do executável.
3. Execute o .exe.
4. Na primeira execução:
   - responda ao prompt de arranque automático no Windows (S/N);
   - autorize o acesso da sua conta Google no navegador.
5. Aguarde o fim da execução e consulte a pasta Classroom_Downloads na sua home.

### Resultado esperado

- Estrutura de ficheiros organizada por disciplina/categoria.
- Relatório de execução com nome dinâmico:
  - _download_report_2026-04-15_16-30-00.json

## Para Developers (Como Clonar e Configurar)

### 1) Clonar o repositório

```bash
git clone <URL_DO_REPOSITORIO>
cd classroom_downloader
```

### 2) Criar ambiente virtual (recomendado)

No Windows (PowerShell):

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

No macOS/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3) Instalar dependências

```bash
pip install -r requirements.txt
```

### 4) Executar em modo script

```bash
python classroom_downloader.py
```

## Configuração da API Google

Tutorial rápido para gerar credentials.json:

1. Aceda ao Google Cloud Console: https://console.cloud.google.com/
2. Crie um projeto novo (ou use um existente).
3. Ative as APIs:
   - Google Classroom API
   - Google Drive API
4. Em APIs e Serviços > Credenciais, clique em Criar Credenciais.
5. Escolha ID do Cliente OAuth.
6. Tipo de aplicação: Aplicação para computador (Desktop App).
7. Faça download do JSON e renomeie para credentials.json.
8. Coloque credentials.json na raiz da execução (ao lado do .py ou .exe).

## Compilação

O projeto inclui o script [build.bat](build.bat), que empacota a aplicação num executável único com PyInstaller.

Conteúdo do build.bat:

```bat
@echo off
echo A instalar PyInstaller...
pip install pyinstaller
pyinstaller --onefile classroom_downloader.py
```

Para compilar:

```bat
build.bat
```

Após a compilação, o executável será gerado na pasta dist.

## Estrutura de Saída (Exemplo)

```text
Classroom_Downloads/
├── Disciplina_A/
│   ├── Anuncios/
│   │   ├── [Publicacao]_Anexo.docx
│   │   └── Publicacao - Texto.txt
│   ├── Trabalhos/
│   │   └── Trabalho_1/
│   │       ├── Enunciado.pdf
│   │       └── Trabalho_1 - Texto.txt
│   └── Materiais/
│       └── Tema_1/
│           ├── Slides.pptx
│           └── Tema_1 - Texto.txt
└── _download_report_YYYY-MM-DD_HH-MM-SS.json
```

## Boas Práticas e Notas

- Mantenha credentials.json e token.pickle fora de partilhas públicas.
- Se alterar scopes da API, remova token.pickle para forçar nova autenticação.
- Para execuções recorrentes, o modo incremental reduz tempo e consumo de quota.

## Licença

Uso educacional e pessoal. Ajuste a política de licenciamento conforme o seu contexto de publicação no GitHub.
