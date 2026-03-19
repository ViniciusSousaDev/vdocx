# Leitor de Cartão CNPJ — Guia de Instalação e Empacotamento

## Pré-requisitos

- Python 3.9 ou superior → https://www.python.org/downloads/
- pip (já vem com o Python)

---

## 1. Instalação das dependências (desenvolvimento)

```bash
pip install -r requirements.txt
```

Rodar o app em modo desenvolvimento:

```bash
python main.py
```

---

## 2. Gerar executável instalável (.exe) com PyInstaller

### 2.1 Instalar o PyInstaller

```bash
pip install pyinstaller
```

### 2.2 Gerar o .exe (arquivo único)

```bash
pyinstaller --onefile --windowed --name "LeitorCNPJ" main.py
```

| Flag | Significado |
|------|-------------|
| `--onefile` | Empacota tudo em um único .exe |
| `--windowed` | Não abre o terminal (modo GUI) |
| `--name "LeitorCNPJ"` | Nome do executável gerado |

O executável gerado estará em:

```
dist/LeitorCNPJ.exe
```

### 2.3 (Opcional) Adicionar ícone personalizado

Coloque um arquivo `icone.ico` na pasta do projeto e adicione ao comando:

```bash
pyinstaller --onefile --windowed --name "LeitorCNPJ" --icon=icone.ico main.py
```

---

## 3. Distribuir para outras máquinas

Basta copiar o arquivo `dist/LeitorCNPJ.exe` para qualquer máquina Windows.
**Não é necessário instalar Python** nas máquinas de destino.

---

## 4. (Avançado) Criar instalador com Inno Setup

Para um instalador profissional com atalho na área de trabalho:

1. Instale o **Inno Setup**: https://jrsoftware.org/isinfo.php
2. Crie um script `.iss` apontando para `dist/LeitorCNPJ.exe`
3. Compile → gera um `setup.exe` para distribuição

---

## Estrutura do projeto

```
leitor-cnpj/
├── main.py            ← código principal
├── requirements.txt   ← dependências Python
├── README.md          ← este arquivo
├── icone.ico          ← (opcional) ícone do app
└── dist/
    └── LeitorCNPJ.exe ← executável gerado pelo PyInstaller
```
