
<img width="1665" height="512" alt="VisuMakerLogo" src="https://github.com/user-attachments/assets/0c49a6bf-66b6-4161-bf79-37cae05cf934" />

# VisuMaker

Aplicação para geração automática de documentos visuais a partir de modelos personalizáveis.

## Descrição

O VisuMaker é uma ferramenta desenvolvida em Python com interface gráfica que permite criar e personalizar documentos visuais em massa.

<img width="1202" height="732" alt="image" src="https://github.com/user-attachments/assets/629dfbf6-acf4-4893-ad4a-83f1279c9a57" />

### Funcionalidades Principais

- Interface gráfica intuitiva para design de documentos visuais
- Importação de dados a partir de ficheiros CSV
- Posicionamento preciso de texto e imagens
- Suporte para variáveis/placeholders que são substituídos pelos dados
- Envio automático de documentos por email via SMTP
- Pré-visualização do resultado final
- Salvar e carregar layouts para reutilização

## Requisitos

- [Python 3.8 ou superior](https://www.python.org/downloads/) 
- Bibliotecas Python (ver `requirements.txt`)
- Sistema operativo Windows (algumas funcionalidades são específicas para Windows)
- Conta de email com suporte SMTP (Gmail, Outlook, etc.)

## Instalação

1. Clone este repositório:
   ```
   git clone https://github.com/IvoAndre/visumaker.git
   cd visumaker
   ```

2. Execute o script de instalação:
   ```
   install.bat
   ```
   
   Ou instale manualmente as dependências:
   ```
   py -m pip install -r requirements.txt
   ```

3. Copie o ficheiro de configuração e edite conforme necessário:
   ```
   copy config.py.default config.py
   ```

## Configuração

Edite o ficheiro `config.py` para configurar:

### Configuração de Email SMTP

```python
SMTP_SERVER = 'smtp.gmail.com'  # ou seu servidor SMTP
SMTP_PORT = 587                 # 587 para TLS, 465 para SSL
SMTP_USER = 'seu_email@gmail.com'
SMTP_PASSWORD = 'sua_senha_ou_token'
SMTP_FROM_NAME = 'VisuMaker'
SMTP_USE_TLS = True             # True para TLS, False para SSL
```

#### Exemplos de Servidores SMTP Comuns:

**Gmail:**
```
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SMTP_USE_TLS = True
```

**Outlook/Hotmail:**
```
SMTP_SERVER = 'smtp-mail.outlook.com'
SMTP_PORT = 587
SMTP_USE_TLS = True
```

**SMTP Personalizado:**
Consulte o seu fornecedor de email para os detalhes do servidor SMTP.

### Outras Configurações

- Configurações padrão de email (assunto, corpo, etc.)
- Fontes e estilos padrão
- Placeholders globais

## Utilização

1. Execute a aplicação:
   ```
   run.bat
   ```
   Ou:
   ```
   python app.py
   ```

2. Carregue uma imagem de modelo para o fundo do documento
3. Adicione texto e imagens conforme necessário
4. Importe dados dos participantes através de um ficheiro CSV
5. Gere os certificados e, opcionalmente, envie-os por email
