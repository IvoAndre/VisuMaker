# Quick Start - Sistema SMTP

Começar a usar o novo sistema SMTP é muito simples!

## 1. Setup Inicial (3 minutos)

```powershell
# Copia arquivo de configuração
copy config.py.default config.py

# Edita config.py com suas credenciais SMTP
# (veja exemplos abaixo)

# Instala dependências
pip install -r requirements.txt
```

## 2. Exemplos de Configuração SMTP

### Gmail (Recomendado)

1. Ative 2FA: https://myaccount.google.com/security
2. Crie App Password: https://myaccount.google.com/apppasswords
3. Configure:

```python
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SMTP_USER = 'seu.email@gmail.com'
SMTP_PASSWORD = 'senha_app_gerada'  # 16 caracteres, sem espaços
SMTP_FROM_NAME = 'Meu Nome'
SMTP_USE_TLS = True
```

### Outlook.com

```python
SMTP_SERVER = 'smtp-mail.outlook.com'
SMTP_PORT = 587
SMTP_USER = 'seu.email@outlook.com'
SMTP_PASSWORD = 'sua_senha'
SMTP_FROM_NAME = 'Meu Nome'
SMTP_USE_TLS = True
```

### Servidor SMTP Personalizado

Contacte seu provedor para obter:
- Endereço do servidor SMTP
- Porta (normalmente 587 ou 465)
- Usa TLS? (normalmente sim)
- Username e password

## 3. Usar a Aplicação

```powershell
# Windows
run.bat

# Ou direto
python app.py

# Com modo verbose (debug)
python app.py -v
```

## 4. Fluxo Básico

1. **Carregar Modelo**: Carregue uma imagem PNG/JPG como base
2. **Adicionar Elementos**: Texto e imagens onde desejar
3. **Carregar CSV**: CSV com dados dos participantes (ex: nome, email)
4. **Configurar Email**: Customize assunto e corpo (opcional)
5. **Gerar & Enviar**: Clique no botão para processar

## 5. Estrutura do CSV

Exemplo mínimo:
```csv
nome,email
João Silva,joao@example.com
Maria Santos,maria@example.com
```

Exemplo completo:
```csv
nome,email,titulo,empresa,data
João Silva,joao@example.com,Especialista,ACME Corp,2026-01-26
Maria Santos,maria@example.com,Gerente,Tech Solutions,2026-01-26
```

## 6. Placeholders Disponíveis

Use em Texto e Email:
- `{nome}` - Nome do participante
- `{email}` - Email do participante
- `{data_atual}` - Data de hoje (DD/MM/YYYY)
- Qualquer coluna do CSV

Exemplo texto:
```
Certificado de Conclusão

Certificamos que ___________________
                 {nome}

Completou o curso com sucesso.

Data: {data_atual}
```

## 7. Troubleshooting

### "Erro de autenticação"
- Verifique email e senha em config.py
- Para Gmail, use App Password, não a senha regular
- Confirme SMTP_SERVER está correto

### "Timeout"
- Verifique firewall permite porta 587/465
- Tente outro servidor SMTP
- Verifique conexão de internet

### "SSL: CERTIFICATE_VERIFY_FAILED"
Tente uma destas:
```python
# Opção 1: Usar TLS na porta 587
SMTP_PORT = 587
SMTP_USE_TLS = True

# Opção 2: Usar SSL na porta 465
SMTP_PORT = 465
SMTP_USE_TLS = False
```

## 8. Segurança

✅ **BOM**: Credenciais em `config.py` (local, não no git)  
❌ **EVITAR**: Colocar `config.py` no git (use `.gitignore`)  
❌ **NUNCA**: Usar senhas de conta pessoal (use App Password para Gmail)  

## 9. Documentação Completa

- `README.md` - Descrição geral
- `CHANGELOG.md` - Mudanças da versão
- `MIGRACAO.md` - Se está a migrar do O365
- `NOVO_REPO.md` - Para criar novo repositório Git

---

**Pronto!** Tudo configurado. Boa sorte com seus documentos! 🎉

Dúvidas? Consulte a documentação completa ou os exemplos de configuração.
