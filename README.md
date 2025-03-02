# Email Reminder Birthday

Este projeto automatiza o envio de lembretes de aniversário via **Google Apps Script**. Ele envia e-mails diários para aniversariantes do dia e um resumo semanal dos aniversariantes do mês.

## 📌 Funcionalidades
- Obtém aniversariantes do mês e envia e-mail semanal para um administrador.
- Obtém aniversariantes do dia e envia e-mails automáticos de felicitações.
- Permite configurar **triggers automáticos** para envio diário e semanal.

## 🛠️ Tecnologias Utilizadas
- **Google Apps Script**
- **Google Sheets API**
- **Gmail API**

## 🚀 Como Usar
1. **Configure a planilha do Google Sheets** com as abas:
   - `Pessoal.Email` → Lista de e-mails que devem receber notificações.
   - `Pessoal.Saude` → Lista de aniversariantes (coluna G: Data de nascimento, coluna J: E-mail).
2. **Configure o ID do Administrador** no arquivo `src/email_reminder.js`:
   ```js
   const EMAIL_ADMIN = "seuemail@exemplo.com"; 
