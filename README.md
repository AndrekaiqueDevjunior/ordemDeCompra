
---

```markdown
# 📄 Ordem de Compra Automatizada com Google Apps Script

Projeto desenvolvido em **Google Apps Script** com integração ao **Google Sheets** para geração automática de ordens de compra (OC), envio de PDFs e controle de solicitações de forma inteligente.

---

## 🛠️ Funcionalidades

- 🔄 Preenchimento automático da aba **ORDEM DE COMPRA** com base no número do pedido.
- 📦 Geração de número de OC automática e registro em **Histórico** e **Controle Pedidos**.
- 🧾 Geração de **PDF** da ordem de compra com envio automático por e-mail.
- 🧠 Validação e busca de dados de **CNPJ** via API com cache local.
- 📬 Envio de e-mails direcionado por **empresa selecionada**.
- 🛡️ Campos obrigatórios e validações de segurança.
- 🗂️ Armazenamento de dados de empresas no banco local (`Banco de Dados - Empresas`).

---

## 📁 Estrutura do Projeto

```

ordem-de-compra-gas/
├── src/
│   ├── main.gs                # Código principal do Apps Script
│   ├── gerarPDF.gs            # Função para gerar e enviar PDF
│   ├── buscarCNPJ.gs          # Busca de dados via CNPJ
│   ├── utilitarios.gs         # Funções auxiliares e validações
├── .clasp.json                # Configuração do Clasp
├── .gitignore                 # Ignora arquivos sensíveis
├── README.md                  # Você está aqui :)

````

---

## 🚀 Como executar localmente

1. Instale o [Node.js](https://nodejs.org/) e o [Clasp](https://github.com/google/clasp):
   ```bash
   npm install -g @google/clasp
````

2. Faça login com sua conta Google:

   ```bash
   clasp login
   ```

3. Clone este projeto:

   ```bash
   clasp clone <SCRIPT_ID>
   ```

4. Edite os arquivos em `src/`, depois envie ao GAS:

   ```bash
   clasp push
   ```

---

## 🔐 Permissões

Este projeto utiliza permissões para:

* Leitura/escrita no Google Drive
* Envio de e-mails
* Acesso a planilhas
* Requisições externas via URL Fetch (para a API de CNPJ)

---

## 👨‍💻 Autor

**André**
Projeto desenvolvido com foco em automação de processos administrativos internos.
🔗 \[LinkedIn opcional ou contato]

---

## 📌 Observações

> Este projeto foi desenvolvido com fins internos e está em constante evolução. Contribuições são bem-vindas!

---

## 📃 Licença

Este repositório está licenciado sob a [MIT License](LICENSE).

```

---

```
