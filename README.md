# Gerenciador de Chamados de TI

Este repositório contém um script do Google Apps Script que automatiza o processo de registro e notificação de chamados de TI a partir de um formulário do Google. O script é projetado para facilitar a gestão de solicitações de suporte técnico, registrando informações em uma planilha do Google e enviando notificações por e-mail.

## Funcionalidades

- **Captura de Dados do Formulário:** Coleta informações do formulário, incluindo carimbo de data/hora, e-mail do solicitante, nome, local, tipo de chamado, categoria, urgência, título, descrição e anexos.
- **Registro na Planilha:** Registra os dados coletados em uma planilha do Google, atribuindo um ID sequencial e definindo o status inicial como "Aberto".
- **Processamento de Anexos:** Converte documentos do Google (Docs, Sheets, Slides) em PDFs para anexar aos e-mails.
- **Notificação por E-mail:** Envia e-mails de notificação ao responsável técnico e ao solicitante, informando sobre o registro do chamado.
- **Formatação da Planilha:** Aplica formatação condicional à planilha, alterando cores de fundo com base na urgência e no status do chamado, além de centralizar o texto em colunas específicas.
- **Tratamento de Erros:** Registra erros em uma aba específica da planilha ("Log de Erros") e notifica o administrador por e-mail em caso de falhas.

## Como Usar

1. **Configuração do Formulário:**
   - Crie um formulário do Google com os campos necessários (carimbo de data/hora, e-mail, nome, local, tipo, categoria, urgência, título, descrição e anexos).

2. **Configuração da Planilha:**
   - Crie uma planilha do Google onde os dados do formulário serão registrados. Certifique-se de que a estrutura da planilha corresponda aos campos do formulário.

3. **Implementação do Script:**
   - Acesse o Google Apps Script a partir da planilha do Google.
   - Cole o código do script no editor.
   - Configure os endereços de e-mail do responsável técnico e do administrador no script.

4. **Ativação do Trigger:**
   - Configure um trigger para que a função `onFormSubmit` seja acionada sempre que o formulário for enviado.

5. **Testes:**
   - Realize testes enviando o formulário e verifique se os dados são registrados corretamente na planilha e se as notificações por e-mail são enviadas.

## Estrutura do Código

- `onFormSubmit(e)`: Função principal que é acionada ao enviar o formulário.
- `gerarEmailHTML(titulo, corpoHTML, corHeader)`: Gera um template HTML para os e-mails.
- `processarAnexos(linksAnexo)`: Processa e converte anexos do Google Drive.
- `aplicarFormatacaoUrgencia(sheet)`: Aplica formatação condicional à coluna de urgência.
- `aplicarFormatacaoStatus(sheet)`: Aplica formatação condicional à coluna de status.
- `centralizarColunas(sheet)`: Centraliza o texto em colunas específicas.
- `registrarErro(erro, contexto, dadosExtras)`: Registra erros e notifica o administrador.

## Contribuições

Contribuições são bem-vindas! Se você tiver sugestões ou melhorias, sinta-se à vontade para abrir uma issue ou enviar um pull request.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

## Contato

Para mais informações, entre em contato com [seu nome ou e-mail].
