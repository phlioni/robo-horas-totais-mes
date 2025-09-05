# Robô de Resumo Mensal de Apontamento de Horas

Este projeto consiste em um robô (RPA) desenvolvido em Python para automatizar a geração de um relatório mensal sobre o status de apontamento de horas dos profissionais em um portal web específico.

## Descrição da Funcionalidade

O robô foi projetado para ser executado diariamente (sugestão: 12:00) através do Agendador de Tarefas do Windows. Sua função é consolidar e reportar o progresso dos apontamentos de horas de cada profissional dentro do mês corrente.

O fluxo de trabalho é o seguinte:

1.  **Cálculo de Horas Úteis:** Primeiramente, o robô calcula o total de horas de trabalho esperadas para o mês corrente. Ele faz isso identificando todos os dias úteis (segunda a sexta) e subtraindo os feriados nacionais (Brasil), estaduais (São Paulo) e municipais (Santos). A base de cálculo padrão é de 8 horas por dia útil.
2.  **Login e Download:** O robô acessa o portal web, realiza o login e aplica o filtro "Mês Corrente" para baixar um relatório em Excel com todos os apontamentos do mês até a data da execução.
3.  **Processamento dos Dados:** A planilha baixada é processada para agregar o total de horas que cada profissional já lançou no mês.
4.  **Geração do Relatório:** Com os dados processados e o total de horas úteis calculado, o robô gera um resumo contendo:
    * **Horas Lançadas:** Total de horas que o profissional registrou no mês.
    * **Horas Necessárias:** Total de horas que o profissional deve registrar no mês (dias úteis x 8).
    * **Saldo de Horas:** A diferença entre as horas necessárias e as lançadas, indicando o quanto ainda falta apontar.
5.  **Envio de E-mail:** Um e-mail é formatado e enviado para os destinatários configurados. O e-mail contém uma tabela com o resumo para cada profissional e uma tabela final com os totais gerais.
6.  **E-mail de Status:** Ao final da execução, um segundo e-mail é enviado (para um destinatário técnico) informando se o robô rodou com `SUCESSO` ou `FALHA`, incluindo detalhes do erro, se houver.

## Configuração

Todas as configurações sensíveis e personalizáveis estão centralizadas no arquivo `config.py`. Antes de executar, você deve preencher:

* **Credenciais do Site:** `SITE_LOGIN` e `SITE_SENHA`.
* **Credenciais de E-mail:** `EMAIL_REMETENTE` e `EMAIL_SENHA`.
* **Destinatários do Relatório Mensal:** `MENSAL_DESTINATARIO_PRINCIPAL` e `MENSAL_DESTINATARIOS_COPIA`.
* **Destinatário de Status:** `STATUS_EMAIL_DESTINATARIO` para receber o e-mail sobre o sucesso ou falha da execução.

## Como Usar

### Pré-requisitos
* Python 3.x
* Bibliotecas listadas no arquivo `requirements.txt`

### Instalação

1.  Clone ou baixe este repositório.
2.  Abra um terminal na pasta raiz do projeto.
3.  Crie e ative um ambiente virtual (recomendado):
    ```bash
    python -m venv venv
    venv\Scripts\activate
    ```
4.  Instale as dependências:
    ```bash
    pip install -r requirements.txt
    ```

### Execução

1.  Certifique-se de que o arquivo `config.py` está preenchido corretamente.
2.  Execute o script principal a partir do terminal para um teste manual:
    ```bash
    python main.py
    ```
3.  Para automatizar a execução diária, configure uma tarefa no **Agendador de Tarefas do Windows** para executar o `main.py` no horário desejado (ex: 12:00).