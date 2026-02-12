# Pipeline de Auditoria e Saneamento de Dados (Python Script)

Este projeto foi desenvolvido com o objetivo de automatizar a conferência, tabulação e validação de grandes bases de dados, substituindo auditorias manuais por um processo programático, rápido e auditável.

Originalmente concebido para otimizar fluxos na **Perpart S/A**, o script identifica inconsistências em bases cadastrais de RH e Engenharia, garantindo a integridade da informação antes de tomadas de decisão estratégicas.

## Problema vs. Solução

* **Problema:** Bases de dados legadas com CPFs inválidos, erros de preenchimento, datas incoerentes e salários fora de conformidade, exigindo horas de revisão manual.
* **Solução:** Um pipeline em Python que aplica 8 níveis de validação técnica e lógica, gerando um **Relatório de Inconsistências** detalhado para a equipe de Compliance.

## Tecnologias Utilizadas

* **Python 3.x**
* **Pandas:** Manipulação e análise de estruturas de dados de alta performance.
* **Regex (Expressões Regulares):** Validação de padrões complexos (E-mail, Telefone, CEP).
* **Openpyxl:** Integração direta com arquivos Excel (.xlsx).

## Validações Executadas

O script percorre a base de dados e valida os seguintes critérios:

1.  **Integridade Matemática:** Validação de dígito verificador de **CPF**.
2.  **Conformidade de Contato:** Verificação de sintaxe de **E-mail** e **Telefone**.
3.  **Lógica Temporal:** Identificação de datas de admissão inválidas ou no futuro.
4.  **Saneamento Financeiro:** Detecção de valores não numéricos ou negativos em campos de salário.
5.  **Validação Geográfica:** Checagem de siglas de **UF** (Estados brasileiros) e formato de **CEP**.
6.  **Campos Obrigatórios:** Identificação de células vazias em colunas críticas (Nome, ID).

## Impacto Gerado

Ao utilizar esta ferramenta em cenários reais, os ganhos observados foram:

* **Redução de Tempo:** Processamento de bases com milhares de linhas em poucos segundos (antigamente levava-se horas).
* **Confiabilidade:** Eliminação do erro humano na fase de triagem de dados.
* **Auditabilidade:** Geração de um log de erros que serve como evidência de conformidade para processos de auditoria.

## Como Executar

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/kaykyDias04/nome-do-repositorio.git](https://github.com/kaykyDias04/nome-do-repositorio.git)
    ```
2.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```
3.  **Execute o script:**
    Coloque seu arquivo `.xlsx` na pasta raiz e execute:
    ```python
    python validador_excel.py
    ```

