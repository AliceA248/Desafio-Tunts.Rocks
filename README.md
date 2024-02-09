# Desafio-Tunts.Rocks

Controle Acadêmico - Node.js Excel Calculator
Este projeto Node.js utiliza a biblioteca xlsx para realizar o controle acadêmico, calculando a situação de alunos com base em médias e faltas. Ele também determina a nota necessária para aprovação final em caso de exame.

## Como Usar

### Instalação das Dependências:

Execute o seguinte comando para instalar as dependências necessárias:

### `npm install`

### Execução da Aplicação:

Execute o seguinte comando para calcular as situações e notas para aprovação final e gerar uma nova planilha:


### `node index.js`

Certifique-se de que o arquivo Excel original (planilha.xlsx) está presente na pasta planilha.

## Resultados:

Os resultados serão salvos no arquivo nova_planilha.xlsx na pasta planilha.

## Estrutura do Projeto

- index.js: Contém o código principal para ler, processar e escrever na planilha.

- planilha/planilha.xlsx: Arquivo Excel original com os dados dos alunos.

- planilha/nova_planilha.xlsx: Arquivo Excel gerado com as situações e notas para aprovação final.

### Regras de Avaliação
Média < 5: Reprovado por Nota.
5 <= Média < 7: Exame Final.
Média >= 7: Aprovado.
Faltas > 25%: Reprovado por Falta.

### Observações

- As notas da P1, P2 e P3 são consideradas nas colunas D, E e F, respectivamente.

- O resultado é salvo nas colunas G (Situação) e H (Nota para Aprovação Final).

