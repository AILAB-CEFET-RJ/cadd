# CADD

## Instalação

O projeto utiliza **Python 3** (recomendado Python 3.9 ou superior).
Baixe e instale em [https://www.python.org/downloads/](https://www.python.org/downloads/).

Clone este repositório e, dentro da pasta do projeto, instale as dependências usando o `requirements.txt`:

```bash
pip install -r requirements.txt
```

O arquivo `requirements.txt` contém:

```
plotly
openpyxl
XlsxWriter
```

---

## Geração dos relatórios a partir de arquivos CSV

1. **Exportar do SIE** (em formato CSV) o relatório `11.02.05.99.60`.

   * Deve ser gerado **um arquivo para cada curso de graduação**.
   * O nome de cada arquivo deve ser a sigla do curso correspondente (e.g., `GADM.csv`, `BCC.csv` etc).

2. Criar a pasta **`data`** no mesmo nível da pasta **`src`**.

   * Copiar todos os arquivos CSV exportados do SIE para dentro da pasta `data`.

3. Criar a pasta **`resultados`** no mesmo nível da pasta **`src`**.

4. No terminal, acesse a pasta **`src`**:

   ```bash
   cd src
   ```

5. Editar o arquivo **`main.py`** para definir o período letivo desejado.

6. Executar:

   ```bash
   python main.py
   ```

Os relatórios em **Excel** e **HTML** serão gerados dentro da pasta `resultados`.

---

## Execução dos scripts específicos (CSV)

Você também pode rodar cada script separadamente, informando período (`-p`), pasta de entrada (`-i`) e pasta de saída (`-o`).

### Exemplo: relatórios do período **2016.1**

```bash
python main_faixas_criticidade.py -p 2016.1 -i ../data -o ../resultados
python main_integralizacoes.py   -p 2016.1 -i ../data -o ../resultados
python main_reprovacoes.py       -p 2016.1 -i ../data -o ../resultados
```

### Exemplo: relatórios do período **2016.2**

```bash
python main_faixas_criticidade.py -p 2016.2 -i ../data -o ../resultados
python main_integralizacoes.py   -p 2016.2 -i ../data -o ../resultados
python main_reprovacoes.py       -p 2016.2 -i ../data -o ../resultados
```

### Script unificado

Também é possível rodar todos os relatórios de uma vez com:

```bash
python main.py
```

> **Nota**: edite previamente o arquivo `main.py` para atualizar o período letivo.

---

## Processamento direto de arquivos Excel

Além dos relatórios em CSV, há scripts que processam diretamente os arquivos Excel exportados do SIE.

### Exemplo de uso:

```bash
python processar_excel.py -f "C:/Users/SeuUsuario/Dropbox/CEFET/COSAC/ComissaoCADD/11.02.05.99.60-GPROD.xlsx"
```

Esse comando processa o arquivo Excel informado e gera um resumo com aprovações, reprovações, trancamentos e matrículas.
Os resultados podem ser exibidos no terminal e/ou exportados para novos arquivos.

---

## Exemplos de saída

### Saída no terminal

```text
Total de linhas = 12500
# linhas processadas 500
# linhas processadas 1000
# linhas processadas 1500
...
# matriculas = 320
(2019, (350 aprovações, 120 reprovações, 45 trancamentos))
(2020, (400 aprovações, 150 reprovações, 60 trancamentos))
(2021, (280 aprovações, 90 reprovações, 30 trancamentos))
```

### Relatório em Excel (`resultados/GADM-faixas-criticidade.xlsx`)

| MATR\_ALUNO | NOME\_ALUNO     | QTD\_MAX\_REPROVACOES | QTD\_PERIODOS\_CURSADOS | CRITICIDADE |
| ----------- | --------------- | --------------------- | ----------------------- | ----------- |
| 2020123456  | João da Silva   | 1                     | 3                       | AZUL        |
| 2020234567  | Maria Oliveira  | 2                     | 4                       | LARANJA     |
| 2019345678  | Pedro Fernandes | 3                     | 6                       | VERMELHA    |
| 2018234569  | Ana Souza       | 5                     | 8                       | PRETA       |

### Relatório gráfico (`resultados/GADM-reprovacoes.html`)

Um gráfico interativo em **Plotly** será gerado, mostrando:

* No eixo X: quantidade máxima de reprovações em uma disciplina
* No eixo Y: número de alunos nessa situação

Exemplo:

* `= 0 → 350 alunos`
* `= 1 → 120 alunos`
* `>= 5 → 25 alunos`