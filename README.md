# Análise de diagnósticos

Scripts Python para resumir diagnósticos de um ficheiro Excel, classificá-los pela ICD-11 através da API ICD-11 MMS da OMS (v2) e criar tabelas por grupo.

## Preparação

Requer Python 3.10 ou superior. Instala as dependências:

```bash
python -m pip install -r requirements.txt
```

Podes usar o Excel local `data/base_ultra_final.xlsx` ou fornecer diretamente o URL normal de uma folha Google Sheets. A pasta `data/` é local e ignorada pelo Git porque o Excel pode conter dados clínicos. Não publiques esse ficheiro nem resultados que possam identificar doentes.

### Configurar o Google Sheets privado

1. Ativa a Google Sheets API num projeto Google Cloud e cria uma conta de serviço.
2. Guarda o JSON da conta de serviço fora do repositório.
3. Partilha a folha apenas com o email da conta de serviço, com permissão de leitor. Não publiques a folha na web.
4. Define `GOOGLE_APPLICATION_CREDENTIALS` para o caminho do JSON. No Linux/macOS:

```bash
export GOOGLE_APPLICATION_CREDENTIALS="$HOME/.config/google/service-account.json"
```

O script usa apenas o escopo de leitura da Google Sheets API. Não coloques o JSON de credenciais no projeto nem no GitHub.

## Gerar resultados

Executa os comandos a partir da raiz do projeto. A classificação de diagnósticos sem código local requer o serviço ICD-11 disponível em `http://localhost:8000`; os códigos presentes no texto ou na lista de fallback são processados localmente.

```bash
python scripts/contar_diagnosticos.py "https://docs.google.com/spreadsheets/d/ID_DA_FOLHA/edit" \
  --csv \
  -o diagnosticos_contagem.csv \
  --icd11-saida diagnosticos_contagem_icd11.csv

python scripts/agrupar_diagnosticos_icd11.py
```

Os CSVs por grupo e `resumo_totais_icd11.csv` são escritos em `outputs/diagnosticos_por_grupo_icd11/`.
Quando indicas apenas o nome do ficheiro nas opções `-o` e `--icd11-saida`, o script guarda-o em `outputs/` e cria essa pasta se necessário. Também podes passar um caminho completo ou relativo explícito.

Para ler o Excel local em vez do Google Sheets:

```bash
python scripts/contar_diagnosticos.py data/base_ultra_final.xlsx \
  --csv \
  -o diagnosticos_contagem.csv \
  --icd11-saida diagnosticos_contagem_icd11.csv
```

Para gerar o gráfico de idade e sexo:

```bash
python scripts/analise_doentes.py data/base_ultra_final.xlsx
```

Os resultados ficam em `outputs/`. Esta pasta também é ignorada pelo Git.