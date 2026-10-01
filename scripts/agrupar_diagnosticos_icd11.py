import argparse
import csv
import re
import unicodedata
from pathlib import Path


CABECALHOS_OBRIGATORIOS = {
    "Diagnóstico",
    "Grupo ICD-11",
    "Total em Diagnóstico principal",
    "Total em Diagnóstico secundário",
}
ALIASES_DIAGNOSTICOS = {
    "leiomioma útero": "Leiomioma do útero",
    "leiomioma uterino": "Leiomioma do útero",
    "mioma uterino": "Leiomioma do útero",
}
PADRAO_CODIGO_ICD11 = re.compile(
    r"^\s*([0-9A-Z][0-9A-Z]{1,5}(?:\.[0-9A-Z]+)?\.?)\s+(.+?)\s*$",
    re.IGNORECASE,
)
PROJECT_ROOT = Path(__file__).resolve().parent.parent


def nome_diagnostico(diagnostico):
    correspondencia = PADRAO_CODIGO_ICD11.match(diagnostico)
    if correspondencia:
        codigo = correspondencia.group(1)
        if any(carater.isdigit() for carater in codigo):
            diagnostico = correspondencia.group(2).strip()

    chave = re.sub(r"\s+", " ", diagnostico).strip().casefold()
    return ALIASES_DIAGNOSTICOS.get(chave, diagnostico)


def ler_diagnosticos(ficheiro_entrada):
    with open(ficheiro_entrada, "r", newline="", encoding="utf-8-sig") as ficheiro:
        primeira_linha = ficheiro.readline()
        if not primeira_linha:
            raise ValueError("O ficheiro CSV está vazio.")

        if primeira_linha.strip().lower() != "sep=;":
            ficheiro.seek(0)

        leitor = csv.DictReader(ficheiro, delimiter=";")
        cabecalhos = set(leitor.fieldnames or [])
        em_falta = CABECALHOS_OBRIGATORIOS - cabecalhos
        if em_falta:
            nomes = ", ".join(sorted(em_falta))
            raise ValueError(f"Colunas em falta no CSV: {nomes}")

        return list(leitor)


def agrupar_diagnosticos(linhas):
    agregados = {}

    for linha in linhas:
        grupo = (linha["Grupo ICD-11"] or "Não indicado").strip()
        diagnostico = nome_diagnostico(
            (linha["Diagnóstico"] or "").strip()
        )

        if not diagnostico:
            continue

        chave = (grupo, diagnostico.casefold())
        if chave not in agregados:
            agregados[chave] = [grupo, diagnostico, 0, 0]

        agregados[chave][2] += int(
            linha["Total em Diagnóstico principal"] or 0
        )
        agregados[chave][3] += int(
            linha["Total em Diagnóstico secundário"] or 0
        )

    return sorted(
        agregados.values(),
        key=lambda linha: (linha[0].casefold(), linha[1].casefold()),
    )


def nome_seguro(nome):
    nome = unicodedata.normalize("NFKD", nome)
    nome = nome.encode("ascii", "ignore").decode("ascii")
    nome = re.sub(r"[^A-Za-z0-9._-]+", "_", nome).strip("._")
    return nome or "grupo_sem_nome"


def guardar_csv(linhas, pasta_saida):
    cabecalhos = [
        "Diagnóstico",
        "Total em Diagnóstico principal",
        "Total em Diagnóstico secundário",
    ]
    grupos = {}
    for linha in linhas:
        grupos.setdefault(linha[0], []).append(linha[1:])

    pasta_saida = Path(pasta_saida)
    pasta_saida.mkdir(parents=True, exist_ok=True)
    ficheiros = []
    linhas_resumo = []
    total_geral_principal = 0
    total_geral_secundario = 0

    for grupo, diagnosticos in grupos.items():
        ficheiro_saida = pasta_saida / f"grupo_{nome_seguro(grupo)}.csv"
        total_principal = sum(linha[1] for linha in diagnosticos)
        total_secundario = sum(linha[2] for linha in diagnosticos)
        linhas_resumo.append([grupo, total_principal, total_secundario])
        total_geral_principal += total_principal
        total_geral_secundario += total_secundario
        with ficheiro_saida.open("w", newline="", encoding="utf-8-sig") as ficheiro:
            escritor = csv.writer(ficheiro, delimiter=";")
            escritor.writerow(cabecalhos)
            escritor.writerows(diagnosticos)
            escritor.writerow([])
            escritor.writerow(["TOTAL", total_principal, total_secundario])
        ficheiros.append(ficheiro_saida)

    ficheiro_resumo = pasta_saida / "resumo_totais_icd11.csv"
    with ficheiro_resumo.open("w", newline="", encoding="utf-8-sig") as ficheiro:
        escritor = csv.writer(ficheiro, delimiter=";")
        escritor.writerow([
            "Grupo ICD-11",
            "Total em Diagnóstico principal",
            "Total em Diagnóstico secundário",
        ])
        escritor.writerows(linhas_resumo)
        escritor.writerow([
            "TOTAL GERAL",
            total_geral_principal,
            total_geral_secundario,
        ])

    return ficheiros


def main():
    parser = argparse.ArgumentParser(
        description="Agrupa diagnósticos por grupo ICD-11."
    )
    parser.add_argument(
        "ficheiro_entrada",
        nargs="?",
        default=str(PROJECT_ROOT / "outputs" / "diagnosticos_contagem_icd11.csv"),
        help="CSV ICD-11 de entrada",
    )
    parser.add_argument(
        "-o",
        "--saida",
        default=str(PROJECT_ROOT / "outputs" / "diagnosticos_por_grupo_icd11"),
        help="Pasta onde serão criados os CSVs, um por grupo",
    )
    args = parser.parse_args()

    linhas = ler_diagnosticos(args.ficheiro_entrada)
    resultados = guardar_csv(
        agrupar_diagnosticos(linhas),
        args.saida,
    )
    print(f"Tabelas criadas: {len(resultados)}")
    for resultado in resultados:
        print(f"Resultado guardado em: {resultado}")
    print(f"Resumo guardado em: {Path(args.saida) / 'resumo_totais_icd11.csv'}")


if __name__ == "__main__":
    main()