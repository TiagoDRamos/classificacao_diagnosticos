import os
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl


PROJECT_ROOT = Path(__file__).resolve().parent.parent


AGE_GROUPS = [
    ("<=29", 0, 29),
    ("30-39", 30, 39),
    ("40-49", 40, 49),
    ("50-59", 50, 59),
    ("60-69", 60, 69),
    ("70-79", 70, 79),
    ("80>=", 80, 200),
]


def normalize_sexo(value):
    if value is None:
        return None
    normalized = str(value).strip().lower()

    if normalized in {"m", "male", "masculino", "homem"}:
        return "Male"
    if normalized in {"f", "female", "feminino", "mulher"}:
        return "Female"
    return None


def age_group(age):
    try:
        age_value = int(float(age))
    except (TypeError, ValueError):
        return None

    if age_value < 0:
        return None

    for label, start, end in AGE_GROUPS:
        if start <= age_value <= end:
            return label

    return None


def load_xlsx(file_name: str):
    path = Path(file_name)
    workbook = openpyxl.load_workbook(path, data_only=True)
    return workbook


def get_column_index(headers, target_name):
    for index, header in enumerate(headers):
        if header is not None and str(header).strip().lower() == target_name.lower():
            return index
    raise ValueError(f"Column '{target_name}' not found in the Excel file.")


def count_by_age_group_and_sexo(workbook):
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))

    if not rows:
        raise ValueError("The Excel file is empty.")

    headers = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
    sexo_index = get_column_index(headers, "Sexo")
    idade_index = get_column_index(headers, "Idade")

    counts = {label: {"Male": 0, "Female": 0} for label, _, _ in AGE_GROUPS}

    for row in rows[1:]:
        if len(row) <= max(sexo_index, idade_index):
            continue

        sexo = normalize_sexo(row[sexo_index])
        idade_label = age_group(row[idade_index])

        if sexo is None or idade_label is None:
            continue

        counts[idade_label][sexo] += 1

    return counts


def plot_counts(counts, output_file="sexo_idade_plot.pdf"):
    labels = [group for group, _, _ in AGE_GROUPS]
    male_counts = [counts[group]["Male"] for group in labels]
    female_counts = [counts[group]["Female"] for group in labels]

    x = range(len(labels))
    width = 0.35



    fig, ax = plt.subplots(figsize=(10, 6))
    ax.bar([i - width / 2 for i in x], male_counts, width=width, label="Male", color="tab:blue", alpha=0.8)
    ax.bar([i + width / 2 for i in x], female_counts, width=width, label="Female", color="tab:pink", alpha=0.8)
    ax.set_title("Total male and female by age group")
    ax.set_xlabel("Age group")
    ax.set_ylabel("Total")
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels)
    ax.legend()
    plt.tight_layout()
    Path(output_file).parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(output_file, format="pdf")
    plt.close(fig)

    try:
        if sys.platform.startswith("darwin"):
            os.system(f"open {output_file}")
        elif sys.platform.startswith("linux"):
            os.system(f"xdg-open {output_file}")
        elif sys.platform.startswith("win"):
            os.system(f"start {output_file}")
    except Exception:
        pass

    print(f"Chart saved to: {output_file}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python scripts/analise_doentes.py <path_to_excel_file>")
        sys.exit(1)

    file_path = sys.argv[1]
    workbook = load_xlsx(file_path)
    counts = count_by_age_group_and_sexo(workbook)
    plot_counts(counts, PROJECT_ROOT / "outputs" / "sexo_idade_plot.pdf")
