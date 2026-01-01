import csv
from pathlib import Path

PAIRWISE_FUN_FILE = Path("pairwise_fun_comparisons.csv")
PAIRWISE_STRENGTH_FILE = Path("pairwise_strength_comparisons.csv")

for file_path in [PAIRWISE_FUN_FILE, PAIRWISE_STRENGTH_FILE]:
    if not file_path.exists():
        with file_path.open(mode="w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(["Role1", "Role2", "SelectedRole"])

def save_pairwise_comparison(role1: str, role2: str, selected_role: str, category: str):
    file_path = PAIRWISE_FUN_FILE if category == "fun" else PAIRWISE_STRENGTH_FILE
    with file_path.open(mode="a", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow([role1, role2, selected_role])

def generate_role_ranking(category: str):
    file_path = PAIRWISE_FUN_FILE if category == "fun" else PAIRWISE_STRENGTH_FILE
    role_scores = {}
    role_comparisons = {}
    pairwise_results = []

    with file_path.open(mode="r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            role1, role2, selected_role = row["Role1"], row["Role2"], row["SelectedRole"]

            pairwise_results.append((role1, role2, selected_role))
            if role1 not in role_scores:
                role_scores[role1] = 0
                role_comparisons[role1] = 0
            if role2 not in role_scores:
                role_scores[role2] = 0
                role_comparisons[role2] = 0

            if selected_role == role1:
                role_scores[role1] += 1
                role_scores[role2] -= 1
            elif selected_role == role2:
                role_scores[role2] += 1
                role_scores[role1] -= 1

            role_comparisons[role1] += 1
            role_comparisons[role2] += 1

    # transitive weighting
    for role1, role2, selected_role in pairwise_results:
        if selected_role == role1:
            for transitive_role, _, transitive_selected in pairwise_results:
                if transitive_role == role2 and transitive_selected == role2:
                    role_scores[role1] += 0.5
        elif selected_role == role2:
            for transitive_role, _, transitive_selected in pairwise_results:
                if transitive_role == role1 and transitive_selected == role1:
                    role_scores[role2] += 0.5

    # Calculate weighted scores
    weighted_scores = {
        role: score / role_comparisons[role] if role_comparisons[role] > 0 else 0
        for role, score in role_scores.items()
    }

    ranked_roles = sorted(weighted_scores.items(), key=lambda x: x[1], reverse=True)
    return ranked_roles