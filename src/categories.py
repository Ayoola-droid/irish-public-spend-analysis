from pathlib import Path
import pandas as pd

RULES = Path("data/category_rules.csv")


def categorise(description, rules):
    text = str(description).lower()
    for keyword, category in rules:
        if keyword in text:
            return category
    return "Uncategorised"


def add_category(spend):
    spend = spend.copy()
    rules_df = pd.read_csv(RULES)
    rules = list(zip(rules_df["keyword"].str.lower(), rules_df["category"]))
    spend["category"] = spend["description"].map(lambda d: categorise(d, rules))
    return spend


if __name__ == "__main__":
    from combine import load_all

    spend, recon = load_all()
    spend = add_category(spend)
    total = spend["euro"].sum()

    summary = (spend.groupby("category")["euro"]
                    .agg(orders="size", spend="sum")
                    .sort_values("spend", ascending=False))
    summary["share_%"] = (summary["spend"] / total * 100).round(1)
    print(summary.to_string())

    print("\nDescriptions in each category:")
    for cat, group in spend.groupby("category"):
        print(f"\n{cat}: {', '.join(sorted(group['description'].unique()))}")