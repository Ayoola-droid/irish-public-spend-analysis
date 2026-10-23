from combine import load_all
from suppliers import add_supplier_clean

spend, recon = load_all()
spend = add_supplier_clean(spend)

# 1. Identical rows within the same quarter
key = ["quarter", "supplier_clean", "euro", "description"]
spend["dup_flag"] = spend.duplicated(subset=key, keep=False)
extra_copies = spend.duplicated(subset=key)

groups = (spend[spend["dup_flag"]]
          .groupby(key).size().reset_index(name="copies")
          .sort_values("euro", ascending=False))

print(f"Rows flagged as possible duplicates: {spend['dup_flag'].sum()}")
print(f"Duplicate groups: {len(groups)}")
print(f"Value of the extra copies: €{spend.loc[extra_copies, 'euro'].sum():,.2f}")
print("\nLargest duplicate groups:")
print(groups.head(10).to_string(index=False))

# 2. Identical orders that repeat across different quarters
repeat = (spend.groupby(["supplier_clean", "euro", "description"])["quarter"]
               .agg(quarters="nunique", seen_in=lambda q: ", ".join(sorted(q.unique())))
               .reset_index())
repeat = repeat[repeat["quarters"] > 1].sort_values("quarters", ascending=False)
print(f"\nIdentical orders appearing in more than one quarter: {len(repeat)}")
print(repeat.head(10).to_string(index=False))

# 3. What was ABM paid for?
abm = spend[spend["supplier_clean"] == "ABM CONTRACTORS LTD"]
print("\nABM Contractors by description:")
print(abm.groupby("description")["euro"].agg(orders="size", total="sum")
         .sort_values("total", ascending=False).to_string())
# 4. Descriptions that differ only by capital letters
desc = spend["description"].dropna()
variants = desc.groupby(desc.str.lower()).unique()
variants = variants[variants.str.len() > 1]
print(f"\nDescriptions with case-only variants: {len(variants)}")
for v in variants:
    print("  ", list(v))
print(f"Distinct descriptions: {spend['description'].nunique()}")