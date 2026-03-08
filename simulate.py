import pandas as pd
import numpy as np
from datetime import date, timedelta
import random

random.seed(42)
np.random.seed(42)

# ─── Part 1: Policy Sales Data ───────────────────────────────────────────────

TOTAL_CUSTOMERS = 1_000_000
VEHICLE_VALUE = 100_000
PREMIUM_PER_YEAR = 100

tenure_dist = {1: 0.20, 2: 0.30, 3: 0.40, 4: 0.10}
tenures = list(tenure_dist.keys())
probs   = list(tenure_dist.values())

# 1M customers spread evenly across 366 days of 2024 (leap year)
start_2024 = date(2024, 1, 1)
days_2024  = 366  # 2024 is a leap year
per_day    = TOTAL_CUSTOMERS // days_2024
remainder  = TOTAL_CUSTOMERS - per_day * days_2024

records = []
cust_id = 1
for d in range(days_2024):
    purchase_date = start_2024 + timedelta(days=d)
    count = per_day + (1 if d < remainder else 0)
    day_tenures = np.random.choice(tenures, size=count, p=probs)
    for t in day_tenures:
        policy_start = purchase_date + timedelta(days=365)
        policy_end   = policy_start + timedelta(days=int(t) * 365)
        records.append({
            "Customer_ID":        f"C{cust_id:08d}",
            "Vehicle_ID":         f"V{cust_id:08d}",
            "Vehicle_Value":      VEHICLE_VALUE,
            "Premium":            t * PREMIUM_PER_YEAR,
            "Policy_Purchase_Date": purchase_date,
            "Policy_Start_Date":  policy_start,
            "Policy_End_Date":    policy_end,
            "Policy_Tenure":      t,
        })
        cust_id += 1

policies = pd.DataFrame(records)
print(f"Policies created: {len(policies):,}")

# ─── Part 2: Claims Data ──────────────────────────────────────────────────────

CLAIM_AMOUNT = VEHICLE_VALUE * 0.10   # 10,000

claims = []
claim_id = 1

# 2025 claims: vehicles purchased on 7th/14th/21st/28th, 30% file claim on policy start date
special_days = {7, 14, 21, 28}

def is_special(d):
    return d.day in special_days

policies_2025_eligible = policies[
    policies["Policy_Purchase_Date"].apply(is_special)
].copy()

# 30% of these
n_claim_2025 = int(len(policies_2025_eligible) * 0.30)
chosen_2025 = policies_2025_eligible.sample(n=n_claim_2025, random_state=42)

filed_2025_ids = set()
for _, row in chosen_2025.iterrows():
    claim_date = row["Policy_Start_Date"]
    # Policy must be active: claim_date >= start and claim_date < end
    if row["Policy_Start_Date"] <= claim_date < row["Policy_End_Date"]:
        claims.append({
            "Claim_ID":    f"CLM{claim_id:08d}",
            "Customer_ID": row["Customer_ID"],
            "Vehicle_ID":  row["Vehicle_ID"],
            "Claim_Amount": CLAIM_AMOUNT,
            "Claim_Date":  claim_date,
            "Claim_Type":  1,
        })
        filed_2025_ids.add(row["Vehicle_ID"])
        claim_id += 1

print(f"2025 claims: {len(filed_2025_ids):,}")

# 2026 claims (Jan 1 – Feb 28): 10% of 4-year tenure vehicles
# Distributed evenly across 59 days
period_start = date(2026, 1, 1)
period_end   = date(2026, 2, 28)
days_2026    = 59

four_year = policies[policies["Policy_Tenure"] == 4].copy()
n_claim_2026 = int(len(four_year) * 0.10)
chosen_2026 = four_year.sample(n=n_claim_2026, random_state=99)

per_day_2026 = n_claim_2026 // days_2026
rem_2026     = n_claim_2026 - per_day_2026 * days_2026

rows_2026 = chosen_2026.reset_index(drop=True)
idx = 0
for d in range(days_2026):
    claim_date = period_start + timedelta(days=d)
    count = per_day_2026 + (1 if d < rem_2026 else 0)
    for i in range(count):
        row = rows_2026.iloc[idx]
        # Policy active check
        if row["Policy_Start_Date"] <= claim_date < row["Policy_End_Date"]:
            is_second = row["Vehicle_ID"] in filed_2025_ids
            claims.append({
                "Claim_ID":    f"CLM{claim_id:08d}",
                "Customer_ID": row["Customer_ID"],
                "Vehicle_ID":  row["Vehicle_ID"],
                "Claim_Amount": CLAIM_AMOUNT,
                "Claim_Date":  claim_date,
                "Claim_Type":  2 if is_second else 1,
            })
            claim_id += 1
        idx += 1

claims_df = pd.DataFrame(claims)
print(f"Total claims: {len(claims_df):,}")
print(f"2025 claims: {len(claims_df[claims_df['Claim_Date'].apply(lambda x: x.year)==2025]):,}")
print(f"2026 claims: {len(claims_df[claims_df['Claim_Date'].apply(lambda x: x.year)==2026]):,}")

# ─── Part 3: Analytical Queries ───────────────────────────────────────────────

## Q1: Total premium collected 2024
total_premium = policies["Premium"].sum()
print(f"\nQ1 Total Premium 2024: ₹{total_premium:,.0f}")

## Q2: Total claim cost by year, monthly breakdown
claims_df["Year"]  = claims_df["Claim_Date"].apply(lambda x: x.year)
claims_df["Month"] = claims_df["Claim_Date"].apply(lambda x: x.month)

q2_monthly = claims_df.groupby(["Year","Month"])["Claim_Amount"].sum().reset_index()
q2_monthly.columns = ["Year","Month","Total_Claim_Cost"]
q2_yearly  = claims_df.groupby("Year")["Claim_Amount"].sum().reset_index()
print("\nQ2 Annual Claim Costs:")
print(q2_yearly.to_string(index=False))

## Q3: Claim-cost-to-premium ratio by tenure
tenure_premiums = policies.groupby("Policy_Tenure")["Premium"].sum()
claim_by_tenure = claims_df.merge(
    policies[["Vehicle_ID","Policy_Tenure"]], on="Vehicle_ID"
).groupby("Policy_Tenure")["Claim_Amount"].sum()

q3 = pd.DataFrame({
    "Total_Premium": tenure_premiums,
    "Total_Claims":  claim_by_tenure
}).fillna(0)
q3["Loss_Ratio"] = q3["Total_Claims"] / q3["Total_Premium"]
print("\nQ3 Loss Ratio by Tenure:")
print(q3.round(4))

## Q4: Claim-cost-to-premium ratio by month of policy sale
policies["Sale_Month"] = policies["Policy_Purchase_Date"].apply(lambda x: x.month)
month_prem = policies.groupby("Sale_Month")["Premium"].sum()

claims_with_sale = claims_df.merge(
    policies[["Vehicle_ID","Sale_Month"]], on="Vehicle_ID"
)
month_claims = claims_with_sale.groupby("Sale_Month")["Claim_Amount"].sum()

q4 = pd.DataFrame({
    "Total_Premium": month_prem,
    "Total_Claims":  month_claims
}).fillna(0)
q4["Loss_Ratio"] = q4["Total_Claims"] / q4["Total_Premium"]
print("\nQ4 Loss Ratio by Sale Month:")
print(q4.round(4))

## Q5: Potential claim liability — every vehicle that hasn't claimed files exactly one claim
# Vehicles that have already filed at least one claim
vehicles_claimed = set(claims_df["Vehicle_ID"].unique())
vehicles_unclaimed = policies[~policies["Vehicle_ID"].isin(vehicles_claimed)]
# Potential liability = count of unclaimed * claim amount
potential_liability = len(vehicles_unclaimed) * CLAIM_AMOUNT
print(f"\nQ5 Potential Claim Liability: ₹{potential_liability:,.0f}")
print(f"   Unclaimed vehicles: {len(vehicles_unclaimed):,}")

## Q6: Earned premium up to Feb 28 2026 & monthly estimate for remaining
AS_OF = date(2026, 2, 28)
TOTAL_REMAINING_MONTHS = 46

def days_active_until(row, as_of):
    s = row["Policy_Start_Date"]
    e = row["Policy_End_Date"]
    if as_of < s:
        return 0
    return (min(as_of, e) - s).days

policies["Active_Days_To_AsOf"] = policies.apply(lambda r: days_active_until(r, AS_OF), axis=1)
policies["Total_Tenure_Days"]   = policies["Policy_Tenure"] * 365
policies["Daily_Premium"]       = policies["Premium"] / policies["Total_Tenure_Days"]
policies["Earned_To_AsOf"]      = policies["Daily_Premium"] * policies["Active_Days_To_AsOf"]

total_earned = policies["Earned_To_AsOf"].sum()
print(f"\nQ6a Earned Premium to Feb 28 2026: ₹{total_earned:,.2f}")

# Monthly estimate for remaining period
# Only policies still active after Feb 28 2026
def remaining_days_after(row, as_of):
    if row["Policy_End_Date"] <= as_of:
        return 0
    return (row["Policy_End_Date"] - max(row["Policy_Start_Date"], as_of)).days

policies["Remaining_Days"] = policies.apply(lambda r: remaining_days_after(r, AS_OF), axis=1)
policies["Remaining_Premium"] = policies["Daily_Premium"] * policies["Remaining_Days"]

total_remaining_premium = policies["Remaining_Premium"].sum()
monthly_est = total_remaining_premium / TOTAL_REMAINING_MONTHS
print(f"Q6b Estimated Monthly Premium (46 months): ₹{monthly_est:,.2f}")

# ─── Save results ─────────────────────────────────────────────────────────────
print("\nSaving datasets...")
policies.to_csv("/home/claude/policy_sales_data.csv", index=False)
claims_df.to_csv("/home/claude/claims_data.csv", index=False)
q2_monthly.to_csv("/home/claude/q2_monthly_claims.csv", index=False)
q3.reset_index().to_csv("/home/claude/q3_tenure_ratio.csv", index=False)
q4.reset_index().to_csv("/home/claude/q4_month_ratio.csv", index=False)

# Store key numbers for Excel
import json
summary = {
    "total_premium": float(total_premium),
    "total_claim_2025": float(q2_yearly[q2_yearly["Year"]==2025]["Claim_Amount"].values[0]),
    "total_claim_2026": float(q2_yearly[q2_yearly["Year"]==2026]["Claim_Amount"].values[0]) if 2026 in q2_yearly["Year"].values else 0,
    "potential_liability": float(potential_liability),
    "unclaimed_vehicles": int(len(vehicles_unclaimed)),
    "total_vehicles_claimed": int(len(vehicles_claimed)),
    "earned_to_feb2026": float(total_earned),
    "monthly_est_remaining": float(monthly_est),
    "total_remaining_premium": float(total_remaining_premium),
}
with open("/home/claude/summary_numbers.json","w") as f:
    json.dump(summary, f, indent=2)

print("Done!")
