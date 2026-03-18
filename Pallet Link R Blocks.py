import pandas as pd

# File path
file_path = r"C:\Users\KMuthusi\OneDrive - Anheuser-Busch InBev\Desktop\R Blocks Resolution\Pallet Link R Blocks.xlsx"

# Load sheets
sheet1 = pd.read_excel(file_path, sheet_name="Sheet1")
sheet2 = pd.read_excel(file_path, sheet_name="Sheet2")

# Clean column names
sheet1.columns = sheet1.columns.str.strip()
sheet2.columns = sheet2.columns.str.strip()

# Ensure correct types
sheet1["Material Descr(MM60)"] = sheet1["Material Descr(MM60)"].astype(str)
sheet1["Plant"] = sheet1["Plant"].astype(str).str.strip()

sheet2["Plant Code"] = sheet2["Plant Code"].astype(str).str.strip()

# Fix things like "1721.0"
sheet1["Plant"] = sheet1["Plant"].str.replace(".0", "", regex=False)
sheet2["Plant Code"] = sheet2["Plant Code"].str.replace(".0", "", regex=False)

# Tracking
total_records = len(sheet1)
match_count = 0
unmatched_invoices = []

# Loop through Sheet1
for _, row in sheet1.iterrows():

    material_desc = row["Material Descr(MM60)"]
    plant = row["Plant"]
    invoice_no = row["Invoice Document No."]

    matched = False

    # Step 1: Check if material starts with "Pallet"
    if material_desc.strip().lower().startswith("pallet"):

        # Step 2: Match Plant with Plant Code in Sheet2
        match_row = sheet2[sheet2["Plant Code"] == plant]

        if not match_row.empty:

            matched_row = match_row.iloc[0]
            matched = True
            match_count += 1

            # Safe numeric conversion
            po_price = pd.to_numeric(row["PO Unit Price"], errors="coerce")
            inv_price = pd.to_numeric(row["Inv Unit Price"], errors="coerce")
            palkor_price = pd.to_numeric(matched_row["Palkor"], errors="coerce")

            print("========== MATCH ==========")
            print("Invoice Document No.:", invoice_no)
            print("Material Descr(MM60):", material_desc)
            print("Plant:", plant)

            #if pd.notna(po_price):
            #    print("PO Unit Price (/1000):", po_price / 1000)
            # else:
            #    print("PO Unit Price (/1000): N/A")

            print("PO Unit Price:", po_price)
            print("Inv Unit Price:", inv_price)

            print("\nSheet2 (Lookup):")
            print("Plant Name:", matched_row["Plant"])
            print("Plant Code:", matched_row["Plant Code"])
            print("Palkor (Delivered Price):", palkor_price)
            #print("PO Unit Price (/1000):", po_price / 1000)

            print("===========================\n")

    # Track unmatched
    if not matched:
        unmatched_invoices.append(invoice_no)

# ===== SUMMARY =====
print("\n========== SUMMARY ==========")
print(f"Matched: {match_count}/{total_records}")
print(f"Unmatched: {len(unmatched_invoices)}")

if unmatched_invoices:
    print("\nUnmatched Invoice Document Numbers:")
    for inv in unmatched_invoices:
        print("-", inv)

print("=============================")

pd.DataFrame({"Unmatched Invoices": unmatched_invoices}) \
    .to_excel("unmatched_pallets.xlsx", index=False)