# %%

import duckdb

# %%

suppliersFile = "./files/oracle suppliers for prama.csv"
outputFile = "./files/pcard_monthly_output_jan24.csv"
outputExcelFile = "./files/pcard_monthly_output_jan24.xlsx"
monthlyTransactionsFile = "./files/1.31.24-Monthly_w_GLInternal_Audit_Record.xlsx"
cardLimitsMappingsFile = "./files/Pcard Mapping-updated 2.2.24.xlsx"
otherMappingsFile = "./files/Regions Monthly_w_GLRevised_Record 2022-2023-Updated.xlsx"


# %%
con = duckdb.connect()

con.sql("INSTALL spatial;")
con.sql("LOAD spatial;")

# %%

# load pcard monthly transactions

con.sql(
    f"""CREATE TABLE pcard_monthly AS SELECT * FROM st_read('{monthlyTransactionsFile}', open_options=['HEADERS=FORCE']);"""
)

print(con.sql("SELECT COUNT(*) FROM pcard_monthly"))

# %%

column_name_mappings = [
    ("Company Name", "Company.Company Legal Name"),
    ("Hierarchy Name", "Company Unit.Hierarchy Name"),
    ("GL Account", "Financial Codes.General Ledger Code"),
    ("Note", "Transaction User Defined.Note"),
    ("Split Description", "Transaction Splits.Split Description"),
    ("Cost Center", "Financial Codes.Cost Center Code"),
    ("Project Code", "Financial Codes.Project Code"),
    ("Split Amount", "Transaction Splits.Split Amount"),
    ("Merchant", "Transaction.Merchant Name"),
    ("Post Date", "Transaction.Posting Dt"),
    ("Receipts", "Transaction.Receipts"),
    ("Review Date", "Transaction.Review Date"),
    ("Approve Date", "Transaction.Approved Dt"),
    ("Approve 2 Date", "Transaction.Approved2 Dt"),
    ("In Envelope", "Transaction.In Envelope"),
    ("First Name", "Cardholder.First Name"),
    ("Last Name", "Cardholder.Last Name"),
    ("SIC MCC CD", "Transaction.SIC Merchant Category Code CD"),
    ("Transaction Date", "Transaction.Transaction Dt"),
]

# %%

for c in column_name_mappings:
    con.sql(
        "ALTER TABLE pcard_monthly RENAME COLUMN "
        + f"'{c[0]}'"
        + " TO "
        + f"'{c[1]}'"
        + ";"
    )

# %%

# load card limits

con.sql(
    f"""CREATE TABLE card_limits AS SELECT * FROM st_read("{cardLimitsMappingsFile}", layer = 'Mapping - Card Limits-Loc-Categ', open_options=['HEADERS=FORCE']);"""
)

# %%

con.sql(
    """UPDATE card_limits SET "Credit Limit" = regexp_replace("Credit Limit",'US', '', 'g') """
)

con.commit()

# %%

con.sql(
    """UPDATE card_limits SET "Current Balance" = regexp_replace("Current Balance",'US', '', 'g') """
)

con.commit()

# %%

con.sql(
    """UPDATE card_limits SET "Current Statement" = regexp_replace("Current Statement",'US', '', 'g') """
)

con.commit()

# %%

# load card and grouping from pivot mapping

con.sql(
    f"""CREATE TABLE pivot_mapping AS SELECT * FROM st_read('{monthlyTransactionsFile}', layer = 'Pivot-Mapping Example', open_options=['HEADERS=FORCE']);"""
)

print(con.sql("SELECT COUNT(*) FROM pivot_mapping"))

# %%
    
# Testing new logic...
    
con.sql(
    """
CREATE TABLE new_logic_tests_1 AS 
SELECT  t.*
FROM    pcard_monthly t
"""
)


# %%

r = con.sql(
    """
ALTER TABLE new_logic_tests_1 ADD COLUMN FN_LN_1 STRING DEFAULT ' '
"""
)

con.commit()

print(r)

# %%

r = con.sql("""
 UPDATE new_logic_tests_1 SET FN_LN_1 = "Cardholder.First Name" || ' ' || "Cardholder.Last Name"
""")

con.commit()

print(r)

# %%

r = con.sql("""
 UPDATE new_logic_tests_1 SET FN_LN_1 = replace(FN_LN_1, 'N/A ', '')
""")

con.commit()

print(r)

# %%

r = con.sql(
    """
ALTER TABLE new_logic_tests_1 ADD COLUMN MAP_COUNT INTEGER DEFAULT -1
"""
)

# %%

r = con.sql("""
UPDATE new_logic_tests_1 SET MAP_COUNT = (SELECT COUNT(*) FROM card_limits cl WHERE "Company Unit.Hierarchy Name" = cl."Hierarchy Node Name" AND FN_LN_1 = replace(cl."First Name" || ' ' || cl."Last Name", 'N/A ', ''))
""")

# %%

# %%

r = con.sql("""
UPDATE new_logic_tests_1 SET MAP_COUNT = (SELECT COUNT(*) FROM pivot_mapping pm WHERE "Company Unit.Hierarchy Name" = pm."Hierarchy Name" AND "Cardholder.First Name" = pm."First Name" AND "Cardholder.Last Name" = pm."Last Name")
""")



# %%

print(con.sql("select * from new_logic_tests_1 limit 5"))

# %%

con.sql("""COPY new_logic_tests_1 TO './files/test_out.csv' (HEADER, DELIMITER ',');""")

# %%

con.sql(
    """
CREATE TABLE new_logic_tests_2 AS 
SELECT  t.*, m.* 
FROM    new_logic_tests_1 t
LEFT JOIN dept_loc m
    ON  t."Financial Codes.Project Code" = m."Financial Codes.Project Code"
"""
)