# %%

import re
import duckdb
from duckdb.typing import BOOLEAN, VARCHAR

# %%

suppliersFile = "./files/oracle suppliers for prama.csv"

# Jan 2025 - Still has unmatched card names - Need to check
outputFile = "./files/Test/pcard_monthly_output_feb25.csv"
outputExcelFile = "./files/Test/pcard_monthly_output_feb25.xlsx"
monthlyTransactionsFile = "./files/01.31.25-Monthly_w_GLInternal_Audit_Record.xlsx"


# Dec 2024 - Still has unmatched card names - Need to check
# outputFile = "./files/pcard_monthly_output_dec24.csv"
# outputExcelFile = "./files/pcard_monthly_output_dec24.xlsx"
# monthlyTransactionsFile = "./files/12.31.24-Monthly_w_GLInternal_Audit_Record.xlsx"


# Nov 2024 - Still has unmatched card names - Need to check
# outputFile = "./files/pcard_monthly_output_nov24.csv"
# outputExcelFile = "./files/pcard_monthly_output_nov24.xlsx"
# monthlyTransactionsFile = "./files/11.30.24-Monthly_w_GLInternal_Audit_Record.xlsx"


# Oct 2024 - Still has unmatched card names - Need to check
# outputFile = "./files/pcard_monthly_output_oct24.csv"
# outputExcelFile = "./files/pcard_monthly_output_oct24.xlsx"
# monthlyTransactionsFile = "./files/10.31.24 Monthly_w_GLInternal_Audit_Record.xlsx"


# Sep 2024 - Still has unmatched card names - Need to check
# outputFile = "./files/pcard_monthly_output_sep24.csv"
# outputExcelFile = "./files/pcard_monthly_output_sep24.xlsx"
# monthlyTransactionsFile = "./files/09.30.24 Monthly_w_GLInternal_Audit_Record.xlsx"


# Aug 2024 - Still has unmatched card names - Need to check
# outputFile = "./files/pcard_monthly_output_aug24.csv"
# outputExcelFile = "./files/pcard_monthly_output_aug24.xlsx"
# monthlyTransactionsFile = "./files/08.30.24 Monthly_w_GLInternal_Audit_Record.xlsx"


# July 2024 - Final Run - Still has unmatched card names - Need to check
# outputFile = "./files/pcard_monthly_output_july24.csv"
# outputExcelFile = "./files/pcard_monthly_output_july24.xlsx"
# monthlyTransactionsFile = "./files/07.31.24 Final Monthly_w_GLInternal_Audit_Record.xlsx"

# # July 2024 - Early Run
# outputFile = "./files/pcard_monthly_output_july24.csv"
# outputExcelFile = "./files/pcard_monthly_output_july24.xlsx"
# monthlyTransactionsFile = "./files/07.31.24-Monthly_w_GLInternal_Audit_Record Early Run.xlsx"

# June 2024
# outputFile = "./files/pcard_monthly_output_june24.csv"
# outputExcelFile = "./files/pcard_monthly_output_june24.xlsx"
# monthlyTransactionsFile = "./files/06.30.24 Monthly_w_GLInternal_Audit_Record.xlsx"

# May 2024
# outputFile = "./files/pcard_monthly_output_may24.csv"
# outputExcelFile = "./files/pcard_monthly_output_may24.xlsx"
# monthlyTransactionsFile = "./files/05.31.24-Monthly_w_GLInternal_Audit_Record.xlsx"

# Apr 2024
# outputFile = "./files/pcard_monthly_output_apr24.csv"
# outputExcelFile = "./files/pcard_monthly_output_apr24.xlsx"
# monthlyTransactionsFile = "./files/04.30.24 Monthly_w_GLInternal_Audit_Record.xlsx"

# Mar 2024
# outputFile = "./files/pcard_monthly_output_mar24.csv"
# outputExcelFile = "./files/pcard_monthly_output_mar24.xlsx"
# monthlyTransactionsFile = "./files/03.31.24-Monthly_w_GLInternal_Audit_Record.xlsx"

# Feb 2024
# outputFile = "./files/pcard_monthly_output_feb24-new1.csv"
# outputExcelFile = "./files/pcard_monthly_output_feb24-new1.xlsx"
# monthlyTransactionsFile = "./files/InternalAuditPcardPreviousMonth_MonthlywGLInternalAudit_20240305064658.xlsx"

# Jan 2024
# outputFile = "./files/pcard_monthly_output_jan24-new1.csv"
# outputExcelFile = "./files/pcard_monthly_output_jan24-new1.xlsx"
# monthlyTransactionsFile = "./files/1.31.24-Monthly_w_GLInternal_Audit_Record.xlsx" # Jan 24 File

# Full 2 years history prior to 2024
# outputFile = "./files/pcard_monthly_output_full_history-new1.csv"
# outputExcelFile = "./files/pcard_monthly_output_full_history-new1.xlsx"
# monthlyTransactionsFile = "./files/Monthly_w_GLInternal_Audit_Record -2022-2023 - 3.14.24.xlsx" # Full History File

# Used since Final July 2024 Run
cardLimitsMappingsFile = "./files/Master- PCard-Mappings-as of 8.13.24.xlsx"
otherMappingsFile = "./files/Master- PCard-Mappings-as of 8.13.24.xlsx"
pivotMappingFile = "./files/Master- PCard-Mappings-as of 8.13.24.xlsx"

# # Used for Early July 2024 Run
# cardLimitsMappingsFile = "./files/Master- PCard-Mappings-as of 7.22.24.xlsx"
# otherMappingsFile = "./files/Master- PCard-Mappings-as of 7.22.24.xlsx"
# pivotMappingFile = "./files/Master- PCard-Mappings-as of 7.22.24.xlsx"


# cardLimitsMappingsFile = "./files/Go Prama Master- PCard-Mappings-Consolidated-03292024.xlsx"
# otherMappingsFile = "./files/Go Prama Master- PCard-Mappings-Consolidated-03292024.xlsx"
# pivotMappingFile = "./files/Go Prama Master- PCard-Mappings-Consolidated-03292024.xlsx"
# %%

exceptionsList = [
    "	Adult	",
    "	Appliances	",
    "	Artwork	",
    "	Bail	",
    "	Bars	",
    "	Beauty	",
    "	Betting	",
    "	Cash	",
    "	Casino	",
    "	Club	",
    "	Computer	",
    "	Concert	",
    "	Designer	",
    "	Dining	",
    "	Donation	",
    "	Electronics	",
    "	Entertainment	",
    "	Fine	",
    "	Firearms	",
    "	Fragrance	",
    "	Furniture	",
    "	Gaming	",
    "	Hobby	",
    "	Insurance	",
    "	Interest	",
    "	Jewelry	",
    "	Late	",
    "	Legal	",
    "	Leisure	",
    "	Lingerie	",
    "	Loan	",
    "	Lottery	",
    "	Luxury	",
    "	Massage	",
    "	Medical	",
    "	Membership	",
    "	Missing	",
    "	Nightclub	",
    "	Packages	",
    "	Penalty	",
    "	Personal	",
    "	Pet	",
    "	Political	",
    "	Renewal	",
    "	Repairs	",
    "	Spa	",
    "	Sports	",
    "	Storage	",
    "	Streaming	",
    "	Subscription	",
    "	Ticket	",
    "	Tobacco	",
    "	Vacation	",
    "	Vehicle	",
    "	Withdrawal	",
]

# "Travel" only for campuses - not for home office or others

cleanExceptionsList = []

for i in exceptionsList:
    cleanExceptionsList.append(i.strip().lower())


# %%

cleanExceptionsList

# %%

vendorsWatchList = [
    "Staples",
    "GFS",
    "Menards",
    "Lowes",
    "Home Depot",
    "Ivy Tech",
    "Positive Promotions",
]

# %%
con = duckdb.connect()

con.sql("INSTALL spatial;")
con.sql("LOAD spatial;")

# %%

# load pcard monthly transactions

con.sql(
    f"""--sql
    CREATE TABLE pcard_monthly AS SELECT * FROM st_read('{monthlyTransactionsFile}', open_options=['HEADERS=FORCE']);"""
)

print(
    con.sql("""--sql
            SELECT COUNT(*) FROM pcard_monthly""")
)

# %%

print(con.sql("DESCRIBE pcard_monthly"))

# ##

print(con.sql(""" SELECT DISTINCT "Company Name" from pcard_monthly """))

# %%

print(con.sql(""" SELECT DISTINCT "Cost Center" from pcard_monthly """))


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
    con.sql("ALTER TABLE pcard_monthly RENAME COLUMN " + f"'{c[0]}'" + " TO " + f"'{c[1]}'" + ";")

# %%

print(con.sql("DESCRIBE pcard_monthly"))

# ##

print(con.sql(""" SELECT DISTINCT "Company.Company Legal Name" from pcard_monthly """))

# %%

print(con.sql(""" SELECT DISTINCT "Financial Codes.Cost Center Code" from pcard_monthly """))


# %%

# load department to location mapping

con.sql(
    f"""--sql
    CREATE TABLE dept_loc AS SELECT * FROM st_read("{otherMappingsFile}", layer = 'Mapping -Dept-Loc', open_options=['HEADERS=FORCE']);"""
)

print(con.sql("SELECT COUNT(*) FROM dept_loc"))

print(con.sql("DESCRIBE dept_loc"))

# %%

# load account number to name mapping

con.sql(f"""CREATE TABLE acct_number_name AS SELECT * FROM st_read("{otherMappingsFile}", layer = 'Mapping - Acct #', open_options=['HEADERS=FORCE']);""")

print(con.sql("SELECT COUNT(*) FROM acct_number_name"))

print(con.sql("DESCRIBE acct_number_name"))

# %%

# load card limits

con.sql(f"""CREATE TABLE card_limits AS SELECT * FROM st_read("{cardLimitsMappingsFile}", layer = 'Mapping - Card Limits-Loc-Categ', open_options=['HEADERS=FORCE']);""")

# %%

print(con.sql("SELECT COUNT(*) FROM card_limits"))

print(con.sql("DESCRIBE card_limits"))

# %%

con.sql("""UPDATE card_limits SET "Credit Limit" = regexp_replace("Credit Limit",'US', '', 'g') """)

con.commit()

# %%

con.sql("""UPDATE card_limits SET "Current Balance" = regexp_replace("Current Balance",'US', '', 'g') """)

con.commit()

# %%

con.sql("""UPDATE card_limits SET "Current Statement" = regexp_replace("Current Statement",'US', '', 'g') """)

con.commit()

# %%
con.sql("""select "Credit Limit" from card_limits limit 5""")

# %%

print(con.sql(""" SELECT DISTINCT "Credit Limit" from card_limits """))


# %%
# con.sql(""" DROP table card_limits """)

# %%

print(con.sql("SELECT COUNT(*) FROM card_limits"))

print(con.sql("DESCRIBE card_limits"))

# %% load suppliers list

con.sql(f"""CREATE TABLE suppliers_list AS SELECT * FROM st_read("{suppliersFile}");""")

print(con.sql("SELECT COUNT(*) FROM suppliers_list"))

print(con.sql("DESCRIBE suppliers_list"))

# %%

con.sql("SELECT distinct SUPPLIER_TYPE  from suppliers_list")


# %%

res = con.sql("""SELECT distinct("Transaction.Merchant Name") from pcard_monthly""")

results = res.fetchall()
for r in results:
    print(r)

# %%

# join transactions to department location mapping

con.sql(
    """
CREATE TABLE clean_transactions_1 AS 
SELECT  t.*, m.* 
FROM    pcard_monthly t
LEFT JOIN dept_loc m
    ON  t."Financial Codes.Project Code" = m."Financial Codes.Project Code"
"""
)

# %%

con.sql(
    """
UPDATE clean_transactions_1
    SET "Financial Codes.Project Code" = '0' WHERE "Financial Codes.Project Code" = 'N/A'
"""
)

con.commit()

# %%

con.sql(
    """
UPDATE clean_transactions_1
    SET "Transaction Splits.Split Amount" = '0.00' WHERE "Transaction Splits.Split Amount" IS NULL OR "Transaction Splits.Split Amount" = '' OR "Transaction Splits.Split Amount" = ' '
"""
)

con.commit()

# %%

con.sql(
    """
ALTER TABLE clean_transactions_1 ALTER "Transaction Splits.Split Amount" TYPE DECIMAL;    
"""
)

con.commit()

# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_1 ADD COLUMN ALL_MATCHED_RULES STRING DEFAULT ' '
"""
)

con.commit()

# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_1 ADD COLUMN ANY_MATCHED_RULE BOOLEAN DEFAULT False
"""
)

con.commit()

# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_1 ADD COLUMN AmountIsWholeNumber BOOLEAN DEFAULT False
"""
)

con.commit()

print(r)

# %%

r = con.sql(
    """
UPDATE clean_transactions_1 
    SET AmountIsWholeNumber = (CASE WHEN "Transaction Splits.Split Amount" = floor("Transaction Splits.Split Amount") AND "Transaction Splits.Split Amount" / 50 = floor("Transaction Splits.Split Amount" / 50) THEN True ELSE False END)
"""
)

con.commit()

print(r)


# %%

res = con.sql(
    """SELECT AmountIsWholeNumber, count(*) from clean_transactions_1 
    GROUP BY AmountIsWholeNumber
    ORDER BY count(*);
    """
)

con.commit()

print(res)

# %%

con.sql("""
UPDATE clean_transactions_1 SET ALL_MATCHED_RULES = ALL_MATCHED_RULES || 'AmtWhole# ' WHERE AmountIsWholeNumber = True
""")

con.commit()

# %%

con.sql("""
UPDATE clean_transactions_1 SET ANY_MATCHED_RULE = True WHERE AmountIsWholeNumber = True
""")

con.commit()


# %%

print(con.sql("SELECT COUNT(*) FROM clean_transactions_1"))

print(con.sql("DESCRIBE clean_transactions_1"))

# %%

r = con.sql(
    """
SELECT "Financial Codes.General Ledger Code", CASE WHEN len(string_split("Financial Codes.General Ledger Code", '-')) > 1 THEN string_split("Financial Codes.General Ledger Code", '-')[2] ELSE "Financial Codes.General Ledger Code" END AS test FROM clean_transactions_1;
"""
)

print(r)

# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_1 ADD COLUMN AcctNum STRING DEFAULT 0
"""
)

print(r)

# %%

r = con.sql(
    """
UPDATE clean_transactions_1 
    SET AcctNum = (CASE WHEN len(string_split("Financial Codes.General Ledger Code", '-')) > 1 THEN string_split("Financial Codes.General Ledger Code", '-')[2] ELSE "Financial Codes.General Ledger Code" END)
"""
)

con.commit()

print(r)

# %%

print(con.sql("SELECT COUNT(*) FROM clean_transactions_1"))

print(con.sql("DESCRIBE clean_transactions_1"))

# con.sql("COPY clean_transactions_1 TO outputFile (HEADER, DELIMITER ',');")

# %%

con.sql(
    """
UPDATE clean_transactions_1
    SET "AcctNum" = '0' WHERE "AcctNum" = 'N/A'
"""
)

con.commit()

# %%

con.sql(
    """
UPDATE clean_transactions_1
SET "Transaction User Defined.Note" = lower("Transaction User Defined.Note");
"""
)

con.commit()

# %%

print(con.sql("select * from clean_transactions_1 limit 5"))

# %%

print(con.sql("""select distinct "Company.Company Legal Name" from clean_transactions_1 """))

# %%


def check_flagged(desc: str) -> bool:
    if any(re.search(r"\b" + re.escape(flag) + r"\b", desc) for flag in cleanExceptionsList):
        return True
    else:
        return False


# %%

con.sql("ALTER TABLE clean_transactions_1 ADD COLUMN Flagged BOOLEAN DEFAULT False")

con.commit()

# %%

try:
    con.remove_function("check_flagged")
except Exception:
    pass

con.create_function("check_flagged", check_flagged, [VARCHAR], BOOLEAN)

# %%

con.sql(
    """
UPDATE clean_transactions_1
SET Flagged = check_flagged("Transaction User Defined.Note");
"""
)

con.commit()


# %%

con.commit()

# %%

# check_flagged("table covers for marketing events")

# %%

con.sql("""
UPDATE clean_transactions_1 SET ALL_MATCHED_RULES = ALL_MATCHED_RULES || 'Exception ' WHERE Flagged = True
""")

con.commit()

# %%

con.sql("""
UPDATE clean_transactions_1 SET ANY_MATCHED_RULE = True WHERE Flagged = True
""")

con.commit()

# %%
print("after running the set flag...")
res = con.sql("SELECT Flagged, count(*) from clean_transactions_1 GROUP BY Flagged;")
print(res)

# %%

con.sql(
    """
CREATE TABLE clean_transactions_2 AS 
SELECT  t.*, a.* 
FROM clean_transactions_1 t
LEFT JOIN acct_number_name a
    ON  CAST(t.AcctNum as INTEGER) = a."Account Number"
"""
)

print(con.sql("SELECT COUNT(*) FROM clean_transactions_2"))

# %%

print(con.sql("DESCRIBE clean_transactions_2"))


# %%

con.sql(
    """
CREATE TABLE clean_transactions_3 AS 
SELECT  t.*, cl.* 
FROM clean_transactions_2 t
LEFT JOIN card_limits cl
    ON  t."Cardholder.First Name" = cl."First Name" AND t."Cardholder.Last Name" = cl."Last Name" AND t."Financial Codes.Project Code" = cl."Financial Codes.Project Code"
"""
)

print(con.sql("SELECT COUNT(*) FROM clean_transactions_3"))


# %%

con.sql(f"""COPY clean_transactions_3 TO '{outputFile}' (HEADER, DELIMITER ',');""")

# %%

print(con.sql("DESCRIBE clean_transactions_3"))

# %%

for del_col in [
    "Financial Codes.Cost Center Code:1",
    "Financial Codes.Project Code:1",
    "Company.Company Legal Name:1",
    "Financial Codes.Project Code:2",
    "Location Name:1",
    "Location Category:1",
    "Grouping:1",
]:
    try:
        con.sql(f"""ALTER TABLE clean_transactions_3 DROP "{del_col}"; """)
        con.commit()
    except Exception:
        pass
# %%

print(con.sql("SELECT COUNT(*) FROM clean_transactions_3"))

con.sql("""DELETE FROM clean_transactions_3 WHERE "Transaction.Merchant Name" = 'AUTOPAYMENT - THANK YOU' """)

print(con.sql("SELECT COUNT(*) FROM clean_transactions_3"))

# %%

print(con.sql("SELECT COUNT(*) FROM clean_transactions_3"))

con.sql("""DELETE FROM clean_transactions_3 WHERE "Transaction.Merchant Name" = 'WEX MONTHLY FEE' """)

print(con.sql("SELECT COUNT(*) FROM clean_transactions_3"))


# %%

con.sql("ALTER TABLE clean_transactions_3 ADD COLUMN RULE1_GT4 BOOLEAN DEFAULT False")

con.commit()


# %%
print("before running RULE1...")
res = con.sql("SELECT RULE1_GT4, count(*) from clean_transactions_3 GROUP BY RULE1_GT4;")
print(res)

# %%
# Rule 1 - Greater than 4 transactions in a day

# Wrong results??? Seem fine...

con.sql(
    """
UPDATE clean_transactions_3 as outertable
SET RULE1_GT4 = True
FROM (
    SELECT "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name", COUNT(*) as COUNT 
        FROM clean_transactions_3
        GROUP BY "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name"
        HAVING COUNT >= 4
) AS innertable
WHERE 
    outertable."Transaction.Transaction Dt" = innertable."Transaction.Transaction Dt" 
    AND outertable."Cardholder.First Name" = innertable."Cardholder.First Name"
    AND outertable."Cardholder.Last Name" = innertable."Cardholder.Last Name";
"""
)

con.commit()


# %%

con.sql(
    """
    SELECT "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name", COUNT(*) as COUNT 
        FROM clean_transactions_3
        GROUP BY "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name"
        HAVING COUNT >= 4
        ORDER BY COUNT
"""
)

# %%
con.sql(
    """
SELECT "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name", COUNT(*)
        FROM clean_transactions_3
        WHERE RULE1_GT4 = True
        GROUP BY "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name"
        ORDER BY COUNT(*)
"""
)

# %%

print("after running RULE1...")
res = con.sql("SELECT RULE1_GT4, count(*) from clean_transactions_3 GROUP BY RULE1_GT4;")
print(res)

# %%

con.sql("ALTER TABLE clean_transactions_3 ADD COLUMN RULE2_2SAMEVENDOR BOOLEAN DEFAULT False")

con.commit()

# con.sql("""
# UPDATE clean_transactions_3 SET RULE2_2SAMEVENDOR = False
# """)

# %%

con.sql(
    """
UPDATE clean_transactions_3 as outertable
SET RULE2_2SAMEVENDOR = True
FROM (
    SELECT "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name", "Transaction.Merchant Name", COUNT(*) as COUNT 
        FROM clean_transactions_3
        GROUP BY "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name", "Transaction.Merchant Name"
        HAVING COUNT >= 2 and sum("Transaction Splits.Split Amount") >= 250
) AS innertable
WHERE 
    outertable."Transaction.Transaction Dt" = innertable."Transaction.Transaction Dt" 
    AND outertable."Cardholder.First Name" = innertable."Cardholder.First Name"
    AND outertable."Cardholder.Last Name" = innertable."Cardholder.Last Name"
    AND outertable."Transaction.Merchant Name" = innertable."Transaction.Merchant Name";
"""
)

con.commit()

# %%

con.sql("""
UPDATE clean_transactions_3 SET ALL_MATCHED_RULES = ALL_MATCHED_RULES || '2TSVSD ' WHERE RULE2_2SAMEVENDOR = True
""")

con.commit()

# %%

con.sql("""
UPDATE clean_transactions_3 SET ANY_MATCHED_RULE = True WHERE RULE2_2SAMEVENDOR = True
""")

con.commit()

# %%

print("after running RULE2...")
res = con.sql("SELECT RULE2_2SAMEVENDOR, count(*) from clean_transactions_3 GROUP BY RULE2_2SAMEVENDOR;")
print(res)

# %%

con.sql("ALTER TABLE clean_transactions_3 ADD COLUMN RULE4_SAMEDOLLAR BOOLEAN DEFAULT False")

con.commit()

# %%

con.sql(
    """
UPDATE clean_transactions_3 as outertable
SET RULE4_SAMEDOLLAR = True
FROM (
    SELECT "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name", "Transaction.Merchant Name", "Transaction Splits.Split Amount", COUNT(*) as COUNT 
        FROM clean_transactions_3
        GROUP BY "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name", "Transaction.Merchant Name", "Transaction Splits.Split Amount"
        HAVING COUNT >= 2
) AS innertable
WHERE 
    outertable."Transaction.Transaction Dt" = innertable."Transaction.Transaction Dt" 
    AND outertable."Cardholder.First Name" = innertable."Cardholder.First Name"
    AND outertable."Cardholder.Last Name" = innertable."Cardholder.Last Name"
    AND outertable."Transaction.Merchant Name" = innertable."Transaction.Merchant Name"
    AND outertable."Transaction Splits.Split Amount" = innertable."Transaction Splits.Split Amount";
"""
)

con.commit()

# %%

con.sql("""
UPDATE clean_transactions_3 SET ALL_MATCHED_RULES = ALL_MATCHED_RULES || '2TSVSDS$ ' WHERE RULE4_SAMEDOLLAR = True
""")

con.commit()

# %%

con.sql("""
UPDATE clean_transactions_3 SET ANY_MATCHED_RULE = True WHERE RULE4_SAMEDOLLAR = True
""")

con.commit()

# %%

print("after running RULE4...")
res = con.sql("SELECT RULE4_SAMEDOLLAR, count(*) from clean_transactions_3 GROUP BY RULE4_SAMEDOLLAR;")
print(res)

# %%

# load card and grouping from pivot mapping

con.sql(f"""CREATE TABLE pivot_mapping AS SELECT * FROM st_read('{pivotMappingFile}', layer = 'Final Mapping', open_options=['HEADERS=FORCE']);""")

# f"""CREATE TABLE pivot_mapping AS SELECT * FROM st_read('{pivotMappingFile}', layer = 'Pivot-Mapping Example', open_options=['HEADERS=FORCE']);"""

print(con.sql("SELECT COUNT(*) FROM pivot_mapping"))


# %%
# New logic for mapping card names

# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_3 ADD COLUMN FN_LN_1 STRING DEFAULT ' '
"""
)

con.commit()

print(r)

# %%

r = con.sql("""
 UPDATE clean_transactions_3 SET FN_LN_1 = "Cardholder.First Name" || ' ' || "Cardholder.Last Name"
""")

con.commit()

print(r)

# %%

r = con.sql("""
 UPDATE clean_transactions_3 SET FN_LN_1 = replace(FN_LN_1, 'N/A ', '')
""")

con.commit()

print(r)

# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_3 ADD COLUMN MAP_COUNT_TO_CL INTEGER DEFAULT -1
"""
)

con.commit()

# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_3 ADD COLUMN MAP_COUNT_TO_PM INTEGER DEFAULT -1
"""
)

con.commit()

# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_3 ADD COLUMN MAP_COUNT_TO_CL_NEW INTEGER DEFAULT -1
"""
)

con.commit()

# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_3 ADD COLUMN MAP_COUNT_TO_NEW_PIVOT_FN_LN INTEGER DEFAULT -1
"""
)

con.commit()

# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_3 ADD COLUMN UPDATED_MAPPING STRING DEFAULT ' '
"""
)

con.commit()

# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_3 ADD COLUMN FINAL_CARD_NAME STRING DEFAULT ' '
"""
)

con.commit()

# %%

r = con.sql("""
UPDATE clean_transactions_3 SET MAP_COUNT_TO_CL = (SELECT COUNT(*) FROM card_limits cl WHERE "Company Unit.Hierarchy Name" = cl."Hierarchy Node Name" AND FN_LN_1 = replace(cl."First Name" || ' ' || cl."Last Name", 'N/A ', ''))
""")

con.commit()

# %%

r = con.sql("""
UPDATE clean_transactions_3 SET MAP_COUNT_TO_PM = (SELECT COUNT(*) FROM pivot_mapping pm WHERE "Company Unit.Hierarchy Name" = pm."Hierarchy Name" AND "Cardholder.First Name" = pm."First Name" AND "Cardholder.Last Name" = pm."Last Name")
""")

con.commit()

# %%

r = con.sql("""
UPDATE clean_transactions_3 SET MAP_COUNT_TO_NEW_PIVOT_FN_LN = (SELECT COUNT(*) FROM pivot_mapping pm WHERE FN_LN_1 = pm."First name + Last Name")
""")

con.commit()

# %%

r = con.sql("""
 UPDATE clean_transactions_3 SET FINAL_CARD_NAME = (SELECT pm."Final Card Name Mapping" FROM pivot_mapping pm WHERE FN_LN_1 = pm."First name + Last Name") where MAP_COUNT_TO_NEW_PIVOT_FN_LN = 1
""")

con.commit()

# %%

r = con.sql("""
 UPDATE clean_transactions_3 SET UPDATED_MAPPING = (SELECT pm."Assigned Grouping" FROM pivot_mapping pm WHERE FN_LN_1 = pm."First name + Last Name") where MAP_COUNT_TO_NEW_PIVOT_FN_LN = 1
""")

con.commit()

# %%

r = con.sql("""
UPDATE clean_transactions_3 SET MAP_COUNT_TO_NEW_PIVOT_FN_LN = 11 where MAP_COUNT_TO_NEW_PIVOT_FN_LN = 1
""")

con.commit()

# %%

r = con.sql("""
UPDATE clean_transactions_3 SET MAP_COUNT_TO_NEW_PIVOT_FN_LN = (SELECT COUNT(*) FROM pivot_mapping pm WHERE "Company Unit.Hierarchy Name" = pm."Hierarchy Name" AND FN_LN_1 = pm."First name + Last Name") WHERE (MAP_COUNT_TO_NEW_PIVOT_FN_LN > 1 and MAP_COUNT_TO_NEW_PIVOT_FN_LN != 11)
""")

con.commit()

# %%

r = con.sql("""
 UPDATE clean_transactions_3 SET FINAL_CARD_NAME = (SELECT pm."Final Card Name Mapping" FROM pivot_mapping pm WHERE "Company Unit.Hierarchy Name" = pm."Hierarchy Name" AND FN_LN_1 = pm."First name + Last Name") where MAP_COUNT_TO_NEW_PIVOT_FN_LN = 1
""")

con.commit()

# %%

r = con.sql("""
 UPDATE clean_transactions_3 SET UPDATED_MAPPING = (SELECT pm."Assigned Grouping" FROM pivot_mapping pm WHERE "Company Unit.Hierarchy Name" = pm."Hierarchy Name" AND FN_LN_1 = pm."First name + Last Name") where MAP_COUNT_TO_NEW_PIVOT_FN_LN = 1
""")

con.commit()

# %%

r = con.sql("""
UPDATE clean_transactions_3 SET MAP_COUNT_TO_NEW_PIVOT_FN_LN = 11 where MAP_COUNT_TO_NEW_PIVOT_FN_LN = 1
""")

con.commit()

# %%

r = con.sql("""
UPDATE clean_transactions_3 SET MAP_COUNT_TO_CL_NEW = (SELECT COUNT(*) FROM card_limits cl WHERE FN_LN_1 = cl."Hierarchy Node Name")
""")

con.commit()

# %%

print(con.sql("select * from clean_transactions_3 limit 5"))


# %%

# r = con.sql(
#     """
# ALTER TABLE clean_transactions_3 ADD COLUMN UPDATED_MAPPING STRING DEFAULT ' '
# """
# )

# con.commit()

# # %%

# r = con.sql("""
#  UPDATE clean_transactions_3 SET UPDATED_MAPPING = (SELECT pm."Final Card Name Mapping" FROM pivot_mapping pm WHERE "Company Unit.Hierarchy Name" = pm."Hierarchy Name" AND "Cardholder.First Name" = pm."First Name" AND "Cardholder.Last Name" = pm."Last Name")
# """)

# #  UPDATE clean_transactions_3 SET UPDATED_MAPPING = (SELECT pm."Assigned Grouping based off of Hierarchy name to mapping table" FROM pivot_mapping pm WHERE "Company Unit.Hierarchy Name" = pm."Hierarchy Name" AND "Cardholder.First Name" = pm."First Name" AND "Cardholder.Last Name" = pm."Last Name")

# con.commit()

# # %%

# r = con.sql(
#     """
# ALTER TABLE clean_transactions_3 ADD COLUMN FINAL_CARD_NAME STRING DEFAULT ' '
# """
# )

# con.commit()

# # %%

# r = con.sql("""
#  UPDATE clean_transactions_3 SET FINAL_CARD_NAME = (SELECT pm."Final Card Name Mapping" FROM pivot_mapping pm WHERE "Company Unit.Hierarchy Name" = pm."Hierarchy Name" AND "Cardholder.First Name" = pm."First Name" AND "Cardholder.Last Name" = pm."Last Name")
# """)

# con.commit()


# %%

print(con.sql("SELECT * FROM clean_transactions_3 LIMIT 5;"))


# %%

con.sql(f"""COPY clean_transactions_3 TO '{outputFile}' (HEADER, DELIMITER ',');""")

# %%

out_df = con.sql("SELECT * from clean_transactions_3").df()

# %%
# CLEAN COMPANY NAMES


def clean_company_name(cn):
    return re.sub(r"([^a-zA-Z ]+?)| US | PO | INC|DD DOORDASH | Inc|DOORDASH |EZCATER|PAYPAL |SQ |TLF |TST |WWW|SP |TM ", "", cn)


translations = {
    r"AMZN Mktp.*": r"AMAZON",
    r"MCDONALDS.*": r"MCDONALDS",
    r"WM SUPERCENTER.*": r"WALMART",
    r"KFC .*": r"KFC",
    r"DROPBOX .*": r"DROPBOX",
    r"AMAZONCOM .*": r"AMAZON",
    r"Amazoncom .*": r"AMAZON",
    r"Amazon Prime .*": r"AMAZON PRIME",
    r"AMZN MKTP.*": r"AMAZON",
    r"COSTCO.*": r"COSTCO",
    r"Amazon Music.*": r"AMAZON MUSIC",
    r"Amazoncom.*": r"AMAZON",
    r"HARDEES.*": r"HARDEES",
    r"KING BUFFET.*": r"KING BUFFET",
    r"RBRT .*": r"RBRT",
    r"ACE HARDWARE.*": r"ACE HARDWARE",
    r"ACE HDWE.*": r"ACE HARDWARE",
    r"ADOBE.*": r"ADOBE",
    r"AGAVE  RYE.*": r"AGAVE  RYE",
    r"Aging Connections.*": r"Aging Connections",
    r"AIRBNB.*": r"AIRBNB",
    r"AMC .*": r"AMC",
    r"APPLEBEES.*": r"APPLEBEES",
    r"APPLECOM.*": r"APPLE",
    r"ARBYS.*": r"ARBYS",
    r"Audible.*": r"AUDIBLE",
    r"BEST BUY.*": r"BEST BUY",
    r"BEST WESTERN.*": r"BEST WESTERN",
    r"BIG BOY.*": r"BIG BOY",
    r"BIGBOY.*": r"BIG BOY",
    r"BIG LOTS.*": r"BIG LOTS",
    r"BUFFALO WILD.*": r"BUFFALO WILD WINGS",
    r"BWW .*": r"BUFFALO WILD WINGS",
    r"BURGER KING.*": r"BURGER KING",
    r"CRACKER BARREL.*": r"CRACKER BARREL",
    r"CRUMBL.*": r"CRUMBL",
    r"CULVER.*": r"CULVERS",
    r"DAIRY QUEEN.*": r"DAIRY QUEEN",
    r"DELTA.*": r"DELTA",
    r"Etsycom.*": r"ETSY",
    r"FACEBK.*": r"FACEBOOK",
    r"FAIRFIELD INN.*": r"FAIRFIELD INN",
    r"FEDEX.*": r"FEDEX",
    r"FIRST WATCH.*": r"FIRST WATCH",
    r"FIVE BELOW.*": r"FIVE BELOW",
    r"FIVE GUYS.*": r"FIVE GUYS",
    r"GOODWILL.*": r"GOODWILL",
    r"HAMPTON INN.*": r"HAMPTON INN",
    r"HILTON GARDEN INN.*": r"HILTON GARDEN INN",
    r"HOLIDAY INN.*": r"HOLIDAY INN",
    r"HOOTERS.*": r"HOOTERS",
    r"JIMMY JOHNS.*": r"JIMMY JOHNS",
    r"KROGER.*": r"KROGER",
    r"LAROSAS.*": r"LAROSAS",
    r"MARATHON PETRO.*": r"MARATHON PETRO",
    r"MCALISTERS.*": r"MCALISTERS",
    r"MEIJER.*": r"MEIJER",
    r"MENARDS.*": r"MENARDS",
    r"MICROSOFT.*": r"MICROSOFT",
    r"Microsoft.*": r"MICROSOFT",
    r"MSFT .*": r"MICROSOFT",
    r"OCHARLEYS.*": r"OCHARLEYS",
    r"OFFICEMAX.*": r"OFFICE MAX",
    r"PANDA EXPRESS.*": r"PANDA EXPRESS",
    r"PANERA BREAD.*": r"PANERA BREAD",
    r"PAPA MURPHYS.*": r"PAPA MURPHYS",
    r"PENN STATION.*": r"PENN STATION",
    r"PF CHANGS.*": r"PF CHANGS",
    r"PIZZA KING.*": r"PIZZA KING",
    r"PIZZAKING.*": r"PIZZA KING",
    r"PIZZAHUT.*": r"PIZZA HUT",
    r"Prime Video.*": r"PRIME VIDEO",
    r"RESIDENCE INN.*": r"RESIDENCE INN",
    r"SAMS CLUB.*": r"SAMS CLUB",
    r"SAMSCLUB.*": r"SAMS CLUB",
    r"SKYLINE CHILI.*": r"SKYLINE CHILI",
    r"SPEEDWAY.*": r"SPEEDWAY",
    r"STARBUCKS.*": r"STARBUCKS",
    r"UPS .*": r"UPS",
    r"WALMART.*": r"WALMART",
    r"WINGS ETC.*": r"WINGS ETC",
}

# %%

out_df["CleanMerchantName"] = out_df["Transaction.Merchant Name"].apply(lambda x: clean_company_name(x))

# %%
out_df["CleanMerchantName"].replace(translations, regex=True, inplace=True)

# %%
out_df["CleanMerchantName"] = out_df["CleanMerchantName"].str.strip()


# %%

out_df.head()

# %%

out_df.to_excel(outputExcelFile)

# %%
# Merge Files

# import pandas as pd
# import os

# files_list = [
#     './files/pcard_monthly_output_full_history-new1.xlsx',
#     './files/pcard_monthly_output_jan24-new1.xlsx',
#     './files/pcard_monthly_output_feb24-new1.xlsx',
#     './files/pcard_monthly_output_mar24.xlsx',
#     './files/pcard_monthly_output_apr24.xlsx',
#     './files/pcard_monthly_output_may24.xlsx',
#     './files/pcard_monthly_output_june24.xlsx',
#     './files/pcard_monthly_output_july24.xlsx',
#     './files/pcard_monthly_output_aug24.xlsx',
#     './files/pcard_monthly_output_sep24.xlsx',
#     './files/pcard_monthly_output_oct24.xlsx',
#     './files/pcard_monthly_output_nov24.xlsx',
#     './files/pcard_monthly_output_dec24.xlsx',
#     './files/pcard_monthly_output_jan25.xlsx',
# ]

# df_list = []

# for f in files_list:
#     data = pd.read_excel(f)
#     df_list.append(data)

# df = pd.concat(df_list)

# df.to_excel("./files/pcard_monthly_output_full_new_jan25.xlsx")

# %%
