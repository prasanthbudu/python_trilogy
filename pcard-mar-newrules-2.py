# %%

import re
import duckdb
from duckdb.typing import BOOLEAN, VARCHAR

# %%

# %%

suppliersFile = "./files/oracle suppliers for prama.csv"
outputFile = "./files/pcard_monthly_output_feb24.csv"
outputExcelFile = "./files/pcard_monthly_output_feb24.xlsx"
monthlyTransactionsFile = "./files/InternalAuditPcardPreviousMonth_MonthlywGLInternalAudit_20240305064658.xlsx"
cardLimitsMappingsFile = "./files/Go Prama Master- PCard-Mappings-Consolidated.xlsx"
otherMappingsFile = "./files/Go Prama Master- PCard-Mappings-Consolidated.xlsx"
pivotMappingFile = "./files/Go Prama Master- PCard-Mappings-Consolidated.xlsx"
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
    f"""CREATE TABLE pcard_monthly AS SELECT * FROM st_read('{monthlyTransactionsFile}', open_options=['HEADERS=FORCE']);"""
)

print(con.sql("SELECT COUNT(*) FROM pcard_monthly"))

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
    con.sql(
        "ALTER TABLE pcard_monthly RENAME COLUMN "
        + f"'{c[0]}'"
        + " TO "
        + f"'{c[1]}'"
        + ";"
    )

# %%

print(con.sql("DESCRIBE pcard_monthly"))

# ##

print(con.sql(""" SELECT DISTINCT "Company.Company Legal Name" from pcard_monthly """))

# %%

print(
    con.sql(
        """ SELECT DISTINCT "Financial Codes.Cost Center Code" from pcard_monthly """
    )
)


# %%

# load department to location mapping

con.sql(
    f"""CREATE TABLE dept_loc AS SELECT * FROM st_read("{otherMappingsFile}", layer = 'Mapping -Dept-Loc', open_options=['HEADERS=FORCE']);"""
)

print(con.sql("SELECT COUNT(*) FROM dept_loc"))

print(con.sql("DESCRIBE dept_loc"))

# %%

# load account number to name mapping

con.sql(
    f"""CREATE TABLE acct_number_name AS SELECT * FROM st_read("{otherMappingsFile}", layer = 'Mapping - Acct #', open_options=['HEADERS=FORCE']);"""
)

print(con.sql("SELECT COUNT(*) FROM acct_number_name"))

print(con.sql("DESCRIBE acct_number_name"))

# %%

# load card limits

con.sql(
    f"""CREATE TABLE card_limits AS SELECT * FROM st_read("{cardLimitsMappingsFile}", layer = 'Mapping - Card Limits-Loc-Categ', open_options=['HEADERS=FORCE']);"""
)

# %%

print(con.sql("SELECT COUNT(*) FROM card_limits"))

print(con.sql("DESCRIBE card_limits"))

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

print(
    con.sql(
        """select distinct "Company.Company Legal Name" from clean_transactions_1 """
    )
)

# %%


def check_flagged(desc: str) -> bool:
    if any(
        re.search(r"\b" + re.escape(flag) + r"\b", desc) for flag in cleanExceptionsList
    ):
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

con.sql(
    """DELETE FROM clean_transactions_3 WHERE "Transaction.Merchant Name" = 'AUTOPAYMENT - THANK YOU' """
)

print(con.sql("SELECT COUNT(*) FROM clean_transactions_3"))


# %%

con.sql(f"""COPY clean_transactions_3 TO '{outputFile}' (HEADER, DELIMITER ',');""")

# %%

# r = con.sql(
#     """
# ALTER TABLE clean_transactions_3 ADD COLUMN ALL_MATCHED_RULES STRING DEFAULT ' '
# """
# )

# con.commit()

# %%

# r = con.sql(
#     """
# ALTER TABLE clean_transactions_3 ADD COLUMN ANY_MATCHED_RULE BOOLEAN DEFAULT False
# """
# )

# con.commit()

# %%

con.sql("ALTER TABLE clean_transactions_3 ADD COLUMN RULE1_GT4 BOOLEAN DEFAULT False")

con.commit()


# %%
print("before running RULE1...")
res = con.sql(
    "SELECT RULE1_GT4, count(*) from clean_transactions_3 GROUP BY RULE1_GT4;"
)
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
        HAVING COUNT > 4
) AS innertable
WHERE 
    outertable."Transaction.Transaction Dt" = innertable."Transaction.Transaction Dt" 
    AND outertable."Cardholder.First Name" = innertable."Cardholder.First Name"
    AND outertable."Cardholder.Last Name" = innertable."Cardholder.Last Name";
"""
)

con.commit()

# %%

# con.sql("""
# UPDATE clean_transactions_3 SET ALL_MATCHED_RULES = ALL_MATCHED_RULES || 'RULE1 ' WHERE RULE1_GT4 = True
# """)

# con.commit()

# %%

# con.sql("""
# UPDATE clean_transactions_3 SET ANY_MATCHED_RULE = True WHERE RULE1_GT4 = True
# """)

# con.commit()

# %%

con.sql(
    """
    SELECT "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name", COUNT(*) as COUNT 
        FROM clean_transactions_3
        GROUP BY "Transaction.Transaction Dt", "Cardholder.First Name", "Cardholder.Last Name"
        HAVING COUNT > 4
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
res = con.sql(
    "SELECT RULE1_GT4, count(*) from clean_transactions_3 GROUP BY RULE1_GT4;"
)
print(res)

# %%

con.sql(
    "ALTER TABLE clean_transactions_3 ADD COLUMN RULE2_2SAMEVENDOR BOOLEAN DEFAULT False"
)

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
res = con.sql(
    "SELECT RULE2_2SAMEVENDOR, count(*) from clean_transactions_3 GROUP BY RULE2_2SAMEVENDOR;"
)
print(res)

# %%

con.sql(
    "ALTER TABLE clean_transactions_3 ADD COLUMN RULE4_SAMEDOLLAR BOOLEAN DEFAULT False"
)

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
        HAVING COUNT > 2
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
res = con.sql(
    "SELECT RULE4_SAMEDOLLAR, count(*) from clean_transactions_3 GROUP BY RULE4_SAMEDOLLAR;"
)
print(res)

# %%

# load card and grouping from pivot mapping

con.sql(
    f"""CREATE TABLE pivot_mapping AS SELECT * FROM st_read('{pivotMappingFile}', layer = 'Pivot-Mapping Example', open_options=['HEADERS=FORCE']);"""
)

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
UPDATE clean_transactions_3 SET MAP_COUNT_TO_CL_NEW = (SELECT COUNT(*) FROM card_limits cl WHERE FN_LN_1 = cl."Hierarchy Node Name")
""")

con.commit()

# %%

print(con.sql("select * from clean_transactions_3 limit 5"))


# %%

r = con.sql(
    """
ALTER TABLE clean_transactions_3 ADD COLUMN UPDATED_MAPPING STRING DEFAULT ' '
"""
)

con.commit()

# %%

r = con.sql("""
 UPDATE clean_transactions_3 SET UPDATED_MAPPING = (SELECT pm."Assigned Grouping based off of Hierarchy name to mapping table" FROM pivot_mapping pm WHERE "Company Unit.Hierarchy Name" = pm."Hierarchy Name" AND "Cardholder.First Name" = pm."First Name" AND "Cardholder.Last Name" = pm."Last Name")
""")

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
 UPDATE clean_transactions_3 SET FINAL_CARD_NAME = (SELECT pm."Final Card Name Mapping" FROM pivot_mapping pm WHERE "Company Unit.Hierarchy Name" = pm."Hierarchy Name" AND "Cardholder.First Name" = pm."First Name" AND "Cardholder.Last Name" = pm."Last Name")
""")

con.commit()


# %%

print(con.sql("SELECT * FROM clean_transactions_3 LIMIT 5;"))


# %%

con.sql(f"""COPY clean_transactions_3 TO '{outputFile}' (HEADER, DELIMITER ',');""")

# %%

out_df = con.sql("SELECT * from clean_transactions_3").df()

# %%

out_df.head()

# %%

out_df.to_excel(outputExcelFile)

# %%
