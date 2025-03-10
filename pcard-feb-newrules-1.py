# %%

import duckdb
from duckdb.typing import BOOLEAN, VARCHAR

# %%

outputFile = "./files/pcard_monthly_output_jan24.csv"
monthlyTransactionsFile = "./files/1.31.24-Monthly_w_GLInternal_Audit_Record.xlsx"
cardLimitsMappingsFile = "./files/Pcard Mapping-updated 2.2.24.xlsx"

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

# %%

res = con.sql("""SELECT distinct("Transaction.Merchant Name") from pcard_monthly""")

results = res.fetchall()
for r in results:
    print(r)

# %%
# load card limits mapping

con.sql(
    f"""CREATE TABLE card_limits_mapping AS SELECT * FROM st_read("{monthlyTransactionsFile}", layer = 'Pivot-Mapping Example', open_options=['HEADERS=FORCE']);"""
)

# %%

print(con.sql("SELECT COUNT(*) FROM card_limits_mapping"))

print(con.sql("DESCRIBE card_limits_mapping"))



# %%

con.sql(
    """
CREATE TABLE clean_transactions_1 AS 
SELECT  t.*, clm.* 
FROM pcard_monthly t
LEFT JOIN card_limits_mapping clm
    ON  t."Cardholder.First Name" = clm."First Name" AND 
    t."Cardholder.Last Name" = clm."Last Name" AND
    t."Company Unit.Hierarchy Name" = clm."Hierarchy Name"
"""
)

# %%

print(con.sql("SELECT COUNT(*) FROM clean_transactions_1"))

# %%
out_df = con.sql("""
    SELECT * from clean_transactions_1;
""")
# %%
out_df.to_csv("./files/test_out.csv")

# %%
