# %%
import duckdb
import re

# %%
transactions_file = "./files/pcard_monthly_output_full_history-new1.csv"

# %%
con = duckdb.connect()

con.sql("INSTALL spatial;")
con.sql("LOAD spatial;")

# %%
con.sql(
    f"""CREATE TABLE full_transactions_file AS SELECT * FROM st_read("{transactions_file}", open_options=['HEADERS=FORCE']);"""
)

# %%
company_names = con.sql(
    """ SELECT DISTINCT "Transaction.Merchant Name" from full_transactions_file """
).to_df()

# %%
print(len(company_names))
# %%
company_names.head(50)


# %%
def clean_company_name(cn):
    # return re.findall(
    #     "[\s*<]+\d+$|[\s*<]+(?![A-Z]{6}.*)\w*\d[\w>]*$|\d{6,}$|[\s*<]+[A-Z]{6}$|(?![A-Z]+$)(?<=[A-Z])\w{6}$",
    #     cn,
    # )
    return re.sub (r'([^a-zA-Z ]+?)| US | PO | INC|DD DOORDASH | Inc|DOORDASH |EZCATER|PAYPAL |SQ |TLF |TST |WWW|SP |TM ', '', cn)



# %%
company_names['CleanMerchantName'] = company_names['Transaction.Merchant Name'].apply(lambda x: clean_company_name(x))

# %%
company_names.head(50)
# %%

translations = {
    r'AMZN Mktp.*': r'AMAZON', 
    r'MCDONALDS.*': r'MCDONALDS', 
    r'WM SUPERCENTER.*': r'WALMART', 
    r'KFC .*': r'KFC', 
    r'DROPBOX .*': r'DROPBOX', 
    r'AMAZONCOM .*': r'AMAZON', 
    r'Amazoncom .*': r'AMAZON',
    r'Amazon Prime .*': r'AMAZON PRIME',
    r'AMZN MKTP.*': r'AMAZON',
    r'COSTCO.*': r'COSTCO',
    r'Amazon Music.*': r'AMAZON MUSIC',
    r'Amazoncom.*': r'AMAZON',
    r'HARDEES.*': r'HARDEES',
    r'KING BUFFET.*': r'KING BUFFET',
    r'RBRT .*': r'RBRT',
    r'ACE HARDWARE.*': r'ACE HARDWARE',
    r'ACE HDWE.*': r'ACE HARDWARE',
    r'ADOBE.*': r'ADOBE',
    r'AGAVE  RYE.*': r'AGAVE  RYE',
    r'Aging Connections.*': r'Aging Connections',
    r'AIRBNB.*': r'AIRBNB',
    r'AMC .*': r'AMC',
    r'APPLEBEES.*': r'APPLEBEES',
    r'APPLECOM.*': r'APPLE',
    r'ARBYS.*': r'ARBYS',
    r'Audible.*': r'AUDIBLE',
    r'BEST BUY.*': r'BEST BUY',
    r'BEST WESTERN.*': r'BEST WESTERN',
    r'BIG BOY.*': r'BIG BOY',
    r'BIGBOY.*': r'BIG BOY',
    r'BIG LOTS.*': r'BIG LOTS',
    r'BUFFALO WILD.*': r'BUFFALO WILD WINGS',
    r'BWW .*': r'BUFFALO WILD WINGS',
    r'BURGER KING.*': r'BURGER KING',
    r'CRACKER BARREL.*': r'CRACKER BARREL',
    r'CRUMBL.*': r'CRUMBL',
    r'CULVER.*': r'CULVERS',
    r'DAIRY QUEEN.*': r'DAIRY QUEEN',
    r'DELTA.*': r'DELTA',
    r'Etsycom.*': r'ETSY',
    r'FACEBK.*': r'FACEBOOK',
    r'FAIRFIELD INN.*': r'FAIRFIELD INN',
    r'FEDEX.*': r'FEDEX',
    r'FIRST WATCH.*': r'FIRST WATCH',
    r'FIVE BELOW.*': r'FIVE BELOW',
    r'FIVE GUYS.*': r'FIVE GUYS',
    r'GOODWILL.*': r'GOODWILL',
    r'HAMPTON INN.*': r'HAMPTON INN',
    r'HILTON GARDEN INN.*': r'HILTON GARDEN INN',
    r'HOLIDAY INN.*': r'HOLIDAY INN',
    r'HOOTERS.*': r'HOOTERS',
    r'JIMMY JOHNS.*': r'JIMMY JOHNS',
    r'KROGER.*': r'KROGER',
    r'LAROSAS.*': r'LAROSAS',
    r'MARATHON PETRO.*': r'MARATHON PETRO',
    r'MCALISTERS.*': r'MCALISTERS',
    r'MEIJER.*': r'MEIJER',
    r'MENARDS.*': r'MENARDS',
    r'MICROSOFT.*': r'MICROSOFT',
    r'Microsoft.*': r'MICROSOFT',
    r'MSFT .*': r'MICROSOFT',
    r'OCHARLEYS.*': r'OCHARLEYS',
    r'OFFICEMAX.*': r'OFFICE MAX',
    r'PANDA EXPRESS.*': r'PANDA EXPRESS',
    r'PANERA BREAD.*': r'PANERA BREAD',
    r'PAPA MURPHYS.*': r'PAPA MURPHYS',
    r'PENN STATION.*': r'PENN STATION',
    r'PF CHANGS.*': r'PF CHANGS',
    r'PIZZA KING.*': r'PIZZA KING',
    r'PIZZAKING.*': r'PIZZA KING',
    r'PIZZAHUT.*': r'PIZZA HUT',
    r'Prime Video.*': r'PRIME VIDEO',
    r'RESIDENCE INN.*': r'RESIDENCE INN',
    r'SAMS CLUB.*': r'SAMS CLUB',
    r'SAMSCLUB.*': r'SAMS CLUB',
    r'SKYLINE CHILI.*': r'SKYLINE CHILI',
    r'SPEEDWAY.*': r'SPEEDWAY',
    r'STARBUCKS.*': r'STARBUCKS',
    r'UPS .*': r'UPS',
    r'WALMART.*': r'WALMART',
    r'WINGS ETC.*': r'WINGS ETC',
}

# %%
company_names['CleanMerchantName'].replace(translations, regex=True, inplace=True)

# %%
company_names['CleanMerchantName'] = company_names['CleanMerchantName'].str.strip()

# %%
company_names.head(50)

# %%
company_names.to_csv("./files/clean_company_names.csv")
# %%
