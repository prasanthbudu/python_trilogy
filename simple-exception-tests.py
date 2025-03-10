# %%
import re


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
desc = "sams club purchase for careshelf - february.  please code to 61231."
print(any(re.search(r"\b" + re.escape(flag) + r"\b", desc) for flag in cleanExceptionsList))
# %%
for flag in cleanExceptionsList:
    if flag in desc:
        print("true ", flag)
# %%
