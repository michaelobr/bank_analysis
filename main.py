# Pandas is a great python module for organizing excel data. It's super fast, lightweight, and has amazing online documentation. The "as pd" is the standard nomenclature.
import pandas as pd 
import time

# I like to always assign the filepath to a variable because it's much easier to simply reuse the variable rather than the full path every time
filepath = "data.csv"

# This will read in the data from the sheet into a pandas dataframe object, which will allow for easy manipulation and cleaning of the data
## header = [8] tells the dataframe to begin at the 9th row (0 indexed hence we use 8), and to use that row as the header names of our columns
df = pd.read_csv(filepath, header=[8], parse_dates=["Order Date", "Ship Date"])

# Remove all rows that don't have a value in the OrderID column. This will remove "This header is for testing purposes" as well as empty rows between tables, and repeated table headers
df = df.dropna(subset = ["OrderID"])

# Set all numeric values within the dataframe to their corresponding dtype, ignoring dates
df = df.apply(pd.to_numeric, errors="ignore") 

# This code will ensure each OrderID is 6 digits in length with leading zeroes 
df["OrderID"] = df["OrderID"].astype("int") # the one downside to the one liner in line 16 is that it interprets the OrderIDs as floats, so here we have to recast them as integers
df["OrderID"] = df["OrderID"].astype("str")
df['OrderID'] = df['OrderID'].str.zfill(6)

# Uncomment the two below lines to output .csv and .xlsx outputs of cleaned data (this is before we start top priority)
#df.to_csv("target.csv", index=False)
#df.to_excel("target.xlsx", index=False) #This will keep formatting (i.e. cell P3 "Kleencut®" vs KleencutÂ® in csv format)

# Create a new dataframe called Top Priority, we'll use our previous dataframe as the base. Also remove columns labeled "Product Container" and "Quota Margins"
df_top_priority = df.drop(["Product Container", "Quota Margins"], axis=1)

# We only want order priority with High or Critical now
df_top_priority = df_top_priority[(df_top_priority["Order Priority"] == "High") | (df_top_priority["Order Priority"] == "Critical")]

# Sort by Order Priority (descending), Order Date (descending), and lastly Ship Date (descending)
df_top_priority = df_top_priority.sort_values(by=["Order Priority", "Order Date", "Ship Date"], ascending=False)

# Write to .csv and .xlsx outputs
df_top_priority.to_csv("top_pri.csv", index=False)
df_top_priority.to_excel("top_pri.xlsx", index=False) # Again, xlsx keeps encoding format much better than csv

print(len(list(df_top_priority["OrderID"])), "recrods were imported to high priority.")