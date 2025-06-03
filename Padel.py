# -*- coding: utf-8 -*-
"""
Created on Sun Feb 23 20:12:13 2025

@author: JColc
"""

import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

credentials_path = r"C:\Users\JColc\Downloads\our-rock-451820-r5-919f8b93adff.json"

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
client = gspread.authorize(creds)
#%%

# Open the Google Sheet
sheet = client.open("Padel Tracker").worksheet('Scores')

# Get all data as a list of lists
data = sheet.get_all_values()

# Convert to Pandas DataFrame
df = pd.DataFrame(data)

array = df.to_numpy()
#%%
a=[]
for i in range(len(array)):
    if array[i,1] == "":
        break
    else:
        a.append(array[i,:])   


JW =0
JL =0
RW =0
RL =0
GW=0
GL=0
PW=0
PL=0
TW=0
TL=0
DW=0
DL=0
JGW=0
JGL=0
GGW=0
GGL=0
RGW=0
RGL=0
PGW=0
PGL=0
TGW=0
TGL=0
DGW=0
DGL=0
JP=0
GP=0
TP=0
PP=0
RP=0
DP=0

c= [JGW,JGL,GGW,GGL,TGW,TGL,PGW,PGL,RGW,RGL,DGW,DGL]

for i in range(len(a)-1):
    name=["Jamie"]
    s=a[i+1][1:3] 
    res = any(elem in s for elem in name)
    if res == True:
        JP +=1
        JGW += int(a[i+1][5])
        JGL += int(a[i+1][6])
        if a[i+1][5]>a[i+1][6]:
            JW += 1
        else:
            JL += 1
    s=a[i+1][3:5] 
    res = any(elem in s for elem in name) 
    if res == True:
        JP +=1
        JGW += int(a[i+1][6])
        JGL += int(a[i+1][5])
        if a[i+1][6]>a[i+1][5]:
            JW += 1
        else:
            JL += 1
    name=["George"]
    s=a[i+1][1:3] 
    res = any(elem in s for elem in name)
    if res == True:
        GP +=1
        GGW+= int(a[i+1][5])
        GGL+= int(a[i+1][6])
        if a[i+1][5]>a[i+1][6]:
            GW +=1
        else:
            GL +=1
    s=a[i+1][3:5] 
    res = any(elem in s for elem in name) 
    if res == True:
        GP +=1
        GGW+= int(a[i+1][6])
        GGL+= int(a[i+1][5])
        if a[i+1][6]>a[i+1][5]:
            GW +=1
        else:
            GL +=1
    name=["Paul"]
    s=a[i+1][1:3] 
    res = any(elem in s for elem in name)
    if res == True:
        PP +=1
        PGW+= int(a[i+1][5])
        PGL+= int(a[i+1][6])
        if a[i+1][5]>a[i+1][6]:
            PW +=1
        else:
            PL +=1
    s=a[i+1][3:5] 
    res = any(elem in s for elem in name) 
    if res == True:
        PP +=1
        PGW+= int(a[i+1][6])
        PGL+= int(a[i+1][5])
        if a[i+1][6]>a[i+1][5]:
            PW +=1
        else:
            PL +=1
    name=["Robyn"]
    s=a[i+1][1:3] 
    res = any(elem in s for elem in name)
    if res == True:
        RP +=1
        RGW+= int(a[i+1][5])
        RGL+= int(a[i+1][6])
        if a[i+1][5]>a[i+1][6]:
            RW +=1
        else:
            RL +=1
    s=a[i+1][3:5] 
    res = any(elem in s for elem in name) 
    if res == True:
        RP +=1
        RGW+= int(a[i+1][6])
        RGL+= int(a[i+1][5])
        if a[i+1][6]>a[i+1][5]:
            RW +=1
        else:
            RL +=1
    name=["Tim"]
    s=a[i+1][1:3] 
    res = any(elem in s for elem in name)
    if res == True:
        DP +=1
        DGW+= int(a[i+1][5])
        DGL+= int(a[i+1][6])
        if a[i+1][5]>a[i+1][6]:
            DW +=1
        else:
            DL +=1
    s=a[i+1][3:5] 
    res = any(elem in s for elem in name) 
    if res == True:
        DP +=1
        DGW+= int(a[i+1][6])
        DGL+= int(a[i+1][5])
        if a[i+1][6]>a[i+1][5]:
            DW +=1
        else:
            DL +=1
    name=["Tom"]
    s=a[i+1][1:3] 
    res = any(elem in s for elem in name)
    if res == True:
        TP +=1
        TGW+= int(a[i+1][5])
        TGL+= int(a[i+1][6])
        if a[i+1][5]>a[i+1][6]:
            TW +=1
        else:
            TL +=1
    s=a[i+1][3:5] 
    res = any(elem in s for elem in name) 
    if res == True:
        TP +=1
        TGW+= int(a[i+1][6])
        TGL+= int(a[i+1][5])
        if a[i+1][6]>a[i+1][5]:
            TW +=1
        else:
            TL +=1
    
b= [JW,JL,GW,GL,TW,TL,PW,PL,RW,RL,DW,DL]
  
c= [JGW,JGL,GGW,GGL,TGW,TGL,PGW,PGL,RGW,RGL,DGW,DGL]

P= [JP,GP,TP,PP,RP,DP]


perc=[]
gamesdif = []
gamesdif2 = []
for i in range(6):
    o=b[i*2]/(b[i*2]+b[(i*2)+1])
    o=o*100
    p=c[i*2]-c[(i*2)+1]
    m= p/P[i]
    perc.append(o)
    gamesdif.append(p)
    gamesdif2.append(m)
   
#%%
matches_won = [b[0], b[2], b[4], b[6], b[8], b[10]]
matches_lost = [b[1], b[3], b[5], b[7], b[9], b[11]]
games_won= [c[0], c[2], c[4], c[6], c[8], c[10]]
games_lost= [c[1], c[3], c[5], c[7], c[9], c[11]]
perc = list(perc)
gamesdif= list(gamesdif)
gamesdif2=list(gamesdif2)
Player = ["Jamie","George","Tom","Paul","Robyn","Tim"]
#%%
# Create DataFrame
df2 = pd.DataFrame({
    "Player": Player,
    "Matches Played": P,
    "Matches Won": matches_won,
    "Matches Lost": matches_lost,
    "Win %": perc,
    "Games Won": games_won,
    "Games Lost": games_lost,
    "GD": gamesdif,
    "GD per Match": gamesdif2
})

# Sort DataFrame by Win % and reset index
df2_sorted = df2.sort_values(by="Win %", ascending=False).reset_index(drop=True)
df2_sorted.index = df2_sorted.index + 1  # Make the index start from 1

# Calculate and add Rank column (higher Win % gets higher rank)
df2_sorted['Rank'] = df2_sorted['Win %'].rank(ascending=False, method='min')

# Move 'Rank' column to the first position
df2_sorted = df2_sorted[['Rank'] + [col for col in df2_sorted.columns if col != 'Rank']]

# Apply formatting styles for Win % and GD per Match
styled_df = df2_sorted.style.format({
    "Win %": "{:.2f}%",  # Format Win % with 2 decimal places
    "GD per Match": "{:.2f}"  # Format GD per Match with 2 decimal places
}).background_gradient(subset=["Win %"], cmap="YlGnBu")  # Apply color gradient for Win %
styled_df = styled_df.background_gradient(subset=["GD per Match"], cmap="coolwarm_r")  # Apply color gradient for GD per Match

# Apply table border styles
styled_df = styled_df.set_table_styles([
    {'selector': 'table', 'props': [('border-collapse', 'collapse')]},  # Collapse table borders
    {'selector': 'th', 'props': [('border', '1px solid black'), ('padding', '5px')]},  # Header borders
    {'selector': 'td', 'props': [('border', '1px solid black'), ('padding', '5px')]}  # Cell borders
])

# Save the styled table to an HTML file (for visualization purposes)
styled_df.to_html("styled_table.html")

#%% Upload to Google Sheets
from gspread_dataframe import set_with_dataframe
from gspread_formatting import *
import gspread

# Authenticate with Google Sheets
gc = gspread.service_account(filename=r"C:\Users\JColc\Downloads\our-rock-451820-r5-919f8b93adff.json")  # Replace with your actual JSON file path

# Open the target spreadsheet and worksheet
spreadsheet = gc.open("Padel Tracker")  # Replace with your Google Sheets name
worksheet = spreadsheet.worksheet("Leaderboard 2")  # Replace with your target sheet name

# Ensure 'Win %' and 'GD per Match' columns are numeric (and handle any errors gracefully)
df2_sorted['Win %'] = pd.to_numeric(df2_sorted['Win %'], errors='coerce')
df2_sorted['GD per Match'] = pd.to_numeric(df2_sorted['GD per Match'], errors='coerce')

# Round Win % and GD per Match to 2 decimal places and convert to string for display
df2_sorted['Win %'] = df2_sorted['Win %'].apply(lambda x: f"{round(x, 2):.2f}")
df2_sorted['GD per Match'] = df2_sorted['GD per Match'].astype(float).map("{:.2f}".format)
# Sort the DataFrame by Rank to maintain the correct order
df2_sorted = df2_sorted.sort_values('Rank')

# Upload the sorted DataFrame (without styling, as Google Sheets doesn't support pandas styling)
set_with_dataframe(worksheet, df2_sorted)

#%% Formatting and Styling in Google Sheets
# Apply cell formatting for the entire range
format_cell_range(worksheet, 'A1:Z100', CellFormat(
    backgroundColor=color(0.9, 0.9, 0.9),  # Light gray background
    textFormat=textFormat(bold=True)  # Bold text for headers
))

# Apply a light yellow background to the 'Win %' column (F2:F7)
format_cell_range(worksheet, 'F2:F7', CellFormat(
    backgroundColor=color(0.9, 0.9, 0.2),  # Light yellow background
    textFormat=textFormat(bold=True)  # Bold text for 'Win %' column
))

format_cell_range(worksheet, 'J2:J7', CellFormat(
    numberFormat=NumberFormat(type='NUMBER', pattern='0.00')))

format_cell_range(worksheet, 'E2:E7', CellFormat(
    numberFormat=NumberFormat(type='NUMBER', pattern='0')))

# Format 'GD per Match' column (J2:J7) with 2 decimal places
#format_cell_range(worksheet, 'J2:J7', CellFormat(
 #   numberFormat=NumberFormat(type='NUMBER', pattern='#.00')  # 2 decimal places for GD per Match
#))

# Freeze the header row to keep it visible while scrolling
set_frozen(worksheet, rows=1)


























