#
# Script compares contents of two Excel files
#
# May 18, 2023
# Eric Schares

# Flow:
# Go to https://www.elsevier.com/about/policies/pricing and download the article-publishing-charge.xlsx file
# Read in yesterday's APC price list
# For each ISSN, look up the price in today's file
# Report any changes, including title, delta, old, and new price
# Report any ISSNs in yesterday's file that are no longer in today's (removed)
# 
# Read in today's APC price list
# Already did the deltas, so don't need to do that again
# Report any ISSNs in today's file that are not in yesterday's (new titles added)
# Those new titles will start getting tracked when this file turns into yesterday's
#
#

import pandas as pd
import requests
import os
import datetime
import shutil


def compare_excel_files(file1, file2):
    df1 = pd.read_excel(file1, header=3, usecols=[0,1,2,3], names=['ISSN', 'Title', 'Business Model', 'USD'])
    df2 = pd.read_excel(file2, header=3, usecols=[0,1,2,3], names=['ISSN', 'Title', 'Business Model', 'USD'])

    diff_df = df1.compare(df2)

    if diff_df.empty:
        print("No differences found.")
    else:
        print(str(diff_df.shape[0]) + " Differences Found:")
        for idx, row in diff_df.iterrows():
            #print(f"Row {idx}:")
            #print(row)
            delta = df2.iloc[idx]['USD'] - df1.iloc[idx]['USD']
            print(df1.iloc[idx]['Title'] + " increased the APC by $" + str(delta) + ", from $" + str(df1.iloc[idx]['USD']) + " to $"+ str(df2.iloc[idx]['USD']))


def scrape_apc_list():
    url = 'https://www.elsevier.com/books-and-journals/journal-pricing/apc-pricelist'
    filename = 'article-publishing-charge.xlsx'

    response = requests.get(url)

    if response.status_code == 200:
        with open(filename, 'wb') as file:
            file.write(response.content)
        print(f"File '{filename}' downloaded successfully.")
    else:
        print("Failed to download the file.")

def rename_file_with_date(file_path):
    # Extract the file directory and base filename
    directory, filename = os.path.split(file_path)
    
    # Get today's date
    today = datetime.date.today()
    
    # Append the date to the filename
    new_filename = f"{os.path.splitext(filename)[0]}_{today.strftime('%Y-%m-%d')}{os.path.splitext(filename)[1]}"
    
    # Construct the new file path
    new_file_path = os.path.join(directory, 'files', new_filename)
    
    # Rename the file
    #os.rename(file_path, new_file_path)
    shutil.move(file_path, new_file_path)
    
    print(f"The file has been renamed to: {new_file_path}")

def find_files_with_date(directory, days_back):
    # Get date
    the_date = datetime.date.today() - datetime.timedelta(days=days_back)  # days_back=1 to get yesterday, 0 to get today
    date_string = the_date.strftime('%Y-%m-%d')
    
    # List all files in the directory
    files = os.listdir(directory)
    
    # Filter files that contain the date
    matching_files = [file for file in files if date_string in file]
    
    return matching_files


############ Main ################

if (0):
    # Download Elsevier's article-publishing-charge.xlsx file
    scrape_apc_list()

    # Change filename to have today's date
    file_path = 'article-publishing-charge.xlsx'
    rename_file_with_date(file_path)



# Find the APC price lists corresponding to today (0) and yesterday (1)
directory = '../files'
matching_files = find_files_with_date(directory, 1)     # pass delta of 1 to get file with yesterday's date
if len(matching_files) != 1:
    print("Something went wrong looking for yesterday's file")  # Should only find one file, if 0 or 2+ something is wrong
# Assign yesterday's file to file1
file1 = matching_files[0]

matching_files = find_files_with_date(directory, 0)     # pass delta of 0 to get file with today's date
if len(matching_files) != 1:
    print("Something went wrong looking for today's file")  # Should only find one file, if 0 or 2+ something is wrong
# Assign today's file to file2
file2 = matching_files[0]

print (f"Yesterday's file: {file1}\nToday's file:     {file2}\n")     # keep the weird spacing so the print output looks nice



# Compare them
#compare_excel_files(file1, file2)
