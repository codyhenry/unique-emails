import pandas
import os

def check_folder_validity(root):
    numFiles = 0
    for item in os.listdir(root):
        if item.endswith('.csv' or '.xlsx') and os.path.isfile(os.path.join(root,item)):
            numFiles += 1
            if(numFiles > 2):
                raise Exception('Too many excel files exist in this folder')

def get_worksheet(file):
    if entry.name.endswith('.xlsx'):
        return pandas.read_excel(entry.name)
    elif entry.name.endswith('.csv'):
        return pandas.read_csv(entry.name)

def find_unique_from_robly(robly,leads):
    for email in robly:
        if email in leads.values:
           index = leads.loc[leads['email'] == email].index[0]
           leads = leads.drop([index])
    return leads

def find_duplicates(leads,robly):
    result = leads.query('email in @robly')
    print(str(len(result)) + ' duplicate emails found')
    return result

def find_unique_from_leads(leads,robly):
    
    duplicates = find_duplicates(leads,robly)
    temp = pandas.concat([leads, duplicates])
    result = temp.drop_duplicates(keep=False)
    print(result)
    return result

def create_new_csv(path,data):
    data.to_csv(path, index=False)

root  = 'C:\\Users\\cody\\Documents\\PythonProjects\\excel'
subfolder = root + '\\output'
csvOutput = 'output.csv'
txtOutputl = 'output.txt'
os.chdir(root)

check_folder_validity(root)

for entry in os.scandir(root): 
    if entry.is_file():
        dataFrame = get_worksheet(entry)
        #determine if Robly from headers
        if entry.name == 'robly.csv':
            robly = set(dataFrame['email'])
            print(str(len(robly)) + ' entries in robly')
        else:
            leads = dataFrame[["name","email"]]
            print(str(len(leads)) + ' entries in leads')
          
if len(robly) < len(leads):
    uniqueLeads = find_unique_from_robly(robly,leads)
else:
    uniqueLeads = find_unique_from_leads(leads,robly)

print(str(len(uniqueLeads)) + ' new leads')
        
if not os.path.exists(subfolder):
    os.makedirs('./excel/output')

pathName = os.path.join(root,subfolder)
create_new_csv(os.path.join(pathName, csvOutput), uniqueLeads)
