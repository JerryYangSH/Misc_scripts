import os
import pandas as pd

## please run "pip install pandas wheel xlrd xlsxwriter" before hand

old_files_dir = "old_files"
new_files_dir = "new_files"
output_files_dir = 'output_files'
old_cas_rns = set()

print("***************************************************************************************************************")
print("\tNote! Please put your old files in {} and new files in {}".format(old_files_dir, new_files_dir))
print("\tBefore run this script, please install python and all its dependent libs by \"pip install pandas wheel xlrd xlsxwriter\"")
print("***************************************************************************************************************")
for old_file in os.listdir(old_files_dir):
    print("reading old file : ", os.path.join(old_files_dir, old_file))
    old = pd.read_excel(os.path.join(old_files_dir, old_file))
    old_cas_rns |= set(old['CAS_RN'])

for new_file in os.listdir(new_files_dir):
    print("reading new file : ", os.path.join(new_files_dir, new_file))
    new = pd.read_excel(os.path.join(new_files_dir, new_file))
    added = new[~new['CAS_RN'].isin(old_cas_rns)]
    #output added file
    output_file_name = os.path.join(output_files_dir, new_file)
    print('generated output : ', output_file_name)
    writer = pd.ExcelWriter(output_file_name, 'xlsxwriter')
    added.to_excel(writer, sheet_name='added', index=False)
    writer.save()