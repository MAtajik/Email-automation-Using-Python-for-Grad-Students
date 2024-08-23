import os
import shutil
import pandas as pd
from docx2pdf import convert
from docxtpl import DocxTemplate
import win32com as win

source_path = r"E:\desktop\USA AUto\\"

email_df = pd.read_excel('Mother_of_all_emails.xlsx', sheet_name='email')
cv_df = pd.read_excel('Mother_of_all_emails.xlsx', sheet_name='cv')

email_doc = DocxTemplate("Email-template.docx")
cv_doc = DocxTemplate("CurriculumVitae-template.docx")
my_email = {}
my_cv = {}

for index, row in email_df.iterrows():
    for colum_name in email_df:
        my_email[colum_name] = row[colum_name]
    email_doc.render(my_email)
    email_doc.save(f"Email.docx")
    file_name = f"Email.docx"
    directory = 'Dr.' + my_email['last_name_of_professor']
    parent_dir = r"E:\desktop\USA AUto\Final Email\\"
    path = os.path.join(parent_dir + directory)
    os.mkdir(path)
    shutil.move(source_path + file_name, path + "\\" + file_name)

for index, row in cv_df.iterrows():
    for colum_name in cv_df:
        my_cv[colum_name] = row[colum_name]
    cv_doc.render(my_cv)
    cv_doc.save(f"Curriculum Vitae.docx")
    convert(input_path=f"Curriculum Vitae.docx", output_path=f"Curriculum Vitae.pdf")
    os.remove(f"Curriculum Vitae.docx")
    file_name = f"Curriculum Vitae.pdf"
    directory = 'Dr.' + my_email['last_name_of_professor']
    parent_dir = r"E:\desktop\USA AUto\Final Email\\"
    path = os.path.join(parent_dir + directory)
    shutil.move(source_path + file_name, path + "\\" + file_name)
