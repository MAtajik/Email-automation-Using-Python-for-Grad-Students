import pandas as pd
import win32com as win

email_df = pd.read_excel('Mother_of_all_emails.xlsx', sheet_name='email')
cv_df = pd.read_excel('Mother_of_all_emails.xlsx', sheet_name='cv')

# last_name_of_professor = input("Enter the Last Name of the Professor: ")
# Email = input("Enter the Email: ")
# intended_program = input("Enter the Intended Program: ")
# university_name = input("Enter the University Name: ")
# intended_session = input("Enter Intended Session: ")
# professor_research_field = input("Enter Professor's Research Field You\'r interested in: ")
# article_1 = input("Enter Professor's Article No-1: ")
# article_2 = input("Enter Professor's Article No-2: ")
# comments_on_article_1 = input("Write your view on Article No-1: ")
# comments_on_article_2 = input("Write your view on Article No-2: ")
# topics_i_have_worked_on = input("The topic you have worked on: ")
# my_research_interests = input("My Research interest in CV for this email: ")

# my_email = {'last_name_of_professor': last_name_of_professor,
#             'Email': Email,
#             'intended_program': intended_program,
#             'university_name': university_name,
#             'intended_session': intended_session,
#             'professor_research_field': professor_research_field,
#             'article_1': article_1,
#             'comments_on_article_1': comments_on_article_1,
#             'article_2': article_2,
#             'comments_on_article_2': comments_on_article_2,
#             'topics_i_have_worked_on': topics_i_have_worked_on}
#
# my_cv = {'my_research_interests': my_research_interests}

# for colum_name in email_df:
#     my_email[colum_name] = colum_name
# print(my_email)

#for column_name in email_df:
#    for key, value in my_email.items():
#       if column_name == key:
            
#    my_email[column_name]
#     pd.concat(my_email)
#
#
# for index, row in cv_df.iterrows():
#     cv_df.concat(my_cv)


