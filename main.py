import pandas as pd
import openpyxl
import re


class MathEnglish:
    columns = ['Student ID', 'Student Name','Student Email', 'LCP', 'Major','Ed Plan', 'English', 'Math', 'Total Units', 'Course List', 'M/E']
    lcp_df = pd.DataFrame(columns=columns)
    ed_plans = ['ASEP: Abbreviated Educational Plan', 'CSEP: Comprehensive Educational Plan']
    math_courses = ['MATH-104', 'MATH-110A', 'MATH-112', 'MATH-112S', 'MATH-114', 'MATH-116', 'MATH-140']
    english_courses = ['ENGL-100', 'ENGL-100S']

    def __init__(self, id, student_df):
        self.id = id
        self.df = df
        self.student_df = student_df

    def enrollments(self):
        english_course = 'None'
        math_course = 'None'
        total_units = 0
        course_list = []

        for i in range(len(self.student_df)):
            if self.student_df.loc[i, 'Course Number'] in MathEnglish.english_courses:
                english_course = self.student_df.loc[i, 'Course Number']

            if self.student_df.loc[i, 'Course Number'] in MathEnglish.math_courses:
                math_course = self.student_df.loc[i, 'Course Number']


            if self.student_df.loc[i, 'Course Number'] not in course_list:
                course_list.append(self.student_df.loc[i, 'Course Number'])
                total_units = total_units + self.student_df.loc[i, 'Credit Hours']

        # if english_course == None:
        #     english_course = 'None'
        # if math_course == None:
        #     math_course = 'None'
        return english_course, math_course, self.id, total_units, course_list


    def lcp(self):

        lcps = ['Arts, Humanities, & Communication', 'Applied Technology & Skilled Trades',
                'Health Sciences & Wellness',
                'Social & Behavioral Sciences', 'Science, Engineering, & Math', 'Business, Accounting, & Law',
                'Education & Human Services', 'Exploration & Discovery', 'VETS', 'Umoja', 'Transfer Academy']

        for i in range(len(self.student_df)):
            lcp = None
            categorie = self.student_df.loc[0, 'Categories']
            for lcp_item in lcps:
                if re.search(pattern=lcp_item, string=categorie):
                    lcp = lcp_item
                    break
                if lcp == None:
                    lcp = 'None'

        return lcp

    def ed_plan(self):

        categorie_list = []
        ed_plan = None
        categorie_list.append(self.student_df.loc[0, 'Categories'])
        string='  '.join([str(item) for item in categorie_list])
        categorie_list.append(self.student_df.loc[0, 'Categories'])
        string = '  '.join([str(item) for item in categorie_list])
        string = string.split(',')
        for item in string:
            if item in MathEnglish.ed_plans:
                ed_plan = item

        return ed_plan

    def report_generator(self, english, math, lcp, ed_plan, total_units, course_list):

        length = len(MathEnglish.lcp_df)
        MathEnglish.lcp_df.loc[length, 'Student ID'] = self.id
        MathEnglish.lcp_df.loc[length, 'Student Name'] = self.student_df.loc[0, 'Student Name']
        MathEnglish.lcp_df.loc[length, 'Student Email'] = self.student_df.loc[0, 'Student E-mail']
        MathEnglish.lcp_df.loc[length, 'LCP'] = lcp
        MathEnglish.lcp_df.loc[length, 'Major'] = self.student_df.loc[0, 'Major']
        MathEnglish.lcp_df.loc[length, 'Ed Plan'] = ed_plan
        MathEnglish.lcp_df.loc[length, 'English'] = english
        MathEnglish.lcp_df.loc[length, 'Math'] = math
        MathEnglish.lcp_df.loc[length, 'Total Units'] = total_units
        MathEnglish.lcp_df.loc[length, 'Course List'] = course_list


        return MathEnglish.lcp_df

def calculations(lcp_df):

    english_total_df = MathEnglish.lcp_df[MathEnglish.lcp_df['English'] != 'None']
    english_total = len(english_total_df)
    math_total_df = MathEnglish.lcp_df[MathEnglish.lcp_df['Math'] != 'None']
    math_total = len(math_total_df)
    for i in range(len(MathEnglish.lcp_df)):
        eng = 'no'
        math = 'no'
        if MathEnglish.lcp_df.loc[i, 'English'] != 'None':
            eng = 'yes'
        if MathEnglish.lcp_df.loc[i, 'Math'] != 'None':
            math = 'yes'

        if eng == 'yes' and math == 'yes':
            MathEnglish.lcp_df.loc[i, 'M/E'] = 'yes'
    both_total_df = MathEnglish.lcp_df[MathEnglish.lcp_df['M/E'] == 'yes']
    both_total = len(both_total_df)

    print(english_total,math_total, both_total)

def lcp_sheets(lcp_df):
    lcps = ['Arts, Humanities, & Communication', 'Applied Technology & Skilled Trades',
            'Health Sciences & Wellness','Social & Behavioral Sciences', 'Science, Engineering, & Math',
            'Business, Accounting, & Law', 'Education & Human Services', 'Exploration & Discovery', 'VETS', 'Umoja',
            'Transfer Academy']

    for lcp_item in lcps:
        print(lcp_df)
        lcp_df2 = lcp_df[lcp_df['LCP'] == lcp_item]
        lcp_df2.to_excel(lcp_item + '.xlsx')


df = pd.read_csv('C:/Users/flmix/OneDrive/Documents/Python Programming/EAB FY Student List.csv')
# df = pd.read_csv(
#     'C:/Users/fmixson/Desktop/Dashboard_files/LCP_English_Math/campus-v2report-enrollment-2023-05-11 (1).csv')
pd.set_option('display.max_columns', None)
df = df[df['Dropped'] == 'No'].reset_index()
id_list = []
for i in range(len(df)):
    if df.loc[i, 'Student ID'] not in id_list:
        id_list.append(df.loc[i, 'Student ID'])

for id in id_list:
    student_df = df[df['Student ID'] == id]
    student_df = student_df.reset_index()
    student = MathEnglish(id=id, student_df=student_df)
    engl_course, math_course, id, total_units, course_list = student.enrollments()
    lcp = student.lcp()
    ed_plan = student.ed_plan()
    lcp_df = student.report_generator(english=engl_course, math=math_course, lcp=lcp, ed_plan=ed_plan, total_units=total_units,
                                      course_list=course_list)
calculations(lcp_df=lcp_df)
lcp_sheets(lcp_df=lcp_df)
lcp_df.to_excel('EM_df.xlsx')
# df[['First', 'Second', 'Third']] = df.Dropped.str.split(',', expand=True)
# print(df)d', 'Third']] = df.Dropped.str.split(',', expand=True)
# print(df)