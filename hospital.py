import os
import openpyxl as xl
from openpyxl.styles import Font
import re

class Hospital:
    def __init__(self):
       self.boldFont = Font(bold=True)
       self.sheets = ['Auth','Patients','Doctors','Departments','Appointments']
       self.wb = self.create_db_if_not_exists()

    def create_db_if_not_exists(self):
        if not os.path.isfile('db.xlsx'):
            wb = xl.Workbook()
            for i in range(len(self.sheets)):
                sheet = wb.create_sheet(self.sheets[i],i)
                sheet['A1'].font = self.boldFont
                if i==0:
                    sheet['A1'].value = 'Rule'
                    sheet['B1'].value = 'Secret'
                    sheet['B1'].font = self.boldFont
                    sheet['A2'].value = 'admin'
                    sheet['B2'].value = 'admin123'
                    sheet['A3'].value = 'user'
                    sheet['B3'].value = 'user123'
                else:
                    for j in ['A','B','C','D']: 
                        sheet.column_dimensions[j].width=22
                        sheet[j+str(1)].font = self.boldFont
            wb['Patients']['A1'].value = 'ID'
            wb['Patients']['B1'].value = 'NAME'
            wb['Patients']['C1'].value = 'ADDRESS'
            wb['Patients']['D1'].value = 'AGE'

            wb['Doctors']['A1'].value = 'ID'
            wb['Doctors']['B1'].value = 'NAME'
            wb['Doctors']['C1'].value = 'DEPARTMENT'

            wb['Appointments']['A1'].value = 'ID'
            wb['Appointments']['B1'].value = 'DOCTOR'
            wb['Appointments']['C1'].value = 'PATIENT'
            wb['Appointments']['D1'].value = 'DATE'

            wb['Departments']['A1'].value = 'ID'
            wb['Departments']['B1'].value = 'NAME'
            wb.save('db.xlsx')
            return wb
        else: return xl.load_workbook('db.xlsx')

    def validate(self,field,pattern,err):
        while True:
            x = input('\nEnter '+field+': ')
            if pattern.match(str(x)): return x
            else: print(err)

    def authenticate(self,level,pwd):
        auth = self.wb['Auth']
        return auth[level].value == pwd

    def show_main_menu(self):
        while True:
            if self.rule=='admin':
                x = input('\n1: Manage patients\n2: Manage doctors\n3: Manage appointments\n4: Manage departments\n0: To go back\n\n'+self.rule+'$ ')
                if x=='0': break
                elif x == '1':  self.show_admin_menu('Patients')
                elif x == '2':  self.show_admin_menu('Doctors')
                elif x == '3':  self.show_admin_menu('Appointments')
                elif x == '4':  self.show_admin_menu('Departments')
                else: print('\nInvalid input!\n')    
            else: 
                x = input('\n1: View patients\n2: View doctors\n3: View appointments\n4: View departments\n0: To go back\n\n'+self.rule+'$ ')
                if x=='0': break
                elif x=='1':
                    self.view('Patients')

    def show_admin_menu(self,title):
        while True:
            x = input('\n1: Add '+title+'\n2: Delete '+title+'\n3: View all '+title+'\n0: To go back\n\n'+self.rule+'$ ')
            if x == '0': break
            elif x == '1': self.prompt_addition(title)
            elif x == '2': 
                x = self.validate(title+' id',re.compile(r'\d+'),'Invalid ID! id must be a number')
                self.delete_record(title,int(x))
            elif x == '3': self.view(title)
            else: print('\nInvalid input!\n')

    def add_record(self,title,data):
        sheet = self.wb[title]
        max_row = sheet.max_row
        max_id = sheet.cell(row=max_row,column=1).value
        if type(max_id) != int: max_id = 0
        for i in range(len(data)+1):
            sheet.cell(row=max_row+1,column=i+1).value = max_id+1 if i==0  else data[i-1]
        self.wb.save('db.xlsx')

    def prompt_addition(self,title):
        if title == 'Patients':
            n = self.validate('patient name',re.compile(r'^[a-zA-z]{1}.*'),'Invalid Name! name should start with character')
            ad = self.validate('patient address',re.compile(r'^[a-zA-z]{1}.*'),'Invalid Address! address should start with character')
            a = self.validate('patient age',re.compile(r'\d+'),'Invalid Age! age must be a number')
            self.add_record(title,[n,ad,a])
        elif title == 'Doctors':
            n = self.validate('doctor name',re.compile(r'^[a-zA-z]{1}.*'),'Invalid Name! name should start with character')
            d = self.validate('department',re.compile(r'^[a-zA-z]{1}.*'),'Invalid Department! department should start with character')
            self.add_record(title,[n,d])
        elif title == 'Appointments':
            dc = self.validate('doctor name',re.compile(r'^[a-zA-z]{1}.*'),'Invalid Name! name should start with character')
            p = self.validate('patient name',re.compile(r'^[a-zA-z]{1}.*'),'Invalid Name! name should start with character')
            d = self.validate('date',re.compile(r'^[0-9]{4}.[0-1]?[0-9].[0-3]?[0-9]$'),'Invalid Date! date should follow this format yyyy-mm-dd')
            self.add_record(title,[dc,p,d])
        elif title == 'Departments':
            n = self.validate('department name',re.compile(r'^[a-zA-z]{1}.*'),'Invalid Name! name should start with character')
            self.add_record(title,[n])

    def view(self,title):
        sheet = self.wb[title]
        if sheet.max_row==1: print('\n No '+title+' registered yet!\n'); return
        items = []
        maxLens = []
        print('\n'+title+'\n')
        for i in range(1,sheet.max_column+1):
            rowLen = 0
            row = []
            for j in range(1,sheet.max_row+1):
                rowValue = sheet.cell(row=j,column=i).value
                if rowValue==None:break
                row.append(rowValue)
                if len(str(rowValue))> rowLen: rowLen=len(str(rowValue))
            else:
                items.append(row)
                maxLens.append(rowLen)
        for i in range(len(items[0])):
            x=''
            for j in range(len(items)):
                x += str(items[j][i]).ljust(maxLens[j]+10)
            print(x)

    def delete_record(self,title,id):
        sheet = self.wb[title]
        if sheet.max_row==1: print('\nNo '+title+' registered yet!\n');return
        for i in range(2,sheet.max_row+1):
            if sheet.cell(row=i,column=1).value == id:
                if i < sheet.max_row:
                    for j in range(i+1,sheet.max_row+1):
                        diff = (j-i-1)+i
                        for k in range(1,sheet.max_column+1):
                            sheet.cell(row=diff,column=k).value = sheet.cell(row=j,column=k).value
                sheet.delete_rows(sheet.max_row,1)
                self.wb.save('db.xlsx')
                break
        else: print('\nNo '+title+' found with the id '+str(id)+'.')


    def start(self):
        while True:
            print('\n\n\n',' Hospital Management System '.center(46,'*'))
            print('\n\n','Select your rule'.center(36,'-'))
            print('\n1: Admin\n2: User\n\nEnter Q to quit.\n')
            x = input('')
            if x == 'q' or x=='Q': print('\nbye.\n'); break
            elif x=='1':
                while True:
                    a = input('\nEnter Admin Password: ')
                    if self.authenticate('B2',a): self.rule = 'admin'; self.show_main_menu(); break
                    else: print('\nIncorrect password!\n')
            elif x=='2':
                while True:
                    u = input('\nEnter User Password: ')
                    if self.authenticate('B3',u): self.rule = 'user'; self.show_main_menu(); break
                    else: print('\nIncorrect password!\n')
            else: print('\nInvalid input!\n')

    
        
