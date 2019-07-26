### It checks whether paid file is exists or not, if not then creates
### It compare Master, Maintenance and Payment (Bank statement) and create two new files, they are, Paid and unpaid files in respective folders.
### Based on Month and Year given in the 'Input sheet', check the Month and Year values in Payment sheet only
### Based on Due list file it checks Master excel file and take Names and compare with Payment sheet
### Prepare Final payment and dues list with dues
### Convert due file to xlsx and try to accumulate previous dues --- Test carefully for Feb2018
import datetime
import xlrd
import csv
import os
import os.path
from os import path
import argparse
import csv
import sys
from openpyxl import Workbook

format_str = '%d/%m/%Y'
INSERT_DT = datetime.datetime.now()
curr_dir = os.getcwd().replace("\\","/")
curr_dir_file = os.path.join(curr_dir,'Maintenance_input_files.xlsx')
filename = curr_dir_file.replace("\\","/")
print(filename)
due_list = xlrd.open_workbook(filename)
due_list_sheet = due_list.sheets()
for x in due_list_sheet:
        f0 = x.row_values(0)
        f1 = x.row_values(1)
        f2 = x.row_values(2)
        f3 = x.row_values(3)
        f4 = x.row_values(6)
        f5 = x.row_values(5)
        f6 = x.row_values(4)
print(f0)
print(f1)
print(f2)
print(f3)
print(f4)
print(f5)
print(f6)
extr_mnth_yr = str(f1[3])
c = extr_mnth_yr.replace('"','')
#print(c)
conv_mn_yr_dt = datetime.datetime.strptime(c,"%b-%Y")
conv_mn_yr_dt_str = datetime.date.strftime(conv_mn_yr_dt,'%m/%Y')
#print(conv_mn_yr_dt)
#print("conv_mn_yr_dt_str",conv_mn_yr_dt_str)
year = str(INSERT_DT.year)
month = INSERT_DT.strftime("%B")
paid = f3[0]
paid1 = paid.replace(".csv","")
paid_file1 = paid1+"_"+c+".csv"
#print(paid_file1)
maintbook = xlrd.open_workbook(f0[0])
maintsheet = maintbook.sheets()
mbook = xlrd.open_workbook(f2[0])
msheets = mbook.sheets()
paybook = xlrd.open_workbook(f1[0])
paymbook = paybook.sheets()
due_list_rep = f4[0].replace(".csv","")
due_list_mnth = due_list_rep+"_"+c+".csv"
due_list_xlsx = due_list_rep+"_"+c+".xlsx"
due_list = open(due_list_mnth,'w+')
paid_dummy = open(f5[0],'w+')
master_dummy = open(f6[0],'w+')
if str(path.exists(paid_file1)):
        #print("File not exists so it is created")
        paid_file = open(paid_file1, "w+")
        paid_file.write("FLAT_NO"+","+"PAID_DATE"+","+"PAID_DESC"+","+"PAID_AMOUNT"+"DUE_AMOUNT"+ "\n")
i = 0
for msheet in msheets:
        mrows = msheet.nrows
        loop_end = 0
        for maint_sheet_row in maintsheet:
            maint_rows = maint_sheet_row.nrows
            k = 0
            m_inc_val = 1
            s = 0
            f = 0
            for payment in paymbook:
                        paycheck = payment.row_values(0)
                        pay_rows = payment.nrows
                        while k < mrows:
                                master = msheet.row_values(k)
                                mNo = str(master[0])
                                mName = master[1]
                                i = 0
                                while i < maint_rows:
                                    FLAT_NO = mNo.replace(".0","")
                                    maints = maint_sheet_row.row_values(i)
                                    maint_flatNo = str(maints[0])
                                    ma_flatNo = maint_flatNo.replace(".0","")
                                    no_of_fields = len(maints)
                                    if FLAT_NO == ma_flatNo:
                                        Amount = maints[no_of_fields - 1 ]
                                        #print("Mainteance Flat no:",ma_flatNo)
                                        #print("Maintenance Amount:",Amount)
                                        j = 0
                                        matched_rows = 0
                                        while j < pay_rows:
                                            paychk = payment.row_values(j)
                                            #print(paychk)
                                            if conv_mn_yr_dt_str in paychk[4]:
                                                    #print(paychk)
                                                    matched_rows = matched_rows + 1
                                                    #print("Total matched rows:",matched_rows)
                                                    payamt = paychk[17]
                                                    if payamt != ' ':
                                                            paydt = paychk[4]
                                                            payamt1 = payamt.replace(".00",".0")
                                                            pay = payamt1.replace(",","")
                                                            paydesc = paychk[9]
                                                            maint_Amt = Amount.replace(".00",".0")
                                                            maintAmt = maint_Amt.replace(",","")
                                                            remvstr_Amt = maintAmt.replace("Rs.","")
                                                            Amt = float(remvstr_Amt)
                                                            print("Amount in payment:",payamt)
                                                            payAmt = float(pay)
                                                            print("Master Flat no:",ma_flatNo)
                                                            print("Amount in Maintenace:",Amt)
                                                            print("Amount in payment:",payAmt)
                                                            if payAmt == Amt:
                                                                print("Total matched rows:",matched_rows)
                                                                due = 0
                                                                paid_file.write(ma_flatNo+","+paydt+","+paydesc+","+pay+","+str(due)+ "\n")
                                                                paid_dummy.write(ma_flatNo+ "\n")
                                            j = j + 1  
                                    i = i + 1
                                master_dummy.write(FLAT_NO+ "\n")
                                k = k + 1
                                m_inc_val = m_inc_val + 1
                        s = s + 1
loop_end = loop_end + 1
paid_file.close()
paid_dummy.close()
master_dummy.close()
with open(master_dummy.name,'r') as f12:
        #print(f12)
        d=set(f12.readlines())
with open(paid_dummy.name,'r') as f21:
        e=set(f21.readlines())
with open(due_list.name,'a') as f31:
     for line in list(d-e):
           #print("First due list:",d-e)
           f31.write(line)
f31.close()
f12.close()
f31.close()
due_list.close()
x = paid_dummy.name

#### Compare both Due list and Payment sheet
def main():
    input_file = due_list.name
    wb = Workbook()
    worksheet = wb.active
    for row in csv.reader(open(input_file), delimiter="\t"):
        worksheet.append([_convert_to_number(cell) for cell in row])
    wb.save(input_file.replace(".csv", ".xlsx"))

def _convert_to_number(cell):
    if cell.isnumeric():
        return int(cell)
    try:
        return float(cell)
    except ValueError:
        return cell
main()
#print("Master sheet name:",f2[0])
dbook = xlrd.open_workbook(due_list_xlsx)
dsheets = dbook.sheets()
d = 0
for due in dsheets:
        drows = due.nrows
        while d < drows:
                #print(drows)
                dvalue = due.row_values(d)
                #print("Dues Flatno:",dvalue)
                due_FNo = str(dvalue[0])
                dFno = due_FNo.replace(".0","")
                ma = 0
                for msheet in msheets:
                        marows = msheet.nrows
                        loop_end = 0
                        while ma < marows:
                                master = msheet.row_values(ma)
                                mNo = str(master[0])
                                mName = master[1]
                                #print("Master Name:",mName)
                                mFNo = mNo.replace(".0","")
                                #print(dFno)
                                #print(mFNo)
                                if mFNo == dFno:
                                        #print("Due No",dFno)
                                        #print("master No",mFNo)
                                        #print("Master Name:",mName)
                                        for payment in paymbook:
                                                #paycheck = payment.row_values(0)
                                                pay_rows = payment.nrows
                                                #print(pay_rows)
                                                j = 0
                                        #matched_rows = 0
                                        while j < pay_rows:
                                            paychk = payment.row_values(j)
                                            if conv_mn_yr_dt_str in paychk[4]:
                                                    if mName in paychk[9]:
                                                        main = 0          
                                                        for maint_sheet_row in maintsheet:
                                                                maint_rows = maint_sheet_row.nrows
                                                                while main < maint_rows:
                                                                    maints = maint_sheet_row.row_values(main)
                                                                    maint_flatNo = str(maints[0])
                                                                    ma_flatNo = maint_flatNo.replace(".0","")
                                                                    no_of_fields = len(maints)
                                                                    if dFno == ma_flatNo:
                                                                        Amount = maints[no_of_fields - 1 ]
                                                                        payamt = paychk[17]
                                                                        payamt1 = payamt.replace(".00",".0")
                                                                        pay = payamt1.replace(",","")
                                                                        Amount1 = Amount.replace("Rs.","")
                                                                        Amt = float(Amount1.replace(",",""))
                                                                        #print("Amount in Maintenance sheet:",Amt)
                                                                        #print("Amount in payment sheet:",float(pay))
                                                                        #Amount1 = str(Amount)

                                                                        if Amt > float(pay):
                                                                                
                                                                                due1 = Amt - float(pay)
                                                                        elif Amt < float(pay):
                                                                                due1 = float(pay) - Amt
                                                                        elif Amt == float(pay):
                                                                                due1 = 0
                                                                                
                                                                        paydt = paychk[4]
                                                                        paydesc = paychk[9]
                                                                        with open(paid_file.name,'a') as paid:
                                                                                paid.write(mFNo+","+paydt+","+paydesc+","+pay+","+str(due1)+ "\n")
                                                                    main = main + 1   
                                            j = j + 1
                                ma = ma + 1
                        d = d + 1
paid.close()
paid_f_dummy = open("paid_f_dummy.txt",'w+')

with open(paid_file.name,'r') as paid_file:
        for paid in paid_file:
                if paid[0:3] != "FLA":
                        paid_f_dummy.write(paid[0:3]+ "\n")
                        
paid_file.close()
paid_f_dummy.close()
dues_final = "Dues_Final"+"_"+c+".csv"
due_list_dummy = open(dues_final,'w+')
with open(paid_f_dummy.name,'r') as f12:
        d=set(f12.readlines())
with open(due_list.name,'r') as f21:
        e=set(f21.readlines())
with open(due_list_dummy.name,'a') as f31:
     for line in list(e-d):
           f31.write(line)
f12.close()
f21.close()
f31.close()
os.remove(paid_dummy.name)
os.remove(master_dummy.name)
os.remove(due_list.name)
os.remove(due_list_xlsx)
os.remove(paid_f_dummy.name)
f31.close()
due_list_dummy.close()
with open(due_list.name,'w+') as due_final:
        with open(due_list_dummy.name,'r') as f13:
                for a in f13:
                        due_final.write(a+ '\n')
f13.close()
due_final.close()
os.remove(dues_final)                        
if loop_end == msheets:
    myconn.close()          
