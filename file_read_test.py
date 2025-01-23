import os
import xlwings as xw
import time

refresh_interval = 0.1
f = open(r"C:\Users\kmini\AppData\Roaming\MetaQuotes\Terminal\Common\Files\buffer.csv","r")
xbook = xw.Book(r"openpy_tester1.xlsx")
print(" === LOG : opened  the excel file === ")
xsheet1 = xbook.sheets['Sheet1']
xsheet1.range('A1').value = "connected to mt4 and excel !"

xsheet2 = xbook.sheets['Sheet2']
xsheet3 = xbook.sheets['Sheet3']

prevpos = f.seek(0, os.SEEK_END)
pos = prevpos

lines = f.read().count("\n")

def find_row (col,search_string) :
    col=int(col)
    last_row = xsheet3.api.UsedRange.Rows.Count
    for row in range(last_row, 0, -1):
        val = xsheet3.range(row, col).value
        if (str(val) == "") :
            continue
        if str(search_string) == str(val):
            #print("seen in row " + str(row), end=" ")#print row that symbol is found in
            return row
    print ("not found")
    return 0

def append_sheet2 (enter_string):
    last_row = xsheet2.api.UsedRange.Rows.Count
    cell_no = "A" + str(last_row + 2)
    xsheet2.range(cell_no).value = enter_string

try:
    while 1 > 0:
        time.sleep(refresh_interval)
        xsheet1.range('A1').value = int(time.time())
        pos = f.seek(0, os.SEEK_END) # log file last location
        if pos > prevpos:
            f.seek(prevpos) # log file move to location before log append
            lines = f.read().count("\n") #number of new lines , seeks to end
            # print("\n\nlines added = "+ str(lines) + ", bytes added = "+ str(pos - prevpos) )
            f.seek(prevpos) # log file seek cursor to location before log append
            for i in range(lines):
                log = f.readline().strip()#the excel extract
                logarr = log.split(";")#splitted excel extract
                #print (logarr) #print the splitted excel extract
                select_row = find_row(1,logarr[0])
                if select_row == 0:
                    append_sheet2("error log : " + "".join(logarr))
                    print("ERROR : LOG NOT FOUND : " + log )
                else :
                    logarr.pop(0)
                    cell_no = "B" + str(select_row)
                    xsheet3.range(cell_no).value = ";".join(logarr);
                    # print ("LOG: " + log + " PLACED IN "+ str(select_row) +" | ")
                # last_row = xbook.sheets[0].api.UsedRange.Rows.Count
                # cell_no = "A" + str(last_row+1)
                # print ("writing to cell "+cell_no)
            prevpos = pos
except Exception as e:
    print("\nERROR: code stopped \n" )
    print (e)

xsheet1.range('A1').value = 0
f.close()


"""
tasks done :
- string search in row 
- writing to excel 
- reading from excel
- fetch data from log file 
- constant first row with symbols in excel and any cell updating needs to highlight for 10s
- every log neeeds to be updated beside the symbol
- the alerts need to be updated in one of the column 
- RSI crossing and levels
- Stocastic levels 



pending tasks:
- multi TF indicators need to be deployed 
- MACD levels
- MA directions 

"""