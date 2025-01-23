
# file read test's snip

while 1 > 0:
    pos = f.seek(0, os.SEEK_END) # log file last location
    if pos > prevpos:
        f.seek(prevpos) # log file move to location before log append
        lines = f.read().count("\n") #number of new lines , seeks to end
        print("file increased " + str(pos - prevpos) + " lines " + str(lines))
        f.seek(prevpos)# log file seek cursor to location before log append
        for i in range(lines):
            last_row = xbook.sheets[0].api.UsedRange.Rows.Count
            cell_no = "A" + str(last_row+1)
            print ("writing to cell "+cell_no)
            xsheet1.range(cell_no).value = f.readline() # writing the newline in log to excel
        prevpos = pos