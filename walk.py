# !/usr/bin/python

import os
from openpyxl import Workbook, load_workbook

#takes coordinates and woorksheet
#in the column specified, will get all the row values
def getRows(row, col, destc, worksheet, outSheet):
    c = worksheet.cell(row = row, column = col)
    s = outSheet.cell(row = row, column = destc)
    while c.value:
        print c.value
        print row
        print col
        s.value = c.value
        row += 1
        c = worksheet.cell(row = row, column = col)
        s = outSheet.cell(row = row, column = destc)


def main():

    docs = []
    for root, dirs, files in os.walk(".", topdown=False):
        for name in files:
            print(os.path.join(root, name))
            full_path = os.path.join(root, name)
            full_path = full_path[2:]
            if full_path.endswith(".xlsx"):
                docs.append(full_path)
                #text_file.write("%s\n" % full_path)

    rawInput = raw_input('enter a list of terms: ')
    tokens = rawInput.split(' ')
    print tokens


    workbook = load_workbook(docs[0])
    sheets = workbook.sheetnames
    worksheet = workbook[sheets[0]]

    #new workbook to save results to
    outputBook = Workbook()
    ws = outputBook.active
    c = 1
    oc = ws.cell(row = 1, column = c)
    for tok in tokens:
        oc.value = tok
        c += 1
        oc = ws.cell(row = 1, column = c)



    row = 1
    col = 1
    d = worksheet.cell(row = row, column = col)

    while d.value:
        for i, tok in enumerate(tokens):
            print tok
            if tok in str(d.value).lower():
                getRows(row, col, i+1, worksheet, ws)
                print "done"
        col += 1
        d = worksheet.cell(row = row, column = col)


    outputBook.save('results.xlsx')


if __name__ == "__main__":
    main()
