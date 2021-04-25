import openpyxl
import os
import sys

basepath = ""
txt_file_folder = "txtfiles"

def read_xlsx():
    # reading file
    wb = openpyxl.load_workbook('test.xlsx')
    ws = wb.active
    write_txt_file(ws.cell(2,1).value, ws.cell(2,2).value)
    print("reading finished")


def write_txt_file(name , content):
    # write the file
    path = str('./' + txt_file_folder)
    if not os.path.isdir(path):
        os.mkdir(txt_file_folder)
    #
    filename = txt_file_folder+  "/" + name + ".txt"
    file1 = open(filename,"w+")
    file1.write(content)
    file1.close()
    print("write txt file finished")

if __name__ == "__main__":
    basepath = sys.args[1]
    read_xlsx()