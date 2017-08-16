#!/Library/Frameworks/Python.framework/Versions/3.5/bin/python3
import os
import os.path
import sys
import xlrd as xl
import openpyxl as pyxl
import errno
import re
import time
import platform
import datetime
from datetime import time
import thulac
import operator
import optparse
import collections


def creation_date(path_to_file):

    if platform.system() == 'Windows':
        return os.path.getctime(path_to_file)
    else:
        stat = os.stat(path_to_file)
        try:
            full_datetime = datetime.datetime.fromtimestamp(stat.st_birthtime)
            year = full_datetime.year
            month = full_datetime.month
            day = full_datetime.day
            return "-".join([str(year), str(month), str(day)])
        except AttributeError:
            # We're probably on Linux. No easy way to get creation dates here,
            # so we'll settle for when its content was last modified.
            full_datetime = datetime.datetime.fromtimestamp(stat.st_mtime)
            year = full_datetime.year
            month = full_datetime.month
            day = full_datetime.day
            return "-".join([str(year), str(month), str(day)])


def count_frequency(full_text):

    split_text = full_text.split(" ")

    feq_count = dict()

    for each in split_text:
        if each not in feq_count:
            feq_count[each] = 0

        feq_count[each] = feq_count[each] + 1

    feq_combined = dict()
    for w in sorted(feq_count, key=feq_count.get, reverse=True):
        if feq_count[w] not in feq_combined:
            feq_combined[feq_count[w]] = ""
        feq_combined[feq_count[w]] = ",".join([feq_combined[feq_count[w]], w]).lstrip(',')

    return feq_combined


def print_out_result(sheet_name, text_cut):
    filted_str = ""
    for ele in text_cut:
        if ele[-1] != 'u' and ele[-1] != 'y' and ele[-1] != 'r':
            filted_str = " ".join([filted_str, ele[0]])
    text_cut = filted_str.strip()
    freq_result = count_frequency(text_cut)
    od = collections.OrderedDict(sorted(freq_result.items()))
    print("sheet_name: " + sheet_name)
    print("字頻統計: ")

    for k, v in od.items():
        print(str(k) + "次: " + v)

    # for ele in freq_result.keys():
    #     print(str(ele)+"次: "+freq_result[ele])
    print("-----------------------------------")




#單一檔案讀取
def read_single_file(filepath):
    fileDic = dict()
    # thu1 = thulac.thulac(seg_only=True, filt=True)
    # # thu1 = thulac.thulac(seg_only=True)
    thu1 = thulac.thulac(filt=True)
    filename = os.path.basename(filepath)

    if filename.endswith(".xls"):
        # 使用xlrd

        filename_without_extension = filename.replace(".xls", "").strip()
        # fileDic[filename_without_extension] = []

        excelFile = xl.open_workbook(filepath)
        # 所有sheet的名字
        sheet_names = excelFile.sheet_names()

        # sheet iteration
        for i in range(len(sheet_names)):
            # sheet
            work_sheet = excelFile.sheet_by_index(i)

            # 以row數來判斷sheet是否為Empty
            if work_sheet.nrows != 0:
                # print("%s - %s: %d rows" %(filename, sheet_names[i], work_sheet.nrows))
                # file_create_date = creation_date(src)
                sheet_fn = sheet_names[i].strip()

                sheet_contents = ""
                num_rows = work_sheet.nrows - 1
                num_cells = work_sheet.ncols - 1
                curr_row = -1

                while curr_row < num_rows:
                    curr_row += 1
                    # row = work_sheet.row(curr_row)
                    curr_cell = -1

                    while curr_cell < num_cells:
                        curr_cell += 1
                        # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
                        cell_type = work_sheet.cell_type(curr_row, curr_cell)
                        cell_value = work_sheet.cell_value(curr_row, curr_cell)

                        try:
                            if cell_type == 3:
                                cell_value = int(cell_value * 24 * 3600)
                                cell_value = time(cell_value // 3600, (cell_value % 3600) // 60, cell_value % 60)
                        except:
                            cell_value = ""

                        if cell_type != 0 and cell_type != 6 and cell_type != 5:
                            cell_value = re.sub("[\s\d，,<>;：（）().；\-:、/.。《》…~\"\\\%$&#*@~`？]+", "", str(cell_value))
                            sheet_contents = " ".join(
                                [sheet_contents, str(cell_value)]).strip()

                # text_cut = thu1.cut(sheet_contents, text=True).strip()

                text_cut = thu1.cut(sheet_contents)
                print_out_result(sheet_names[i], text_cut)


    elif filename.endswith(".xlsx"):
        # 使用openpyxl

        filename_without_extension = filename.replace(".xlsx", "").strip()
        # fileDic[filename_without_extension] = []

        excelFile = pyxl.load_workbook(filepath)
        #所有sheet的名字
        sheet_names = excelFile.get_sheet_names()

        #sheet iteration
        for i in range(len(sheet_names)):
            #sheet
            work_sheet = excelFile.get_sheet_by_name(sheet_names[i])

            num_rows = work_sheet.max_row
            num_columns = work_sheet.max_column

            if num_columns != 1 and num_rows != 1:
                # not empty
                sheet_contents = ""

                for curr_row in range(1, num_rows + 1):
                    for curr_cell in range(1, num_columns + 1):
                        cell_value = work_sheet.cell(
                            row=curr_row, column=curr_cell).value
                        if cell_value == None:
                            cell_value = ""

                        cell_value = re.sub("[\s\d，,<>;：（）().；\-:、/.。《》…~\"\\\%$&#*@~`？]+", "", str(cell_value))
                        sheet_contents = " ".join(
                            [sheet_contents, str(cell_value)]).strip()

                
                # text_cut = thu1.cut(sheet_contents, text=True).strip()
                text_cut = thu1.cut(sheet_contents)
                print_out_result(sheet_names[i], text_cut)

    else:
        # 非excel檔
        return


def go_through_directory(dir_path):
    os.chdir(dir_path)
    print(os.getcwd())
    for roots, dirs, files in os.walk(os.getcwd()):
        for filename in files:
            src = os.path.join(roots, filename)
            read_single_file(src)




def main():

    parser = optparse.OptionParser(usage = "%prog [options] [file_name or folder_name]\n")

    parser.add_option("-f", "--file",
                      action="store_true",
                      default=False,
                      dest="single_file",
                      help="分析單獨excel檔")

    parser.add_option("-d", "--directory",
                      action="store_true",
                      default=False,
                      dest="all_in_directory",
                      help="分析資料夾裡所有excel檔")

    parser.add_option("-c", "--combinedFiles",
                      action="store_true",
                      default=False,
                      dest="combined_Files",
                      help="綜合分析指定excel檔")


    (options, args) = parser.parse_args()

    if len(args) < 1:
        print("not enough args")
        parser.print_help()    
        return -1


    else:
        
        if options.single_file:
            for i in range(len(args)):
                read_single_file(args[i])
        
                
        elif options.all_in_directory:
            for i in range(len(args)):
                go_through_directory(args[i])

        elif options.combined_Files:
            print("sorry, this function is under construction, please wait")

                
        else:
            parser.print_help()









if __name__ == "__main__":
    main()
