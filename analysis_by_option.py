#!/Library/Frameworks/Python.framework/Versions/3.5/bin/python3
from fileAnalyzer import FileAnalyzer
from tokenizer import Tokenizer
from summerize import Summerize
import optparse


def main():
    theAnalyzer = FileAnalyzer()
    theTokenizer = Tokenizer()
    theSummerizer = Summerize()
    # parser = optparse.OptionParser(usage = "%prog [options] [file_name or folder_name]\n")
    parser = optparse.OptionParser(usage = "%prog [options] [file_name]\n")

    parser.add_option("-f", "--file",
                      action="store_true",
                      default=False,
                      dest="single_file",
                      help="分析單獨excel檔")

    """
    parser.add_option("-d", "--directory",
                      action="store_true",
                      default=False,
                      dest="all_in_directory",
                      help="分析資料夾裡所有excel檔")
                      
    parser.add_option("-l", "--list",
                      action="store_true",
                      default=False,
                      dest="list_all",
                      help="列出所有excel及sheet名")

    parser.add_option("-c", "--combinedFiles",
                      action="store_true",
                      default=False,
                      dest="combined_Files",
                      help="綜合分析指定excel檔")

    """

    (options, args) = parser.parse_args()

    if options.single_file:
        if len(args) < 1:
            print("not enough args")
            parser.print_help()
            return -1
        else:
            
            for i in range(len(args)):
                # analyzer.print_out_result(analyzer.read_single_file(args[i]))
                # analyzer.parse_row_by_row(args[i])
                theAnalyzer.parse_excel_file(args[i])
            
            theTokenizer.cut_and_save_2_sqlite()
            
            theSummerizer.list_same_values()
            print("分析完畢，已將分析結果存入lexicon.db中的表analysis_result")
    else:
        parser.print_help()

    """

    elif options.all_in_directory:
        if len(args) < 1:
            print("not enough args")
            parser.print_help()
            return -1
        else:
            for i in range(len(args)):
               analyzer.go_through_directory(args[i])
                
    elif options.combined_Files:
        analyzer.printFilename = False
        all_data = dict()
        for i in range(len(args)):
            single_file_data = analyzer.read_single_file(args[i])
            # print(single_file_data)
            
            for ele in single_file_data:
                sn = ele[0]
                tokens_in_sheet = ""  # sheet為單位
                for theKey in ele[1].keys():
                    loopTime = int(theKey)
                    wannaStr = ""
                    for ti in range(loopTime):
                        wannaStr = ",".join([wannaStr, ele[1][theKey]])
                    wannaStr = wannaStr.lstrip(",").strip() #分次字數統計
                    
                    tokens_in_sheet = ",".join([tokens_in_sheet, wannaStr])
                tokens_in_sheet = tokens_in_sheet.lstrip(",").strip()
                    
                # print(sn, ":", tokens_in_sheet)
                
                if sn not in all_data:
                    all_data[sn] = tokens_in_sheet
                else:
                    ori_sn_data = all_data[sn]
                    all_data[sn] = ",".join([ori_sn_data, tokens_in_sheet])

        for k, v in all_data.items():
            cleanV = v.replace(","," ")
            count_result = analyzer.count_frequency(cleanV)
            print(k)
            # print(count_result)

            od = collections.OrderedDict(sorted(count_result.items()))
            print("字頻統計: ")
            for od_k, od_v in od.items():
                print(str(od_k) + "次: " + od_v)
            print("-----------------------------------------------------")


    
    elif options.list_all:
        if len(args) < 1:
            print("not enough args")
            parser.print_help()
            return -1
        else:
            for i in range(len(args)):
                analyzer.list_all(args[i])
    """
    

        
if __name__ == "__main__":
    main()
