#!/Library/Frameworks/Python.framework/Versions/3.5/bin/python3

from sqlite_manager import SQLiteManager
import re
from datetime import datetime
import collections

sqlManager = SQLiteManager()
parsingResult = dict() #最後的結果

#員工姓名
employeeList = []

for ele in sqlManager.execute("SELECT keyword FROM lexical where word_type = 1;"):
    for eachResult in ele :
        employeeList.append(eachResult[0])

for name in employeeList:
    if name not in parsingResult:
        parsingResult[name] = dict()

#詞庫
taskLexicons = []
for ele in sqlManager.execute("SELECT keyword FROM lexical where word_type = 3;"):
    for eachResult in ele:
        taskLexicons.append(eachResult[0])
# print(taskLexicons)

sql_query = "CREATE TABLE IF NOT EXISTS %s (doc_name TEXT, sheet_name TEXT, emp_name TEXT, workdate datetime, keyword TEXT, row_id INT, full_text TEXT);" % ('task_analysis')
sqlManager.execute(sql_query)
sql_query = "DELETE FROM %s;" % ('task_analysis')
sqlManager.execute(sql_query)

# partOfQueryStr = ""
# for ele in taskLexicons:
#     lexicon_with_type = ele + " INT"
#     partOfQueryStr = ", ".join([partOfQueryStr, lexicon_with_type]).strip(", ")

# print(partOfQueryStr)

#每一行的全文
for ele in sqlManager.execute("SELECT doc_name, sheet_name, task_id, task_full_text FROM CRM_1707_業助_task;"):
    #task_full_text
    for ele_in_tuple in ele:
        # print(ele_in_tuple)
        doc_name = ele_in_tuple[0]
        sheet_name = ele_in_tuple[1]
        row_id = ele_in_tuple[2]
        full_text = ele_in_tuple[3]

        if doc_name not in parsingResult:
            parsingResult[doc_name] = dict()
        
        if sheet_name not in parsingResult[doc_name]:
            parsingResult[doc_name][sheet_name] = dict()

        #以人作為判斷
        for name in employeeList:
            appear_found_position = [m.start() for m in re.finditer(name, full_text)] #出現位置判斷
            if len(appear_found_position)>0:
                #有找到
                if name not in parsingResult[doc_name][sheet_name]:
                    parsingResult[doc_name][sheet_name][name] = dict()
                #分析日期
                dateRegex = re.compile(r'\d+/\d+/\d+')
                if dateRegex.findall(full_text)[0]:
                    workdate = dateRegex.findall(full_text)[0]
                    workdate = datetime.strptime(workdate, '%Y/%m/%d')
                    if workdate not in parsingResult[doc_name][sheet_name][name]:
                        # parsingResult[name][workdate] = ""
                        parsingResult[doc_name][sheet_name][name][workdate] = dict()

                    for lexicon in taskLexicons:
                        keywordFoundTimes = len([m.start() for m in re.finditer(lexicon, full_text)])
                        if keywordFoundTimes > 0:
                            for i in range(keywordFoundTimes):
                                if lexicon not in parsingResult[doc_name][sheet_name][name][workdate]:
                                    parsingResult[doc_name][sheet_name][name][workdate][lexicon] = dict()
                                if row_id not in parsingResult[doc_name][sheet_name][name][workdate][lexicon]:
                                    parsingResult[doc_name][sheet_name][name][workdate][lexicon][row_id] = []
                                
                                oriData = parsingResult[doc_name][sheet_name][name][workdate][lexicon][row_id]
                                oriData.append(full_text)
                                parsingResult[doc_name][sheet_name][name][workdate][lexicon][row_id] = oriData


        #             originWorkTask = parsingResult[name][workdate]
        #             nowWorkTask = "\n".join([originWorkTask, full_text]).strip("\n")
        #             parsingResult[name][workdate] = nowWorkTask


for k, v in parsingResult.items():
    #k = doc_name
    #v = {sheet_name: 
    #       {員工姓名：
    #           {日期：
    #               {keyword: {row_id : [full_text1, full_text2...] } } } }
    documentName = k
    for k1, v1 in v.items():
        sheet_name = k1
        for k2, v2 in v1.items():
            #k2 = 員工姓名
            #v = {日期：{keyword: {row_id : [full_text1, full_text2...] } }
            employeeName = k2
            sortedDateTime = collections.OrderedDict(sorted(v2.items()))
            for theDate in sortedDateTime:
                #theDate = 日期
                inV = v2[theDate]  #{keyword: {row_id : [full_text1, full_text2...] }
                for keyword, rowAndFull in inV.items():
                    #keyword詞庫的字
                    for the_row_id in sorted(rowAndFull, key=rowAndFull.get, reverse=False):
                        all_full_texts = rowAndFull[the_row_id] #list
                        for each_full_text in all_full_texts:
                            #寫進table
                            sql_query =  "INSERT INTO %s VALUES ('%s', '%s', '%s', '%s', '%s', %d, '%s');" % ('task_analysis', documentName, sheet_name, employeeName, str(theDate), keyword, the_row_id, each_full_text )
                            sqlManager.execute(sql_query)

print("分析完畢，已將分析結果存入lexicon.db, table: ", 'task_analysis')
    # for sortedDate in sorted(v, key=v.get, reverse=False):
    #     #sortedFate = 日期
    #     inV = v[sortedDate]  #{keyword: {row_id : [full_text1, full_text2...] }

    #     for keyword, rowAndFull in inV.items():
    #         #keyword詞庫的字
    #         for the_row_id in sorted(rowAndFull, key=rowAndFull.get, reverse=False):
    #             all_full_texts = rowAndFull[the_row_id] #list

    #             for each_full_text in all_full_texts:
    #                 #寫進table
    #                 sql_query =  "INSERT INTO %s VALUES ('%s', '%s', '%s', %d, '%s');" % ('task_analysis', employeeName, str(sortedDate), keyword, the_row_id, each_full_text )
    #                 sqlManager.execute(sql_query)



#暫時不動這一段
# for k, v in parsingResult.items():
#     #k = 員工姓名
#     #v = {日期：工作內容}
#     print(k, ":")
#     for sortedDate in sorted(v, key=v.get, reverse=False):
#         print("<", sortedDate, ">")        
#         inV = v[sortedDate]
#         taskCountDic = dict()
#         for task in taskLexicons:
#             taskFoundTimes = len([m.start() for m in re.finditer(task, inV)])
#             if taskFoundTimes > 0:
#                 if task not in taskCountDic:
#                     taskCountDic[task] = 0
#                 originalCount = taskCountDic[task]
#                 taskCountDic[task] = originalCount+taskFoundTimes

#         #sort by frequency
#         countString = ""

#         for w in sorted(taskCountDic, key=taskCountDic.get, reverse=True):
#             print("\t", w, "\t", str(taskCountDic[w]), "次")
#             taskDesc = w + "\t" + str(taskCountDic[w])+"次"
#             countString = "\n".join([countString,taskDesc]).strip("\n")
        
#         sql_query =  "INSERT INTO %s VALUES ('%s', '%s', '%s');" % ('task_analysis', k, str(sortedDate), countString)
#         sqlManager.execute(sql_query)

#         # for w in sorted(taskCountDic, key=taskCountDic.get, reverse=True):
#         #     print("\t", w, "\t", str(taskCountDic[w]), "次")



    # for inK, inV in v.items():
    #     #inK = 日期
    #     #inV = 工作內容
    #     print(inK)
    #     taskCountDic = dict()
    #     for task in taskLexicons:
    #         taskFoundTimes = len([m.start() for m in re.finditer(task, inV)])
    #         if taskFoundTimes > 0:
    #             if task not in taskCountDic:
    #                 taskCountDic[task] = 0
    #             originalCount = taskCountDic[task]
    #             taskCountDic[task] = originalCount+taskFoundTimes
        
    #     #sort by frequency
    #     for w in sorted(taskCountDic, key=taskCountDic.get, reverse=True):
    #         print("\t", w, "\t", str(taskCountDic[w]), "次")


        # for taskcountK, taskcountV in taskCountDic.items():
        #     print("\t", taskcountK, "\t", str(taskcountV), "次")


#print parsingResult
# for k, v in parsingResult.items():
#     print(k, ":")
#     for inK, inV in v.items():
#         print(inK)
#         print(inV)

#     print("---------------------------------")
        

