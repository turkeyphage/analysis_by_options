#!/Library/Frameworks/Python.framework/Versions/3.5/bin/python3

from sqlite_manager import SQLiteManager
import re
from datetime import datetime



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


sql_query = "CREATE TABLE IF NOT EXISTS %s (name TEXT, workdate datetime, count text);" % ('task_analysis')
sqlManager.execute(sql_query)
sql_query = "DELETE FROM %s;" % ('task_analysis')
sqlManager.execute(sql_query)




# partOfQueryStr = ""
# for ele in taskLexicons:
#     lexicon_with_type = ele + " INT"
#     partOfQueryStr = ", ".join([partOfQueryStr, lexicon_with_type]).strip(", ")

# print(partOfQueryStr)

#每一行的全文
for ele in sqlManager.execute("SELECT task_full_text FROM CRM_1707_業助_task;"):
    #task_full_text
    for ele_in_tuple in ele:
        full_text = ele_in_tuple[0]

        #以人作為判斷
        for name in employeeList:
            appear_found_position = [m.start() for m in re.finditer(name, full_text)] #出現位置判斷
            if len(appear_found_position)>0:
                #有找到
                #分析日期
                dateRegex = re.compile(r'\d+/\d+/\d+')
                if dateRegex.findall(full_text)[0]:
                    workdate = dateRegex.findall(full_text)[0]

                    workdate = datetime.strptime(workdate, '%Y/%m/%d')
                    # print(datetime_object)
                    # datetime.strptime('Jun 1 2005  1:33PM', '%b %d %Y %I:%M%p')



                    if workdate not in parsingResult[name]:
                        parsingResult[name][workdate] = ""
                    originWorkTask = parsingResult[name][workdate]
                    nowWorkTask = "\n".join([originWorkTask, full_text]).strip("\n")
                    parsingResult[name][workdate] = nowWorkTask



for k, v in parsingResult.items():
    #k = 員工姓名
    #v = {日期：工作內容}
    print(k, ":")
    for sortedDate in sorted(v, key=v.get, reverse=False):
        print("<", sortedDate, ">")        
        inV = v[sortedDate]
        taskCountDic = dict()
        for task in taskLexicons:
            taskFoundTimes = len([m.start() for m in re.finditer(task, inV)])
            if taskFoundTimes > 0:
                if task not in taskCountDic:
                    taskCountDic[task] = 0
                originalCount = taskCountDic[task]
                taskCountDic[task] = originalCount+taskFoundTimes

        #sort by frequency
        countString = ""

        for w in sorted(taskCountDic, key=taskCountDic.get, reverse=True):
            print("\t", w, "\t", str(taskCountDic[w]), "次")
            taskDesc = w + "\t" + str(taskCountDic[w])+"次"
            countString = "\n".join([countString,taskDesc]).strip("\n")
        
        sql_query =  "INSERT INTO %s VALUES ('%s', '%s', '%s');" % ('task_analysis', k, str(sortedDate), countString)
        sqlManager.execute(sql_query)

        # for w in sorted(taskCountDic, key=taskCountDic.get, reverse=True):
        #     print("\t", w, "\t", str(taskCountDic[w]), "次")


    
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

'''
#print parsingResult
for k, v in parsingResult.items():
    print(k, ":")
    for inK, inV in v.items():
        print(inK)
        print(inV)

    print("---------------------------------")
        
'''
