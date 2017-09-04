import jieba
import jieba.posseg as pseg
from sqlite_manager import SQLiteManager

class Tokenizer(SQLiteManager):
    def __init__(self):
        super().__init__()
        jieba.set_dictionary('./extra_dict/dict.txt.big')
        jieba.load_userdict('lexicon_dict.text')

    #回傳中文
    def replace_nonChinese(self, text):
        newStrList = []
        for ele in text:
            if '\u4e00' <= ele <= '\u9fff':
                newStrList.append(ele)
            else:
                pass
        return "".join(newStrList)


    def cut_with_speech_part(self, origin_string):
        origin_string = self.replace_nonChinese(origin_string)

        result_list = []
        words = pseg.cut(origin_string)

        for word, flag in words:
            result_list.append((word, flag))

        return result_list



    def cut_and_save_2_sqlite(self):

        #先確認sqlite table
        def create_task_table(task_table_name):
                sql_query = "CREATE TABLE IF NOT EXISTS %s (doc_name TEXT, sheet_name TEXT, task_id INT, keyword TEXT, part_of_speech TEXT);" % (task_table_name)
                self.execute(sql_query)
                sql_query = "DELETE FROM %s;" % (task_table_name)
                self.execute(sql_query)

        def insert_task_table(task_table_name, valueList):
                sql_query = "INSERT INTO %s VALUES ('%s', '%s', '%s', '%s', '%s');" % (task_table_name, valueList[0], valueList[1], valueList[2], valueList[3], valueList[4])
                self.execute(sql_query)

        create_task_table("task_cut_result")

        sql_query = "SELECT doc_name, sheet_name, task_id , task_desc FROM task_CRM_1707_業助;"
        queryResult = self.execute(sql_query)

        for ele in queryResult[0]:
            filename = ele[0]
            sheetname = ele[1]
            task_id = ele[2]
            task = ele[3]
            cut_results = self.cut_with_speech_part(task)
            for piece in cut_results:
                if len(piece[0])>1:
                    keyword = piece[0]
                    part_of_speech = piece[1]
                    insert_task_table("task_cut_result", [filename, sheetname,task_id, keyword, part_of_speech])