from sqlite_manager import SQLiteManager


class Summerize(SQLiteManager):

    def __init__(self):
        super().__init__()


    def list_same_values(self):

        return_value_total_rows = []
        max_length = 0
        max_length_item = []
        sql_query = "select * from task_CRM_1707_業助 order by task_start_date;"
        for ele in self.execute(sql_query)[0]:
            task_id = ele[2]
            employee = ele[3]
            workdate = ele[4]
            start_time = ele[5]
            task_desc = ele[11]

            keywordlist = []
            sql_query = "select task_id, keyword, part_of_speech from task_cut_result where task_id = %s;" % (task_id)
            for ele_taskcut in self.execute(sql_query)[0]:
                if ele_taskcut[1] not in keywordlist:
                    keywordlist.append(ele_taskcut[1])

            return_value_each_row = [workdate, start_time, employee, task_desc]

            for eachKeyword in keywordlist:
                occurrences = task_desc.count(eachKeyword)
                return_value_each_row.extend([eachKeyword, occurrences])

            if max_length <= len(return_value_each_row):
                max_length = len(return_value_each_row)

            return_value_total_rows.append(return_value_each_row)
        
        #save to table
        query1 = ""
        for i in range(1, int((max_length-4)/2) + 1):
            query1 = query1 + "keyword" + str(i) + " TEXT, frequency" + str(i) + " INT, "
            
        query1 = "date TEXT, time TEXT, name TEXT, work_detail TEXT, " + query1
        query1 = query1.rstrip(', ')
        sql_query = "CREATE TABLE IF NOT EXISTS analysis_result (" + query1 + ");" 
        self.execute(sql_query)
        sql_query = "DELETE FROM analysis_result;"
        self.execute(sql_query)

        
        for ele in return_value_total_rows:
            if len(ele) < max_length:
                for i in range(max_length - len(ele)):
                    ele.append(None)

        for ele in return_value_total_rows:  
            sql_query = "INSERT INTO analysis_result VALUES ("
            for in_ele in ele:
                if isinstance(in_ele, str):
                    sql_query = sql_query + "'" + in_ele + "',"
                elif isinstance(in_ele, int):
                    sql_query = sql_query + str(in_ele) + ","
                else:
                    sql_query = sql_query + 'null' + ","
            sql_query = sql_query.rstrip(",")
            sql_query = sql_query + ");"
            self.execute(sql_query)
        