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
    
