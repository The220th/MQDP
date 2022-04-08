# -*- coding: utf-8 -*- 

import os

from docxparsek import Doc
from docxparsek import Line
from docxparsek import Text
from docxparsek import Run
from docxparsek import Image
from docxparsek import Table
from docxparsek import Row
from docxparsek import Cell

def standardk_run(docPath : str, outPath : str) -> tuple:
    k = standardk_process(docPath, outPath)
    return k.k_run()


# Code below like: https://i.imgur.com/W5AxL6N.jpg


class standardk_process:

    __docPath = None
    __outPath = None
    __imgFolder = None
    __relImgFolder = "imgs"

    __debug_file = None

    __debug = ""

    __lastError = None

    __image_i = None

    __DEBUG_ON = None

    def __init__(self, docPath : str, outPath: str):
        if "MQPD_DEBUG_ON" in os.environ:
            self.__DEBUG_ON = bool(os.environ["MQPD_DEBUG_ON"])
        else:
            self.__DEBUG_ON = False

        self.__docPath = docPath
        self.__outPath = outPath
        self.__debug_file = os.path.join(self.__outPath, "bebug_standartk.txt")
        #self.__imgFolder = os.path.join(self.__outPath, "imgs")
        self.__imgFolder = os.path.join(self.__outPath, self.__relImgFolder)

        self.__image_i = 0
    
    def k_run(self) -> tuple:
        try:
            self.debug(f"docPath = {self.__docPath}, outDir = {self.__outPath}, imgdir = {self.__imgFolder}")

            self.debug(f"openning Doc: {self.__docPath}")
            doc = Doc(self.__docPath)
            self.debug(f"opened Doc: {self.__docPath}")

            self.debug(f"Creating folder: {self.__imgFolder}")
            os.makedirs(self.__imgFolder)
            self.debug(f"Created folder: {self.__imgFolder}")

            TABLEFINDED = False
            for line in doc:
                if(line.isTable()):
                    table = line.getSrc()
                    for row in table:
                        if(row.getRowNum() > 0):
                            q = self.question_depo(row)
                            self.debug(f"\n\n==============================\n\n")
                    TABLEFINDED = True
                    break
                else:
                    #print(line.getSrc())
                    #self.__lastError = "Syntax error. Cannot find table"
                    #raise SyntaxError
                    pass
            if(TABLEFINDED == False):
                self.__lastError = "Syntax error. Cannot find table"
                raise SyntaxError

        except SyntaxError:
            return (self.__lastError, 4, self.__debug)

        return (f"Complete succesessfully", 0, self.__debug)
        #return (f"test: {self.__docPath}, {self.__outPath}", 2, self.__debug)

    def writeBytes(self, where : str, b : bytes):
        with open(where, 'wb') as temp:
            temp.write(b)

    def writeText(self, where : str, s : str):
        with open(where, 'w', encoding="utf-8") as temp:
            temp.write(s)

    def writeTextAppend(self, where : str, s : str):
        with open(where, 'a', encoding="utf-8") as temp:
            temp.write(s)

    '''
    def getRelativeImgPath(self, absPath : str):
        #res = os.path.relpath(self.__imgFolder, self.__outPath)
    '''
    
    def checkColorRight(self, c : str) -> bool:
        '''
        Green -> True
        Red -> False
        else -> None
        '''
        if(c == "auto"):
            return None
        r = int(c[0:2], 16)
        g = int(c[2:4], 16)
        b = int(c[4:6], 16)
        res = True
        #if(r < 50 and g > 100 and b < 50):
        if(g > r and g > b):
            return True
        #elif(r > 100 and g < 50 and b < 50):
        elif(r > g and r > b):
            return False
        else:
            return None

    def checkQuestionRight(self, text) -> bool:
        self.debug(f"Checking is \"{text.getText()}\" is correct...")
        self.debug(f"have=={text.getText().strip()[0] == '='}, bold={text.isBold()}, underline={text.isUnderline()}, color={text.isColored() and self.checkColorRight(text.getColor()) == True}")
        if(text.getText().strip()[0] == "="
        or text.isBold()
        or text.isUnderline() 
        or ( text.isColored() and self.checkColorRight(text.getColor()) == True )
        ):
            self.debug(f"So it is correct")
            return True
        else:
            self.debug(f"So it is incorrect")
            return False
        
    def getNumAnswers(self, cell) -> tuple:
        '''
        return = (allAnswers, rightAnswers)
        '''

        self.debug(f"Checking answers number...")

        ans = []
        rightsNum = 0
        cell_2 = cell
        f = True
        for line in cell_2:
            if(line.isOther()):
                if(f == False):
                    ans_i+=1
                    f = True
            else:
                if(f == True):
                    ans.append("") # len(ans) is current number of question
                if(line.isText()):
                    text = line.getSrc()
                    if(self.checkQuestionRight(text)): # right ans
                        rightsNum += 1
        self.debug(f"Number of all answers = {len(ans)}, number right answers = {rightsNum}")
        return (len(ans), rightsNum)

    def getImageLink(self, image : Image) -> str:
        imgName = f"image{self.__image_i}.png"
        pathToSave = os.path.join(self.__imgFolder, imgName)
        self.debug(f"Trying to save image {pathToSave}")
        self.writeBytes(pathToSave, image.getBytes())
        self.debug(f"{pathToSave} saved")
        res = f"<img src\\=\"@@PLUGINFILE@@/{self.__relImgFolder}/{imgName}\"/>"
        self.debug(f"img link generated = {res}")
        self.__image_i+=1
        return res

    def parse_by_del(self, s : str, begin : str, end : str) -> str:
        '''
        ex: s = "=%53%", begin = "=%", end = "%". Then return = "53"
        ex: s = "=%53%123%52%", begin = "=%", end = "%". Then return = "53"
        '''
        beginfind = s.find(begin)
        if(beginfind != -1):
            res = s[beginfind+len(begin):]
        else:
            return s
        
        endfind = res.find(end)
        if(endfind != -1):
            res = res[:endfind]

        return res

    def isRepresentsInt(self, s : str) -> bool:
        try: 
            int(s)
            return True
        except ValueError:
            return False

    def isRepresentsFloat(self, s : str) -> bool:
        try: 
            float(s)
            return True
        except ValueError:
            return False

    def getMarkdownStyleQuestion(self, cell) -> tuple:
        '''
        return = (question, comment)
        comment string begin with "//"
        '''
        cell_1 = cell

        Q = ""
        Comment = ""
        for line in cell_1:
            if(line.isText()):
                text = line.getSrc()
                if(text.getText().strip()[:2] == "//"):
                    Comment += text.getText()
                else:
                    Q += text.getText()
            elif(line.isImage()):
                img = line.getSrc()
                Q += " " + self.getImageLink(img) + " "
            elif(line.isOther()):
                Q += "\n"
        return (Q, Comment)

    def getMarkdownStyleLineAndImg(self, cell) -> list:
        ans = []
        ans_i = 0
        f = True
        rightsNum = 0
        cell_2 = cell

        for line in cell_2:
            if(line.isOther()):
                if(f == False):
                    ans_i+=1
                    f = True
            else:
                if(f == True):
                    ans.append("") # len(ans) is current number of question

                # O6PA6OTKA begin
                if(line.isText()):
                    text = line.getSrc()
                    ans[ans_i] += text.getText()
                if(line.isImage()):
                    img = line.getSrc()
                    ans[ans_i] += self.getImageLink(img)
                # O6PA6OTKA end

                f = False
        return ans

    def question_depo(self, row : Row) -> str:
        self.debug(f"trying to detect type of row {row.getRowNum()}")
        firstCell = row.getCell(0)
        qmark = "None type"
        for line in firstCell:
            if(line.isText()):
                text = line.getSrc().getText()
                #print(text)
                if(text[0] == 'О' or text[0] == 'М' or text[0] == 'К' or text[0] == 'Ф' or
                text[0] == 'С' or text[0] == 'Ч' or text[0] == 'Э' or
                text[0] == 'о' or text[0] == 'м' or text[0] == 'к' or text[0] == 'ф' or
                text[0] == 'с' or text[0] == 'ч' or text[0] == 'э'):
                    qmark = text[0]
                    break
            else:
                self.__lastError = f"Syntax error. Cannot define type of question in row {row.getRowNum()}"
                raise SyntaxError
        self.debug(f"type of row {row.getRowNum()} is {qmark}")
        
        res = ""
        if(qmark == 'О' or qmark == 'о'):
            res = self.question_OnePick(row)
        elif(qmark == 'М' or qmark == 'м'):
            res = self.question_MulPick(row)
        elif(qmark == 'К' or qmark == 'к'):
            res = self.question_ShortPick(row)
        elif(qmark == 'Ф' or qmark == 'ф'):
            res = self.question_50_50Pick(row)
        elif(qmark == 'С' or qmark == 'с'):
            res = self.question_comparisonPick(row)
        elif(qmark == 'Ч' or qmark == 'ч'):
            res = self.question_numericPick(row)
        elif(qmark == 'Э' or qmark == 'э'):
            res = self.question_superOpenPick(row)
        else:
            self.__lastError = f"Syntax error. Cannot define type of question in row {row.getRowNum()}"
            raise SyntaxError
        return res
    
    def question_OnePick(self, row : Row) -> str:
        cell_1 = row.getCell(1)

        Q = ""
        Comments = ""
        for line in cell_1:
            if(line.isText()):
                text = line.getSrc()
                if(text.getText().strip()[:2] == "//"):
                    Comments += text.getText()
                else:
                    Q += text.getText()
            elif(line.isImage()):
                img = line.getSrc()
                Q += " " + self.getImageLink(img) + " "
            elif(line.isOther()):
                Q += "\n"
        self.debug(f"Question {row.getRowNum()} formed: {Q}")

        cell_2 = row.getCell(2)
        ans = []
        ans_i = 0
        f = True
        rightsNum = 0

        for line in cell_2:
            if(line.isOther()):
                if(f == False):
                    ans_i+=1
                    f = True
            else:
                if(f == True):
                    ans.append("") # len(ans) is current number of question

                # O6PA6OTKA begin
                if(line.isText()):
                    text = line.getSrc()
                    #print(f"{123}: {text.getText()} {text.isBold()} {text.getColor()}")
                    if(self.checkQuestionRight(text)): # right ans
                        rightsNum += 1
                        #print(text.getText())
                        if(text.getText().strip()[0] != "="):
                            ans[ans_i] += "="
                        ans[ans_i] += text.getText()
                    else:
                        ans[ans_i] += text.getText()
                if(line.isImage()):
                    img = line.getSrc()
                    ans[ans_i] += self.getImageLink(img)
                # O6PA6OTKA end

                f = False
        
        if(rightsNum != 1):
            self.__lastError = f"Syntax error. In row {row.getRowNum()} must be only 1 correct answer"
            raise SyntaxError

        for ans_i in range(len(ans)):
            if(ans[ans_i][0] != "="):
                ans[ans_i] = "~" + ans[ans_i]


        res = Comments + "\n"
        res += f"::Вопрос {row.getRowNum()}::{Q}" + "{\n"
        for an in ans:
            res += an + "\n"
        res += "}"

        self.debug(f"answers formed: \n{res}")

        return res

    def mulQuestion_checkRightPercent(self, cell) -> tuple():
        '''
        if okay, errorStr = ""
        if not okay errorStr = error_msg

        okay is <all question not have %%> or <all question right determine %%>
                  (then percents = None)          (then percents = percents)

        return (errorStr, positivePercents, negativePercents)
        '''
        percents = []
        negativePercents = []

        ans = []
        cell_2 = cell
        f = True
        for line in cell_2:
            if(line.isOther()):
                if(f == False):
                    ans_i+=1
                    f = True
            else:
                if(f == True):
                    ans.append("") # len(ans) is current number of question
                if(line.isText()):
                    text = line.getSrc()
                    if(self.checkQuestionRight(text)):
                        if(text.getText().strip()[:2] == "=%"):
                            perS = text.getText().strip()
                            perS = self.parse_by_del(perS, "=%", "%")
                            if(self.isRepresentsInt(perS)):
                                percents.append(int(perS))
                            else:
                                return (f"In answer {ans_i+1} \"{perS}\" is not number. ", None, None)
                    else:
                        if(text.getText().strip()[:2] == "~%"):
                            perS = text.getText().strip()
                            perS = self.parse_by_del(perS, "~%", "%")
                            if(self.isRepresentsInt(perS)):
                                negativePercents.append(int(perS))
                            else:
                                return (f"In answer {ans_i+1} \"{perS}\" is not number. ", None, None)
        
        if(len(percents) + len(negativePercents) == 0):
            return ("", None, None) 
        if(len(ans) != len(percents) + len(negativePercents)):
            return (f"The weight of not all answers is determined. ", None, None)
        
        sum_positive = 0
        sum_negative = 0
        for i in percents:
            sum_positive += i
        for i in negativePercents:
            sum_negative += i

        if(sum_positive != 100):
            return ("The sum of percentages of correct answers is not equal to 100", None, None)
        if(sum_negative != -100):
            return ("The sum of percentages of incorrect answers is not equal to -100", None, None)

        
        return ("", percents, negativePercents)

    def calPercents(self, a : int, r : int) -> tuple:
        '''
        return = (list of correct answers percent, list of incorrect answers percent)
        '''
        self.debug(f"Calculating the percentages for answers={a}, where right={r}...")

        inr = a - r

        c = [100 // r for i in range(r)]
        c[len(c)-1] = c[len(c)-1] + (100 - (100 // r)*r)

        inc = [100 // inr for i in range(inr)]
        inc[len(inc)-1] = inc[len(inc)-1] + (100 - (100 // inr)*inr)
        for i in range(len(inc)):
            inc[i] = -inc[i]

        self.debug(f"Calculating the percentages for answers={a}, where right={r} done: correct={c}, incorrect={inc}")
        return (c, inc)

    def question_MulPick(self, row : Row) -> str:
        '''
        Если каждый новый правильный ответ начинается на =%,
        тогда используем проценты явным образом.

        Если Просто выделены правильные ответы,
        то подсчёт процентов вручную
        '''
        cell_1 = row.getCell(1)

        Q = ""
        Comments = ""
        for line in cell_1:
            if(line.isText()):
                text = line.getSrc()
                if(text.getText().strip()[:2] == "//"):
                    Comments += text.getText()
                else:
                    Q += text.getText()
            elif(line.isImage()):
                img = line.getSrc()
                Q += " " + self.getImageLink(img) + " "
            elif(line.isOther()):
                Q += "\n"
        self.debug(f"Question {row.getRowNum()} formed: {Q}")

        cell_2 = row.getCell(2)
        ans = []
        ans_i = 0
        f = True
        rightsNum = 0

        self.debug(f"Cheking percent type...")
        checkPercent = self.mulQuestion_checkRightPercent(cell_2)
        if(checkPercent[0] != ""):
            self.__lastError = f"Syntax error. In row {row.getRowNum()}: {checkPercent[0]}"
            raise SyntaxError
        percents_pos = []
        percents_pos_i = 0
        percents_neg = []
        percents_neg_i = 0
        if(checkPercent[1] == None):
            self.debug(f"The percentages are NOT set by the user")
            nums = self.getNumAnswers(cell_2)
            self.debug(f"Calculate the percentages manually...")
            percents_pos, percents_neg = self.calPercents(nums[0], nums[1])
            self.debug(f"Percentages calculated. len(Pos) = {len(percents_pos)}, len(Neg) = {len(percents_neg)}")
        else:
            self.debug(f"The percentages are set by the user")
            percents_pos = checkPercent[1]
            percents_neg = checkPercent[2]
            self.debug(f"Percentages calculated. len(Pos) = {len(percents_pos)}, len(Neg) = {len(percents_neg)}")

        for line in cell_2:
            if(line.isOther()):
                if(f == False):
                    ans_i+=1
                    f = True
            else:
                if(f == True):
                    ans.append("") # len(ans) is current number of question

                # O6PA6OTKA begin
                if(line.isText()):
                    text = line.getSrc()
                    #print(f"{123}: {text.getText()} {text.isBold()} {text.getColor()}")
                    if(self.checkQuestionRight(text)): # right ans
                        if(text.getText().strip()[:2] == "=%"):
                            ans[ans_i] += text.getText()
                        else:
                            ans[ans_i] += f"=%{percents_pos[percents_pos_i]}%" + text.getText()
                            percents_pos_i-=-1
                        #print(text.getText())
                    else:
                        if(text.getText().strip()[:2] == "~%"):
                            ans[ans_i] += text.getText()
                        else:
                            ans[ans_i] += f"=%{percents_neg[percents_neg_i]}%" + text.getText()
                            percents_neg_i-=-1
                if(line.isImage()):
                    img = line.getSrc()
                    ans[ans_i] += self.getImageLink(img)
                # O6PA6OTKA end

                f = False

        res = Comments + "\n"
        res += f"::Вопрос {row.getRowNum()}::{Q}" + "{\n"
        for an in ans:
            res += an + "\n"
        res += "}"

        self.debug(f"answers formed: \n{res}")

        return res
        
    def question_ShortPick(self, row):
        '''
        Если каждый новый ответ начинается на =%,
        тогда используем проценты явным образом.

        Иначе только один ответ
        '''

        cell_1 = row.getCell(1)
        Q, Comment = self.getMarkdownStyleQuestion(cell_1)
        self.debug(f"Question {row.getRowNum()} formed: {Q}")

        cell_2 = row.getCell(2)
        answers = self.getMarkdownStyleLineAndImg(cell_2)

        for line in cell_2:
            if(line.isImage()):
                self.__lastError = f"In answers of question {row.getRowNum()} cannot be images. "
                raise SyntaxError

        MANYPICKS = False
        for an in answers:
            if(an.strip()[:2] == "=%"):
                MANYPICKS = True
                break

        self.debug(f"In question {row.getRowNum()}: MANYPICKS={MANYPICKS}\n")
        
        if(MANYPICKS == True):
            for an in answers:
                something = self.parse_by_del(an, "=%", "%")
                if(not self.isRepresentsInt(something)):
                    self.__lastError = f"The weight of not all answers is determined in question {row.getRowNum()}. "
                    raise SyntaxError
                elif(int(something) < 0 or int(something) > 100):
                    self.__lastError = f"The weight \"{int(something)}\" of question {row.getRowNum()} is not determined correctly. "
                    raise SyntaxError

        ans = []
        ans_i = 0
        f = True

        for an in answers:
            ans.append("")
            if(MANYPICKS == False):
                ans[ans_i] += "="
            ans[ans_i] += an
            ans_i+=1

        if(MANYPICKS == False):
            if(len(ans) != 1):
                self.__lastError = f"Too many answers in question {row.getRowNum()}. "
                raise SyntaxError

        res = Comment + "\n"
        res += f"::Вопрос {row.getRowNum()}::{Q}" + "{\n"
        for an in ans:
            res += an + "\n"
        res += "}"

        self.debug(f"answers formed: \n{res}")

        return res

    def question_50_50Pick(self, row):
        '''
        Только один ответ

        Правильный = Верно, верно, да, Да, 1
        Неправильный = Неверно, неверно, нет, Нет, 0
        '''

        cell_1 = row.getCell(1)
        Q, Comment = self.getMarkdownStyleQuestion(cell_1)
        self.debug(f"Question {row.getRowNum()} formed: {Q}")

        cell_2 = row.getCell(2)
        answers = self.getMarkdownStyleLineAndImg(cell_2)
        if(len(answers) != 1):
            self.__lastError = f"Too many or no answers in question {row.getRowNum()}. "
            raise SyntaxError

        for line in cell_2:
            if(line.isImage()):
                self.__lastError = f"In answers of question {row.getRowNum()} cannot be images. "
                raise SyntaxError

        self.debug(f"User\'s answer: {answers[0]}")

        verdict = ""
        if(answers[0].strip()[0] == 'В'
        or answers[0].strip()[0] == 'в'
        or answers[0].strip()[0] == 'Д'
        or answers[0].strip()[0] == 'д'
        or answers[0].strip()[0] == '1'
        ):
            verdict = "TRUE"
        elif(answers[0].strip()[0] == 'Н'
        or answers[0].strip()[0] == 'н'
        #or answers[0].strip()[0] == 'Н' #     =/
        #or answers[0].strip()[0] == 'н' # ¯\_(ツ)_/¯ 
        or answers[0].strip()[0] == '0'
        ):
            verdict = "FALSE"
        else:
            self.__lastError = f"In question {row.getRowNum()} the answer is incorrectly defined. "
            raise SyntaxError

        res = Comment + "\n"
        res += f"::Вопрос {row.getRowNum()}::{Q}" + "\n{"
        res += verdict
        res += "}"

        self.debug(f"answers formed: \n{res}")

        return res

    def question_comparisonPick(self, row):
        '''
        Минимум 3 сопоставления
        a1 = b1
        a2 = b2
        a3 = b3
        '''
        cell_1 = row.getCell(1)
        Q, Comment = self.getMarkdownStyleQuestion(cell_1)
        self.debug(f"Question {row.getRowNum()} formed: {Q}")

        cell_2 = row.getCell(2)
        answers = self.getMarkdownStyleLineAndImg(cell_2)
        if(len(answers) < 3):
            self.__lastError = f"Too few or no answers in question {row.getRowNum()}. There should be at least 3 answers. "
            raise SyntaxError

        for line in cell_2:
            if(line.isImage()):
                self.__lastError = f"In answers of question {row.getRowNum()} cannot be images. "
                raise SyntaxError 

        for an in answers:
            if(an.find('=') == -1):
                self.__lastError = f"In answer \"{an}\" of question {row.getRowNum()} no matching sign \"=\". "
                raise SyntaxError 

        ans = []
        ans_i = 0

        for an in answers:
            ans.append("")

            ans[ans_i] += "="
            match_sign_i = an.find('=')
            match_l = an[:match_sign_i]
            match_r = an[match_sign_i+1:]
            if(match_l.strip() == "" or match_r.strip() == ""):
                self.__lastError = f"In answer \"{an}\" of question {row.getRowNum()} no comparison. "
                raise SyntaxError
            ans[ans_i] += match_l + " -> " + match_r + "\n"

            ans_i+=1

        res = Comment + "\n"
        res += f"::Вопрос {row.getRowNum()}::{Q}" + "{\n"
        for an in ans:
            res += an + "\n"
        res += "}"

        self.debug(f"answers formed: \n{res}")

        return res

    def question_numericPick(self, row):
        '''
        replace " " -> ""
        replace "," -> "."

        3.14 then 3.14
        or
        3.141..3.142 then 3.141..3.142
        or
        3.141...3.142 then 3.141..3.142
        or
        3.1415%0.0005 then 3.1415:0.0005
        '''

        cell_1 = row.getCell(1)
        Q, Comment = self.getMarkdownStyleQuestion(cell_1)
        self.debug(f"Question {row.getRowNum()} formed: {Q}")

        cell_2 = row.getCell(2)
        answers = self.getMarkdownStyleLineAndImg(cell_2)
        if(len(answers) != 1):
            self.__lastError = f"Too many or no answers in question {row.getRowNum()}. "
            raise SyntaxError

        for line in cell_2:
            if(line.isImage()):
                self.__lastError = f"In answers of question {row.getRowNum()} cannot be images. "
                raise SyntaxError
        
        self.debug(f"User\'s answer: {answers[0]}")

        all_num = answers[0]
        all_num = all_num.strip()
        all_num = all_num.replace(" ", "")
        all_num = all_num.replace(",", ".")

        self.debug(f"Answer after clean: {all_num}")

        forparsed = ""

        if(all_num.find("...") != -1):
            self.debug(f"Numeric question type is \"...\"")
            if(all_num.count("...") != 1):
                self.__lastError = f"Syntax error in answers \"{all_num}\" of question {row.getRowNum()}. "
                raise SyntaxError
            sep_i = all_num.find("...")
            first = all_num[:sep_i]
            second = all_num[sep_i+len("..."):]
            if(self.isRepresentsFloat(first) == False):
                self.__lastError = f"In answers of question {row.getRowNum()}: \"{first}\" is not number. "
                raise SyntaxError
            if(self.isRepresentsFloat(second) == False):
                self.__lastError = f"In answers of question {row.getRowNum()}: \"{second}\" is not number. "
                raise SyntaxError
            first_num, second_num = float(first), float(second)
            if(first_num >= second_num):
                self.__lastError = f"In answers of question {row.getRowNum()}: {second} must be more than {first} "
                raise SyntaxError
            forparsed = "#" + first + ".." + second
        elif(all_num.find("..") != -1):
            self.debug(f"Numeric question type is \"..\"")
            if(all_num.count("..") != 1):
                self.__lastError = f"Syntax error in answers \"{all_num}\" of question {row.getRowNum()}. "
                raise SyntaxError
            sep_i = all_num.find("..")
            first = all_num[:sep_i]
            second = all_num[sep_i+len(".."):]
            if(self.isRepresentsFloat(first) == False):
                self.__lastError = f"In answers of question {row.getRowNum()}: \"{first}\" is not number. "
                raise SyntaxError
            if(self.isRepresentsFloat(second) == False):
                self.__lastError = f"In answers of question {row.getRowNum()}: \"{second}\" is not number. "
                raise SyntaxError
            first_num, second_num = float(first), float(second)
            if(first_num >= second_num):
                self.__lastError = f"In answers of question {row.getRowNum()}: {second} must be more than {first} "
                raise SyntaxError
            forparsed = "#" + first + ".." + second
        elif(all_num.find("%") != -1):
            self.debug(f"Numeric question type is \"%\"")
            if(all_num.count("%") != 1):
                self.__lastError = f"Syntax error in answers \"{all_num}\" of question {row.getRowNum()}. "
                raise SyntaxError
            sep_i = all_num.find("%")
            first = all_num[:sep_i]
            second = all_num[sep_i+len("%"):]
            if(self.isRepresentsFloat(first) == False):
                self.__lastError = f"In answers of question {row.getRowNum()}: \"{first}\" is not number. "
                raise SyntaxError
            if(self.isRepresentsFloat(second) == False):
                self.__lastError = f"In answers of question {row.getRowNum()}: \"{second}\" is not number. "
                raise SyntaxError
            first_num, second_num = float(first), float(second)
            if(first_num <= second_num):
                self.__lastError = f"In answers of question {row.getRowNum()}: {second} must be less than {first} "
                raise SyntaxError
            forparsed = "#" + first + ":" + second
        else:
            self.debug(f"Numeric question type is standart")
            if(self.isRepresentsFloat(all_num) == False):
                self.__lastError = f"In answers of question {row.getRowNum()}: \"{all_num}\" is not number. "
                raise SyntaxError
            forparsed = "#" + all_num
        
        res = Comment + "\n"
        res += f"::Вопрос {row.getRowNum()}::{Q}" + "{\n"
        res += forparsed + "\n"
        res += "}"

        self.debug(f"answers formed: \n{res}")

        return res

    def question_superOpenPick(self, row):
        cell_1 = row.getCell(1)
        Q, Comment = self.getMarkdownStyleQuestion(cell_1)

        self.debug(f"Question {row.getRowNum()} formed: {Q}")

        res = Comment + "\n"
        res += f"::Вопрос {row.getRowNum()}::{Q}" + "{}"

        self.debug(f"answers formed: \n{res}")

        return res

    def debug(self, text : str):
        if(self.__DEBUG_ON):
            self.__debug += "\n] "
            self.__debug += text

            self.writeTextAppend(self.__debug_file, f"\n] {text}")