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

class standardk_process:

    __docPath = None
    __outPath = None
    __imgFolder = None
    __relImgFolder = "imgs"

    __debug = ""

    __lastError = None

    __image_i = None

    def __init__(self, docPath : str, outPath: str):
        self.__docPath = docPath
        self.__outPath = outPath
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
        if(r < 50 and g > 100 and b < 50):
            return True
        elif(r > 100 and g < 50 and b < 50):
            return False
        else:
            return None

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


    def question_depo(self, row : Row) -> str:
        self.debug(f"trying to detect type of row {row.getRowNum()}")
        firstCell = row.getCell(0)
        for line in firstCell:
            if(line.isText()):
                text = line.getSrc().getText()
                #print(text)
                if(text[0] == 'О' or text[0] == 'М' or text[0] == 'К' or text[0] == 'Ф' or
                text[0] == 'С' or text[0] == 'Ч' or text[0] == 'Э'):
                    qmark = text[0]
                    break
            else:
                self.__lastError = f"Syntax error. Cannot define type of question in row {row.getRowNum()}"
                raise SyntaxError
        self.debug(f"type of row {row.getRowNum()} is {qmark}")
        
        res = "Somthing else" # !!!
        if(qmark == 'О'):
            res = self.question_OnePick(row)
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
        self.debug(f"Question formed: {Q}")

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
                    ans.append("")

                # O6PA6OTKA begin
                if(line.isText()):
                    text = line.getSrc()
                    #print(f"{123}: {text.getText()} {text.isBold()} {text.getColor()}")
                    if(text.getText().strip()[0] == "="
                    or text.isBold()
                    or text.isUnderline() 
                    or ( text.isColored() and self.checkColorRight(text.getColor()) == True )
                    ): # right ans
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


        res = f"::Вопрос {row.getRowNum()}::{Q}" + "{\n"
        for an in ans:
            res += an + "\n"
        res + "}"

        self.debug(f"answers formed: \n{res}")

        return res


        


    def debug(self, text : str):
        self.__debug += "\n] "
        self.__debug += text