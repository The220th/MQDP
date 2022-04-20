# -*- coding: utf-8 -*-

import zipfile
import re
from bs4 import BeautifulSoup
from typing import Final

#class Image:
#    pass

class Doc:
    __fileName = ""
    __z = None # zip
    __lines = []
    __nlines = 0

    def __init__(self, fileName : str):
        self.__fileName = fileName
        self.__z = zipfile.ZipFile(fileName)
        self.__initLines()
        #print(self.__z.namelist())
    
    def getDocXML(self) -> str:
        return self.__z.open("word/document.xml").read().decode("utf-8")

    def __initLines(self):
        s = self.getDocXML()
        reres = re.findall(r"<w:p.*?<\/w:p>|<w:tbl>.*?<\/w:tbl>", s)
        for xmltag in reres:
            #print("\n\n" + xmltag)
            oneline = Line(xmltag, self)
            self.__lines.append(oneline)

    def __iter__(self):
        self.__nlines = 0
        return self

    def __next__(self):
        if self.__nlines < len(self.__lines):
            result = self.__lines[self.__nlines]
            self.__nlines -=-1
            return result
        else:
            raise StopIteration

    def readFileText(self, path : str) -> str:
        return self.__z.open(path).read().decode("utf-8")

    def getImageBytes(self, image : 'Image') -> bytes:
        return image.getBytes()

    def readImage(self, pathImage : str) -> bytes:
        return self.__z.open(pathImage).read()
    
    def getLines(self):
        return self.__lines



class Line:
    #_TEXT : Final = "text"
    #_IMAGE : Final = "img"
    #_TABLE : Final = "table"
    
    @staticmethod
    def TEXT() -> str: return "=TEXT"

    @staticmethod
    def IMAGE() -> str: return "=IMG"

    @staticmethod
    def TABLE() -> str: return "=TABLE"

    @staticmethod
    def OTHER() -> str: return "=OTHER"

    __type = None

    __src = None

    __doc = None

    def __init__(self, xml : str, doc):

        self.__doc = doc

        if(xml.find("<w:tbl") != -1):     # Таблица
            self.__ifTable(xml)
            #print("\n\n" + f"{xml}" + "\n\n")
        elif(xml.find("<pic:pic") != -1): # Картинка
            self.__ifPic(xml)
        elif(xml.find("<w:t") != -1):     # Текст
            self.__ifText(xml)
        else:                             # Что это?!
            self.__ifOther(xml)
            #raise Exception("Cannot determine the type... ")

    def __ifTable(self, xmltag : str) -> None:
        # pizza time
        #print(xmltag)
        self.__type = Line.TABLE()
        self.__src = Table(xmltag, self.__doc)
        return

    def __ifPic(self, xmltag : str) -> None:
        self.__type = Line.IMAGE()
        self.__src = Image(xmltag, self.__doc)
        return

    def __ifText(self, xmltag : str) -> None:
        self.__type = Line.TEXT()
        self.__src = Text(xmltag)
        return

    def __ifOther(self, xmltag : str) -> None:
        self.__type = Line.OTHER()
        self.__src = xmltag

    def getType(self) -> str:
        return self.__type

    def isText(self) -> bool:
        if(self.__type == Line.TEXT()):
            return True
        else:
            return False

    def isImage(self) -> bool:
        if(self.__type == Line.IMAGE()):
            return True
        else:
            return False

    def isTable(self) -> bool:
        if(self.__type == Line.TABLE()):
            return True
        else:
            return False

    def isOther(self) -> bool:
        if(self.__type == Line.OTHER()):
            return True
        else:
            return False

    def getSrc(self):
        return self.__src

class Text:
    __text = ""
    __bold = None
    __italic = None
    __underline = None
    __color = None

    __soup = None

    __runs = None
    __nruns = None

    def __init__(self, xmlstr : str):

        xmlstr = str(xmlstr)


        ps = re.findall(r"<w:p.*?<\/w:p>", xmlstr)
        for p in ps:
            pbuff = p
            if(pbuff.find("w:tab") != -1):
                if(pbuff.find("<w:t ") != -1):
                    pbuff = pbuff[pbuff.find("<w:t "):]
                elif(pbuff.find("<w:t>") != -1):
                    pbuff = pbuff[pbuff.find("<w:t>"):]
            ws = re.findall(r"<w:t.*?</w:t>", pbuff)
            for w in ws:
                w = w[w.find(">")+1:len(w)-len("</w:t>")]
                self.__text += w
        #print("Text inited")

        self.__soup = BeautifulSoup(xmlstr, "lxml")
        run_blocks = self.__soup.find_all("w:r")

        self.__runs = []
        for run in run_blocks:
            self.__runs.append(Run(run))
        
        self.__checkProperties()

    def __checkProperties(self):
        self.__bold = False
        self.__italic = False
        self.__underline = False
        self.__color = None

        for run in self.__runs:
            if(run.getText() != None and run.getText() != ""):
                if(run.isBold() == True):
                    self.__bold = True
                if(run.isItalic() == True):
                    self.__italic = True
                if(run.isUnderline() == True):
                    self.__underline = True
                if(self.__color == None and run.getColor() != None):
                    if(run.getColor() != "auto"):
                        self.__color = run.getColor()
            #print(run.getColor())

    def isBold(self) -> bool:
        return self.__bold
    def isItalic(self) -> bool:
        return self.__italic
    def isUnderline(self) -> bool:
        return self.__underline
    def isColored(self) -> bool:
        if(self.__color != None):
            return True
        else:
            return False
    def getColor(self) -> str: # https://colorscheme.ru/color-converter.html
        return self.__color

    def getText(self) -> str:
        return self.__text

    def __iter__(self):
        self.__nruns = 0
        return self

    def __next__(self):
        if self.__nruns < len(self.__runs):
            result = self.__runs[self.__nruns]
            self.__nruns+=1
            return result
        else:
            raise StopIteration

    def getRuns(self):
        return self.__runs

class Run:

    __text = None

    __soup = None

    __bold = False # <w:b></w:b>
    # bold false is <w:b w:val="0"></w:b>
    __italic = False # <w:i></w:i>
    # italic false is <w:i w:val="0"></w:i>
    __underline = False # Например: <w:u w:val="single"></w:u>
    # underline false is <w:u w:val="none"></w:u>
    __color = None # Либо вообще нет, либо, например, 
                                #<w:color w:themecolor="accent6" w:val="70AD47"></w:color>

    # Цвет позади: <w:highlight w:val="none"></w:highlight>

    
    def __init__(self, xmlstr : str):
        xmlstr = str(xmlstr)
        #print(xmlstr)
        if(xmlstr.find("</w:b>") != -1):
            self.__bold = True
            minisoup = BeautifulSoup(xmlstr, "lxml")
            minires = minisoup.find("w:b")
            if("w:val" in minires.attrs and minires["w:val"] == "0"):
                self.__bold = False
        if(xmlstr.find("</w:i>") != -1):
            self.__italic = True
            minisoup = BeautifulSoup(xmlstr, "lxml")
            minires = minisoup.find("w:i")
            if("w:val" in minires.attrs and minires["w:val"] == "0"):
                self.__italic = False
        if(xmlstr.find("</w:u>") != -1):
            self.__underline = True
            minisoup = BeautifulSoup(xmlstr, "lxml")
            minires = minisoup.find("w:u")
            if("w:val" in minires.attrs and minires["w:val"] == "none"):
                self.__underline = False
        if(xmlstr.find("</w:color>") != -1):
            buff = xmlstr
            buff = buff[buff.find("<w:color"):] # <w:color w:themecolor="accent6" w:val="70AD47"...
            buff = buff[buff.find("w:val=\""):] # w:val="70AD47"...
            buff = buff[buff.find("\"")+1:] # 70AD47"...
            buff = buff[:buff.find("\"")] # 70AD47
            self.__color = buff

        #print(xmlstr)
        self.__soup = BeautifulSoup(xmlstr, "lxml")
        wt = self.__soup.find("w:t")
        #print(wt)
        if(wt != None):
            self.__text = wt.text
        
    def isBold(self) -> bool:
        return self.__bold
    def isItalic(self) -> bool:
        return self.__italic
    def isUnderline(self) -> bool:
        return self.__underline
    def isColored(self) -> bool:
        if(self.__color != None):
            return True
        else:
            return False
    def getColor(self) -> str: # https://colorscheme.ru/color-converter.html
        return self.__color

    def getText(self):
        return self.__text



class Image:

    __imageID = None # rIdX, где X = число

    __imagePath = None

    __doc = None

    def __init__(self, xmlstr : str, doc : Doc):

        self.__doc = doc

        #print(xmlstr)
        buff = xmlstr
        buff = buff[buff.find("<a:blip"):] # <a:blip...r:embed="rId10"...>
        buff = buff[buff.find("r:embed"):] # r:embed="rId10"...>
        buff = buff[buff.find("\"")+1:] # rId10
        buff = buff[:buff.find("\"")]
        self.__imageID = buff
        #self.__imageID = re.findall(r'<a:blip.+?embed=\s*".+?(?=\s*/>)', xmlstr)[0]
        #self.__imageID = re.findall(r'(?<=").+?(?=")', self.__imageID)[0] 
        #print(xmlstr)
        #print(str(123) + " " + self.__imageID) # rId9
        self.__calcImagePath(self.__doc.readFileText("word/_rels/document.xml.rels"))
    
    def __calcImagePath(self, document_xml_rels : str):
        #print("Id\\s*=\\s*\"" + str(self.__imageID) + "\".+?(?=\"/>)")
        #print(document_xml_rels)
        path = re.findall("Id\\s*=\\s*\"" + str(self.__imageID) + "\".+?(?=\"/>)", document_xml_rels)[0]
        path = re.findall(r'(?<=Target=").+', path)[0]
        path = "word/" + path
        #print(path)
        self.__imagePath = path

    def getBytes(self) -> bytes:
        return self.__doc.readImage(self.__imagePath)




class Table:

    __doc = None

    __rows = None
    __nrows = 0

    __soup = None


    def __init__(self, xmlstr : str, doc : 'Doc'):
        self.__doc = doc

        #print(xmlstr)
        xmlstr = str(xmlstr)
        self.__soup = BeautifulSoup(xmlstr, "lxml")
        blocks = self.__soup.find_all("w:tr")
        self.__rows = []
        li = 0
        for el in blocks:
            self.__rows.append(Row(el, li, self, self.__doc))
            li+=1
        #print(f"rows: {len(blocks)}")

    def getCell(self, i : int, j : int) -> 'Cell':
        buffRow = self.__rows[i]
        buffCell = buffRow.getCell(j)
        return buffCell

    def _setCell(self, i : int, j : int, cell : 'Cell'):
        self.__rows[i]._setCell(j, cell)
    
    def __iter__(self):
        self.__nrows = 0
        return self

    def __next__(self):
        if self.__nrows < len(self.__rows):
            result = self.__rows[self.__nrows]
            self.__nrows += 1
            return result
        else:
            raise StopIteration



class Row:

    __table = None
    __doc = None

    __cells = None
    __ncells = 0

    __soup = None

    __ix_i = 0

    def __init__(self, xmlstr : str, i : int, table : 'Table', doc : 'Doc'):
        self.__ix_i = i
        self.__table = table
        self.__doc = doc

        xmlstr = str(xmlstr)
        #print(xmlstr)
        self.__soup = BeautifulSoup(xmlstr, "lxml")
        blocks = self.__soup.find_all("w:tc")
        
        ncells = len(blocks)
        for cell_i in blocks:
            buffcell = str(cell_i)
            if(buffcell.find("</w:gridspan>") != -1):
                buff = buffcell
                #print(buff)
                buff = buff[buff.find("w:val"):] # w:val="X"...
                buff = buff[buff.find("\"")+1:] # X"...
                buff = buff[:buff.find("\"")] # X
                ncells += (int(buff)-1)
        #print(ncells)
        self.__cells = [None]*ncells

        el_i = 0
        for lj in range(ncells):
            #print(self.__cells[lj])
            if(self.__cells[lj] == None):
                #print(f"{lj}:{self.__cells[lj]}")
                el = blocks[el_i]
                self.__cells[lj] = Cell(el, self.__ix_i, lj, self, self.__table, self.__doc)
                el_i+=1

        #print("Row inited")

    def getRowNum(self):
        return self.__ix_i

    def getCell(self, i : int) -> 'Cell':
        buffCell = self.__cells[i]
        return buffCell

    def _setCell(self, j : int, cell : 'Cell'):
        self.__cells[j] = cell

    def __iter__(self):
        self.__ncells = 0
        return self

    def __next__(self):
        if self.__ncells < len(self.__cells):
            result = self.__cells[self.__ncells]
            self.__ncells += 1
            return result
        else:
            raise StopIteration





class Cell:

    __row = None
    __table = None
    __doc = None

    __lines = None
    __nlines = 0

    __ix_i = 0
    __ix_j = 0

    __isvMerged = None
    __ishMerged = None
    #__mergedCell = None

    def __init__(self, xmlstr : str, i : int, j : int, row : 'Row', table : 'Table', doc : 'Doc', ignor = False):
        if(ignor == False):
            self.__ix_i = i
            self.__ix_j = j
            self.__row = row
            self.__table = table
            self.__doc = doc

            xmlstr = str(xmlstr) #<w:tc>...</w:tc>
            #print(xmlstr)

            self.__isvMerged = False
            self.__ishMerged = False

            #if(xmlstr.find("<w:vMerge/>") != -1): # BeautifulSoup lxml
                                                    # Переводит всё в нижний регистр
                                                    # И "дозакрывает" теги, 6л&%!3


            # Иногда:
            #    <w:vmerge w:val="restart"></w:vmerge> - в инициирующей ячейке
            #    <w:vmerge w:val="continue"></w:vmerge> - в продолжающей ячейке
            # А иногда, 6l%^&:
            #    <w:vmerge w:val="restart"></w:vmerge> - в инициирующей ячейке
            #    <w:vmerge></w:vmerge> - в продолжающей ячейке
            if(xmlstr.find("<w:vmerge>") != -1 or xmlstr.find("<w:vmerge w:val=\"continue\">") != -1):
                #print(f"{i}:{j}:{xmlstr}")
                self.__isvMerged = True
                #print(f"{i}, {j}")
                self.__lines = table.getCell(i-1, j).getLines()
                #self.__mergedCell = table.getCell(i-1, j)
                #table._setCell(i, j, table.getCell(i-1, j))

            #   <w:gridSpan w:val="X"/> - значит занимает X ячеек по горизонтали
            # Но BeautifulSoup lxml, поэтому:
            #    <w:gridspan w:val="X"></w:gridspan>
            elif(xmlstr.find("</w:gridspan>") != -1):
                #print(f"{i}:{j}:{xmlstr}")

                buff = xmlstr
                buff = buff[buff.find("w:val"):] # w:val="X"...
                buff = buff[buff.find("\"")+1:] # X"...
                buff = buff[:buff.find("\"")] # X
                howm = int(buff)

                self.__lines = []
                reres = re.findall(r"<w:p.*?<\/w:p>|<w:tbl>.*?<\/w:tbl>", xmlstr)
                for xmltag in reres:
                    self.__lines.append(Line(xmltag, self.__doc))

                for i in range(howm-1):
                    nextCell = Cell(None, None, None, None, None, None, True)
                    nextCell.__ishMerged = True
                    nextCell.__isvMerged = False
                    nextCell.__lines = self.__lines
                    nextCell.__ix_i = self.__ix_i
                    nextCell.__ix_j = self.__ix_j + i + 1
                    nextCell.__table = self.__table
                    nextCell.__doc = self.__doc
                    table.getCell(i, j-1).getLines()
                    #print(nextCell.__ix_j)
                    row._setCell(nextCell.__ix_j, nextCell)
                    
            else:
                self.__lines = []
                reres = re.findall(r"<w:p.*?<\/w:p>|<w:tbl>.*?<\/w:tbl>", xmlstr)
                for xmltag in reres:
                    self.__lines.append(Line(xmltag, self.__doc))
                #print("Cell inited")

    def getPosition(self):
        return (self.__ix_i, self.__ix_j)

    def getLines(self) -> list:
        return self.__lines

    def isMerged(self):
        if(self.__isvMerged == True or self.__ishMerged == True):
            return True
        else:
            return False

    def is_vMerged(self):
        return self.__isvMerged

    def is_hMerged(self):
        return self.__ishMerged

    def __iter__(self):
        self.__nlines = 0
        return self

    def __next__(self):
        if self.__nlines < len(self.__lines):
            result = self.__lines[self.__nlines]
            self.__nlines -=-1
            return result
        else:
            raise StopIteration
