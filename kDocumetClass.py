import pythoncom
from win32com.client import Dispatch, gencache
import LDefin2D
import MiscellaneousHelpers as MH
import sys
from pathlib import Path

class kAPPLICATION:

      def __init__(self):
            # Подключим константы API Компас
            self.kConstants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
            self.kConstants3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants
            #  Подключим описание интерфейсов API5
            self.kAPI5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
            self.kObject = self.kAPI5.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(self.kAPI5.KompasObject.CLSID, pythoncom.IID_IDispatch))
            MH.iKompasObject = self.kObject
            # Подключим описание интерфейсов API7
            self.kAPI7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
            self.APP = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0).IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(self.kAPI7.IApplication.CLSID, pythoncom.IID_IDispatch))
            MH.iApplication = self.APP
            self.APP.Visible = True
            # Список всех открытых чертежей
            self.docList = []

      def open(self, path):
            self.docList.append(kDOCUMENT(self.APP.Documents.Open(path), self.kAPI7, self.kAPI5, self.kObject))
      def getActiveDocument(self):
            self.docList.append(kDOCUMENT(self.APP.ActiveDocument, self.kAPI7, self.kAPI5, self.kObject))


class kDOCUMENT:

      def __init__(self, iKompasDocument, ikAPI7, ikAPI5, ikObject):
            self.kDocument = iKompasDocument
            self.kAPI7 = ikAPI7
            self.kAPI5 = ikAPI5
            self.kObject = ikObject
            self.kDocument2D = self.kAPI7.IKompasDocument2D(self.kDocument)
            #  Получим активный документ
            # iDocument2D = self.kObject.ActiveDocument2D()
            self.ViewsCount = self.kAPI7.IKompasDocument2D(self.kDocument).ViewsAndLayersManager.Views.Count
            self.iDrawingDocument = self.kAPI7.IDrawingDocument(self.kDocument._oleobj_.QueryInterface(self.kAPI7.IDrawingDocument.CLSID, pythoncom.IID_IDispatch))
            self.SheetsCount = self.kDocument.LayoutSheets.Count
            self.TextsInView = 0
            self.TablesInView = 0
            self.RoughsCount = 0
            self.LineDimensionsCount = 0
            self.contents()
      def stamp(self):
            iStamp = iDocument2D.GetStamp()
            iStamp.ksOpenStamp()
            iStamp.ksColumnNumber(2)
            iTextLineParam = kompas6_api5_module.ksTextLineParam(
            kompas_object.GetParamStruct(kompas6_constants.ko_TextLineParam))
            iTextLineParam.Init()
            iTextLineParam.style = 32768
            iTextItemArray = kompas_object.GetDynamicArray(4)
            iTextItemParam = kompas6_api5_module.ksTextItemParam(
            kompas_object.GetParamStruct(kompas6_constants.ko_TextItemParam))
            iTextItemParam.Init()
            iTextItemParam.iSNumb = 0
            iTextItemParam.s = "new"
            iTextItemParam.type = 0
            iTextItemFont = kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
            iTextItemFont.Init()
            iTextItemFont.bitVector = 4096
            iTextItemFont.color = 0
            iTextItemFont.fontName = "GOST type A"
            iTextItemFont.height = 7
            iTextItemFont.ksu = 1
            iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
            iTextLineParam.SetTextItemArr(iTextItemArray)

            iStamp.ksTextLine(iTextLineParam)
            iStamp.ksCloseStamp()
            print(type(self.kAPI7))
            print(type(application))
            print(type(Documents))
            print("iDocument2D", type(iDocument2D))
            print("kDocument", type(self.kDocument))
            # iDocument2D.ksLineSeg(53.278238290749, 221.464629752129, 122.397130580339, 137.181898847987, 1)
            # iDocument2D.orthoMode
            # print(type(iDocument2D))
            # print(type(kDocument))
            # print(kDocument.LayoutSheets)
            # print(kDocument.LayoutSheets)
            # kDocument.LayoutSheets.Item(1).Format.Format вовращает формат листа
            # print(kompas_object.Coutn)

            # kDocument.LayoutSheets.Item(1).Format.VerticalOrientation = True
            # kDocument.LayoutSheets.Item(1).Update()

            # application.kDocument.LayoutSheets.Item(1)
            # print(dir(kompas6_api5_module))

            # doc2D = application.ActiveDocument._oleobj_.QueryInterface(self.kAPI7.NamesToIIDMap['IDrawingDocument'], pythoncom.IID_IDispatch)

            # doc2D = self.kAPI7.IDrawingDocument(doc2D)
            # fuck = self.kAPI7.IKompasDocument2D(kDocument).ViewsAndLayersManager.Views.View(1)._oleobj_.QueryInterface(self.kAPI7.NamesToIIDMap['IDrawingContainer'], pythoncom.IID_IDispatch)

            # print(dir(application.ActiveDocument._oleobj_.QueryInterface(self.kAPI7.NamesToIIDMap['IDrawingDocument'], pythoncom.IID_IDispatch)))
            # print(dir(self.kAPI7))
            # print(self.kAPI7.IKompasDocument2D(kDocument).ViewsAndLayersManager.Views.View(1)._oleobj_.QueryInterface(self.kAPI7.NamesToIIDMap['IDrawingContainer'], pythoncom.IID_IDispatch))
            # print(dir(self.kAPI7.IDrawingContainer(self.kAPI7.IKompasDocument2D(kDocument).ViewsAndLayersManager.Views.View(1)._oleobj_.QueryInterface(self.kAPI7.NamesToIIDMap['IDrawingContainer'], pythoncom.IID_IDispatch)).DrawingTexts.Count ))
      def contents(self):
            i = 0
            while (i < self.ViewsCount):
                  iSymbols2DContainer = self.kAPI7.IKompasDocument2D(self.kDocument).ViewsAndLayersManager.Views.View(i)._oleobj_.QueryInterface(self.kAPI7.NamesToIIDMap['ISymbols2DContainer'], pythoncom.IID_IDispatch)
                  iSymbols2DContainer = self.kAPI7.ISymbols2DContainer(iSymbols2DContainer)
                  self.TextsInView += self.kAPI7.IDrawingContainer(self.kAPI7.IKompasDocument2D(self.kDocument).ViewsAndLayersManager.Views.View(i)._oleobj_.QueryInterface(self.kAPI7.NamesToIIDMap['IDrawingContainer'], pythoncom.IID_IDispatch)).DrawingTexts.Count
                  self.TablesInView += iSymbols2DContainer.DrawingTables.Count
                  self.RoughsCount += iSymbols2DContainer.Roughs.Count
                  self.LineDimensionsCount += iSymbols2DContainer.LineDimensions.Count
                  i += 1
      def getStamp(self, n):
            return self.kDocument.LayoutSheets.Item(0).Stamp.Text(n).Str
            #self.kDocument.LayoutSheets.Item(0).Update()
            #self.kDocument.LayoutSheets.Item(0).Stamp.Update()
            #print(self.kDocument.LayoutSheets.Item(0).Stamp.Text(2).Str)
            #print(dir(self.kDocument.LayoutSheets.Item(0).Stamp.Text(2).Str))
      def getFormatList(self):
            i = 0
            ListInfo = ""
            ListInfo += " " + str(self.SheetsCount) + " лист(ов):"
            while i < self.SheetsCount:
                  ListInfo += " A" + str(self.kDocument.LayoutSheets.Item(i).Format.Format)
                  if self.kDocument.LayoutSheets.Item(i).Format.FormatMultiplicity > 1:
                        ListInfo += "x" + str(self.kDocument.LayoutSheets.Item(i).Format.FormatMultiplicity)
                  i += 1
            return ListInfo
      def showDrawContent(self):

            print("Документ", self.kDocument.Name, "cодержит: \n",
                  self.SheetsCount, "лист(а)\n",
                  self.ViewsCount, "вид(a/ов)\n",
                  self.TextsInView, "текст(ов)\n",
                  self.TablesInView, "таблиц(ы)\n",
                  self.RoughsCount, "шероховатостей\n",
                  self.LineDimensionsCount, "линейных размеров\n")
      def autoSpecRough(self):
            # проверка правильности символа дополнительной шероховатости
            if (self.RoughsCount > 0) and (self.iDrawingDocument.SpecRough.AddSign != True):
                  self.iDrawingDocument.SpecRough.AddSign = True
            if (self.RoughsCount == 0) and (self.iDrawingDocument.SpecRough.AddSign == True):
                  self.iDrawingDocument.SpecRough.AddSign = False
            self.iDrawingDocument.SpecRough.Update()
      def textReplace(self, __old, __new):
            # string in text
            i = 0
            while (i < self.ViewsCount):
                  iDrawingContainer = self.kAPI7.IKompasDocument2D(self.kDocument).ViewsAndLayersManager.Views.View(i)._oleobj_.QueryInterface(self.kAPI7.NamesToIIDMap['IDrawingContainer'],pythoncom.IID_IDispatch)
                  iDrawingContainer = self.kAPI7.IDrawingContainer(iDrawingContainer)
                  iDrawingText = iDrawingContainer.DrawingTexts
                  ViewTextCount = iDrawingText.Count
                  y = 0
                  while (y < ViewTextCount):
                        DrawingText_i = iDrawingText.DrawingText(y)
                        TextLinesCount = self.kAPI7.IText(DrawingText_i).Count
                        j = 0
                        while (j < TextLinesCount):
                              w_str = self.kAPI7.IText(DrawingText_i).TextLine(j).Str
                              if __old in w_str:
                                    w_str = w_str.replace(__old, __new)
                                    self.kAPI7.IText(DrawingText_i).TextLine(j).Str = w_str
                                    DrawingText_i.Update()
                              j += 1
                        y += 1
                  i += 1
      def ttReplace(self, __old, __new):
            ttStrCount = self.iDrawingDocument.TechnicalDemand.Text.Count
            i = 0
            flag = 0
            while i < ttStrCount:
                ttStr = self.iDrawingDocument.TechnicalDemand.Text.TextLine(i).Str
                if __old in ttStr:
                    print("технические требования:")
                    print(f"[строка {i + 1}/{ttStrCount}]:", self.iDrawingDocument.TechnicalDemand.Text.TextLine(i).Str)
                    ttStr = ttStr.replace(__old, __new)
                    self.iDrawingDocument.TechnicalDemand.Text.TextLine(i).Str = ttStr
                    flag = 1
                i += 1
            if (flag != True):
                  print("технические требования: совпадений не найдено")
            self.iDrawingDocument.TechnicalDemand.Update()
            # IfaceТech = iDrawingDocument.TechnicalDemand.Text.Add().Add().Str = "ну и залупа этот ваш компас"
      def techDemAutoPos(self):
            LTechDam = 176.5
            HTechDam = 218
            y = 60
            l1 = 2
            l2 = 3.5
            _y = y + l1 + l2
            x1 = 27
            # = [411, 62 + 3.5, 587.5, 100]
            #[x, y, x + LTechDam, HTechDam]
            A4V = [27, _y, 27 + LTechDam, _y + HTechDam]
            #A4G =
            A3V = [114, _y, 114 + LTechDam, _y + HTechDam]
            A3G = [237, _y, 237 + LTechDam, _y + HTechDam]
            self.iDrawingDocument.TechnicalDemand.BlocksGabarits = A3G
            #print(self.iDrawingDocument.TechnicalDemand.BlocksGabarits)
            self.iDrawingDocument.TechnicalDemand.Update()
                  #BlocksGabarits
      def getTechDem(self):
            if self.iDrawingDocument.TechnicalDemand.IsCreated == False:
                  print("Технические требования: отсутсвуют")
                  return
            print("Технические требования: ")
            print("\tлистов: ", int(len(self.iDrawingDocument.TechnicalDemand.BlocksGabarits)/4))
            print("\tстрок: ", self.iDrawingDocument.TechnicalDemand.Text.Count)


      def tableReplace(self, __old, __new):
            i = 0
            while (i < self.ViewsCount):
                  iSymbols2DContainer = self.kAPI7.IKompasDocument2D(self.kDocument).ViewsAndLayersManager.Views.View(i)._oleobj_.QueryInterface(self.kAPI7.NamesToIIDMap['ISymbols2DContainer'], pythoncom.IID_IDispatch)
                  iSymbols2DContainer = self.kAPI7.ISymbols2DContainer(iSymbols2DContainer)
                  TableViewCount = iSymbols2DContainer.DrawingTables.Count
                  if TableViewCount == 0:
                        i += 1
                        continue
                  iDrawingTables = iSymbols2DContainer.DrawingTables
                  j = 0
                  while (j < TableViewCount):
                        iDrawingTable = iDrawingTables.DrawingTable(j)
                        iTable = iDrawingTable._oleobj_.QueryInterface(self.kAPI7.ITable.CLSID, pythoncom.IID_IDispatch)
                        iTable = self.kAPI7.ITable(iTable)
                        ColumnsCount = iTable.ColumnsCount
                        iColumnsCount = 0
                        RowsCount = iTable.RowsCount
                        iRowsCount = 0
                        while (iRowsCount < RowsCount):
                              while (iColumnsCount < ColumnsCount):
                                    iTableCell = iTable.Cell(iColumnsCount, iRowsCount)
                                    iText = self.kAPI7.IText(iTableCell.Text._oleobj_.QueryInterface(self.kAPI7.IText.CLSID, pythoncom.IID_IDispatch))
                                    if __old in iText.TextLine(0).Str:
                                          iText.TextLine(0).Str = iText.TextLine(0).Str.replace(__old, __new)
                                    iDrawingTable.Update()
                                    iColumnsCount += 1
                              iRowsCount += 1
                              iColumnsCount = 0
                        j += 1
                  i += 1


kAPPLICATION = kAPPLICATION()


class DirectionTree(object):
    """Создать дерево каталогов
         @ путь: целевой каталог
         @ filename: имя файла для сохранения
    """
    workPath = r"H:\YandexDisk\Травматология и ортопедия\Блокируемые LCP\Лит О1 (2)\Пластины прямые"
    wFile = r"C:\Users\User\Desktop\Новый текстовый документ.txt"

    def __init__(self, pathname=workPath, filename=wFile):
        super(DirectionTree, self).__init__()
        self.pathname = Path(pathname)
        self.filename = filename
        self.tree = ""

    def set_path(self, pathname):
        self.pathname = Path(pathname)

    def set_filename(self, filename):
        self.filename = filename

    def generate_tree(self, n=0):
        if self.pathname.is_file():
            if "cdw" in self.pathname.name:
                self.tree += '    ' * n + ' ' * 4 + self.pathname.name
                #print(self.pathname)
                kAPPLICATION.open(self.pathname)
                self.tree += kAPPLICATION.docList[0].getFormatList()
                self.tree += "\tРазрaб.: " + kAPPLICATION.docList[0].getStamp(110) + ','
                self.tree += " Пров.: " + kAPPLICATION.docList[0].getStamp(111) + '\n'
                kAPPLICATION.docList.pop()
            else:
                pass
        elif self.pathname.is_dir():
            self.tree += '    ' * n + ' ' * 4 + \
                str(self.pathname.relative_to(self.pathname.parent)) + ':' + '\n'
            for cp in self.pathname.iterdir():
                self.pathname = Path(cp)
                self.generate_tree(n + 1)

    def save_file(self):
        with open(self.filename, 'w', encoding='utf-8') as f:
            f.write(self.tree)
def max_str_len(tree):
      max_str_len = 0
      for str in tree:
           if len(str) > max_str_len:
                max_str_len += len(str)
      return max_str_len


dirtree = DirectionTree()
dirtree.generate_tree()
dirtree.save_file()
print(dirtree.tree)
