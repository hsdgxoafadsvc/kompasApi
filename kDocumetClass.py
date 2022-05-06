import pythoncom
from win32com.client import Dispatch, gencache
import LDefin2D
import MiscellaneousHelpers as MH

#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
MH.iKompasObject = kompas_object

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0).IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))
MH.iApplication = application

Documents = application.Documents

#  Получим активный документ
iDocument2D = kompas_object.ActiveDocument2D()
application.Visible = True

class kDocument():

      def __init__(self):
            self.kompas_document = application.ActiveDocument
            self.kompas_document_2d = kompas_api7_module.IKompasDocument2D(self.kompas_document)
            self.iDrawingDocument = kompas_api7_module.IDrawingDocument(self.kompas_document._oleobj_.QueryInterface(kompas_api7_module.IDrawingDocument.CLSID,pythoncom.IID_IDispatch))
            self.SheetsCount = self.kompas_document.LayoutSheets.Count
            self.ViewsCount = kompas_api7_module.IKompasDocument2D(self.kompas_document).ViewsAndLayersManager.Views.Count
            self.TextsInView = 0
            self.TablesInView = 0
            self.RoughsCoutn = 0
            self.LineDimensionsCount = 0
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
            print(type(kompas_api7_module))
            print(type(application))
            print(type(Documents))
            print("iDocument2D", type(iDocument2D))
            print("kompas_document", type(self.kompas_document))
            # iDocument2D.ksLineSeg(53.278238290749, 221.464629752129, 122.397130580339, 137.181898847987, 1)
            # iDocument2D.orthoMode
            # print(type(iDocument2D))
            # print(type(kompas_document))
            # print(kompas_document.LayoutSheets)
            # print(kompas_document.LayoutSheets)
            # kompas_document.LayoutSheets.Item(1).Format.Format вовращает формат листа
            # print(kompas_object.Coutn)

            # kompas_document.LayoutSheets.Item(1).Format.VerticalOrientation = True
            # kompas_document.LayoutSheets.Item(1).Update()

            # application.kompas_document.LayoutSheets.Item(1)
            # print(dir(kompas6_api5_module))

            # doc2D = application.ActiveDocument._oleobj_.QueryInterface(kompas_api7_module.NamesToIIDMap['IDrawingDocument'], pythoncom.IID_IDispatch)

            # doc2D = kompas_api7_module.IDrawingDocument(doc2D)
            # fuck = kompas_api7_module.IKompasDocument2D(kompas_document).ViewsAndLayersManager.Views.View(1)._oleobj_.QueryInterface(kompas_api7_module.NamesToIIDMap['IDrawingContainer'], pythoncom.IID_IDispatch)

            # print(dir(application.ActiveDocument._oleobj_.QueryInterface(kompas_api7_module.NamesToIIDMap['IDrawingDocument'], pythoncom.IID_IDispatch)))
            # print(dir(kompas_api7_module))
            # print(kompas_api7_module.IKompasDocument2D(kompas_document).ViewsAndLayersManager.Views.View(1)._oleobj_.QueryInterface(kompas_api7_module.NamesToIIDMap['IDrawingContainer'], pythoncom.IID_IDispatch))
            # print(dir(kompas_api7_module.IDrawingContainer(kompas_api7_module.IKompasDocument2D(kompas_document).ViewsAndLayersManager.Views.View(1)._oleobj_.QueryInterface(kompas_api7_module.NamesToIIDMap['IDrawingContainer'], pythoncom.IID_IDispatch)).DrawingTexts.Count ))
      def contents(self):
            i = 0
            while (i < self.ViewsCount):
                  iSymbols2DContainer = kompas_api7_module.IKompasDocument2D(self.kompas_document).ViewsAndLayersManager.Views.View(i)._oleobj_.QueryInterface(kompas_api7_module.NamesToIIDMap['ISymbols2DContainer'], pythoncom.IID_IDispatch)
                  iSymbols2DContainer = kompas_api7_module.ISymbols2DContainer(iSymbols2DContainer)
                  self.TextsInView += kompas_api7_module.IDrawingContainer(kompas_api7_module.IKompasDocument2D(self.kompas_document).ViewsAndLayersManager.Views.View(i)._oleobj_.QueryInterface(kompas_api7_module.NamesToIIDMap['IDrawingContainer'], pythoncom.IID_IDispatch)).DrawingTexts.Count
                  self.TablesInView += iSymbols2DContainer.DrawingTables.Count
                  self.RoughsCoutn += iSymbols2DContainer.Roughs.Count
                  self.LineDimensionsCount += iSymbols2DContainer.LineDimensions.Count
                  i += 1
      def showDrawContent(self):
            print("Документ", self.kompas_document.Name, "cодержит: \n",
                  self.SheetsCount, "лист(а)\n",
                  self.ViewsCount, "вид(a/ов)\n",
                  self.TextsInView, "текст(ов)\n",
                  self.TablesInView, "таблиц(ы)\n",
                  self.RoughsCoutn, "шероховатостей\n",
                  self.LineDimensionsCount, "линейных размеров\n")
      def autoSpecRough(self):
            # проверка правильности символа дополнительной шероховатости
            if ((self.RoughsCoutn > 0) and (self.iDrawingDocument.SpecRough.AddSign != True)):
                  self.iDrawingDocument.SpecRough.AddSign = True
            if ((self.RoughsCoutn == 0) and (self.iDrawingDocument.SpecRough.AddSign == True)):
                  self.iDrawingDocument.SpecRough.AddSign = False
            self.iDrawingDocument.SpecRough.Update()
      def textInViews(self, old, new):
            # string in text
            i = 0
            while (i < self.ViewsCount):
                  iDrawingContainer = kompas_api7_module.IKompasDocument2D(self.kompas_document).ViewsAndLayersManager.Views.View(i)._oleobj_.QueryInterface(kompas_api7_module.NamesToIIDMap['IDrawingContainer'],pythoncom.IID_IDispatch)
                  iDrawingContainer = kompas_api7_module.IDrawingContainer(iDrawingContainer)
                  iDrawingText = iDrawingContainer.DrawingTexts
                  ViewTextCount = iDrawingText.Count
                  y = 0
                  while (y < ViewTextCount):
                        DrawingText_i = iDrawingText.DrawingText(y)
                        TextLinesCount = kompas_api7_module.IText(DrawingText_i).Count
                        j = 0
                        while (j < TextLinesCount):
                              w_str = kompas_api7_module.IText(DrawingText_i).TextLine(j).Str
                              if old in w_str:
                                    w_str = w_str.replace(old, new)
                                    kompas_api7_module.IText(DrawingText_i).TextLine(j).Str = w_str
                                    DrawingText_i.Update()
                              j += 1
                        y += 1
                  i += 1
      def tt(self, old, new):
            ttStrCount = self.iDrawingDocument.TechnicalDemand.Text.Count
            i = 0
            flag = 0
            while i < ttStrCount:
                ttStr = self.iDrawingDocument.TechnicalDemand.Text.TextLine(i).Str
                if old in ttStr:
                    print("технические требования:")
                    print(f"[строка {i + 1}/{ttStrCount}]:", self.iDrawingDocument.TechnicalDemand.Text.TextLine(i).Str)
                    ttStr = ttStr.replace(old, new)
                    self.iDrawingDocument.TechnicalDemand.Text.TextLine(i).Str = ttStr
                    flag = 1
                i += 1
            if (flag != True):
                  print("технические требования: совпадений не найдено")
            self.iDrawingDocument.TechnicalDemand.Update()
            # IfaceТech = iDrawingDocument.TechnicalDemand.Text.Add().Add().Str = "ну и залупа этот ваш компас"
      def table(self, old, new):
            i = 0
            while (i < self.ViewsCount):
                  iSymbols2DContainer = kompas_api7_module.IKompasDocument2D(self.kompas_document).ViewsAndLayersManager.Views.View(i)._oleobj_.QueryInterface(kompas_api7_module.NamesToIIDMap['ISymbols2DContainer'], pythoncom.IID_IDispatch)
                  iSymbols2DContainer = kompas_api7_module.ISymbols2DContainer(iSymbols2DContainer)
                  TableViewCount = iSymbols2DContainer.DrawingTables.Count
                  if TableViewCount == 0:
                        i += 1
                        continue
                  iDrawingTables = iSymbols2DContainer.DrawingTables
                  j = 0
                  while (j < TableViewCount):
                        iDrawingTable = iDrawingTables.DrawingTable(j)
                        iTable = iDrawingTable._oleobj_.QueryInterface(kompas_api7_module.ITable.CLSID, pythoncom.IID_IDispatch)
                        iTable = kompas_api7_module.ITable(iTable)
                        ColumnsCount = iTable.ColumnsCount
                        iColumnsCount = 0
                        RowsCount = iTable.RowsCount
                        iRowsCount = 0
                        while (iRowsCount < RowsCount):
                              while (iColumnsCount < ColumnsCount):
                                    iTableCell = iTable.Cell(iColumnsCount, iRowsCount)
                                    iText = kompas_api7_module.IText(iTableCell.Text._oleobj_.QueryInterface(kompas_api7_module.IText.CLSID, pythoncom.IID_IDispatch))
                                    if old in iText.TextLine(0).Str:
                                          iText.TextLine(0).Str = iText.TextLine(0).Str.replace(old, new)
                                    iDrawingTable.Update()
                                    iColumnsCount += 1
                              iRowsCount += 1
                              iColumnsCount = 0
                        j += 1
                  i += 1

document = kDocument()
document.table("ИШПЖ", "МЕШБ")