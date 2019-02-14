using System;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using LexTalionis.StreamTools;

namespace LexTalionis.ExcelTools
{
    /// <summary>
    /// Класс для работы с Excel
    /// </summary>
    public static class Excel
    {
        /// <summary>
        /// Вернуть таблицу данных с первого листа файла
        /// </summary>
        /// <param name="file">имя файла</param>
        /// <returns>таблица</returns>
        public static DataTable GetDt(string file)
        {
            DataTable dataTable;
            using (
                var spreadSheetDocument =
                    SpreadsheetDocument.Open(file, false))
            {
                dataTable = GetDtCore(spreadSheetDocument);
            }
            return dataTable;
        }

        /// <summary>
        /// Вернуть таблицу данных с первого листа файла
        /// </summary>
        /// <param name="file">содержимое файла</param>
        /// <returns>таблица</returns>
        public static DataTable GetDt(Stream file)
        {
            DataTable dataTable;
            var tmp = Path.GetTempFileName();

            using (var tmpfile = new FileStream(tmp, FileMode.Create))
            {
                file.CopyTo(tmpfile);
                tmpfile.Seek(0, SeekOrigin.Begin);

                using (
                var spreadSheetDocument =
                    SpreadsheetDocument.Open(tmpfile, false))
                {
                    dataTable = GetDtCore(spreadSheetDocument);
                }
            }

            File.Delete(tmp);
            return dataTable;
        }

        /// <summary>
        /// Ядро процесса получения таблицы данных
        /// </summary>
        /// <param name="spreadSheetDocument">документ</param>
        /// <returns>таблица</returns>
        private static DataTable GetDtCore(SpreadsheetDocument spreadSheetDocument)
        {
            var dataTable = new DataTable();

                var sheets =
                    spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                var relationshipId = sheets.First().Id.Value;
                var worksheetPart =
                    (WorksheetPart) spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                var workSheet = worksheetPart.Worksheet;
                var sheetData = workSheet.GetFirstChild<SheetData>();
                var rows = sheetData.Descendants<Row>().ToArray();

                foreach (Cell cell in rows.ElementAt(0))
                {
                    dataTable.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }
                foreach (var row in rows)
                {
                    var dataRow = dataTable.NewRow();
                    for (var i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        dataRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                    }

                    dataTable.Rows.Add(dataRow);
                }
            dataTable.Rows.RemoveAt(0);
            return dataTable;
        }

        /// <summary>
        /// Получить данные из ячейки
        /// </summary>
        /// <param name="document">документ</param>
        /// <param name="cell">ячейка</param>
        /// <returns>данные</returns>
        private static string GetCellValue(SpreadsheetDocument document, CellType cell)
        {
            string result = null;
            var stringTablePart = document.WorkbookPart.SharedStringTablePart;
            var cellvalue = cell.CellValue;
            if (cellvalue != null)
            {
                var value = cellvalue.InnerXml;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                }
                result = value;
            }
            return result;
        }

        /// <summary>
        /// Копировать вкладку
        /// </summary>
        /// <param name="input">откуда</param>
        /// <param name="output">куда</param>
        /// <param name="pageFrom">имя вкладки в исходнике</param>
        /// <param name="pageTo">новое имя вкладки</param>
        public static void CopySheet(Stream input, Stream output, string pageFrom, string pageTo)
        {
            using (var inputReport = SpreadsheetDocument.Open(input, false))
            {
                var workbookPart = inputReport.WorkbookPart;
                var sourceSheetPart = GetWorkSheetPart(workbookPart, pageFrom);
                //var sharedStringTable = workbookPart.SharedStringTablePart;
                //var wbsp = workbookPart.WorkbookStylesPart;
                using (var outputReport = SpreadsheetDocument.Open(output, true))
                {
                    var newWorkbookPart = outputReport.WorkbookPart;
                    var newWorksheetPart = newWorkbookPart.AddPart(sourceSheetPart);

                    //Table definition parts are somewhat special and need unique ids...so let's make an id based on count
                    var numTableDefParts = sourceSheetPart.GetPartsCountOfType<TableDefinitionPart>();

                    //Clean up table definition parts (tables need unique ids)
                    if (numTableDefParts != 0)
                        FixupTableParts(newWorksheetPart, numTableDefParts);
                    CleanView(newWorksheetPart);
                    var wb = newWorkbookPart.Workbook;
                    var index = UInt32.Parse(wb.Sheets.Last().GetAttributes().First(x => x.LocalName == "sheetId").Value);
                    var sheets = wb.GetFirstChild<Sheets>();
                    var copiedSheet = new Sheet
                        {
                            Name = pageTo,
                            Id = newWorkbookPart.GetIdOfPart(newWorksheetPart),
                            SheetId = ++index
                        };
                    sheets.Append(new OpenXmlElement[]{copiedSheet});
                    var tmp = sheets.Descendants<Sheet>().FirstOrDefault(x => x.Name == pageFrom);
                    if (tmp != null)
                        tmp.Remove();
                    newWorksheetPart.Worksheet.Save();
                    wb.Save();
                }
            }
        }
        /// <summary>
        /// Получить отношение
        /// </summary>
        /// <param name="workbookPart">отношение в книге</param>
        /// <param name="sheetName">имя вкладки</param>
        /// <returns>отношение в листах</returns>
        private static WorksheetPart GetWorkSheetPart(WorkbookPart workbookPart, string sheetName)
        {
            //Get the relationship id of the sheetname
            string relId = workbookPart.Workbook.Descendants<Sheet>().First(s => s.Name.Value.Equals(sheetName))
                .Id;

            return (WorksheetPart)workbookPart.GetPartById(relId);
        }
        /// <summary>
        /// Очистить представление
        /// </summary>
        /// <param name="worksheetPart">отношения во вкладках</param>
        private static void CleanView(WorksheetPart worksheetPart)
        {
            var views = worksheetPart.Worksheet.GetFirstChild<SheetViews>();

            if (views != null)
            {
                views.Remove();
                worksheetPart.Worksheet.Save();
            }
        }
        private static void FixupTableParts(WorksheetPart worksheetPart, int tableId)
        {
            foreach (var tableDefPart in worksheetPart.TableDefinitionParts)
            {
                tableId++;
                tableDefPart.Table.Id = (uint)tableId;
                tableDefPart.Table.DisplayName = "CopiedTable" + tableId;
                tableDefPart.Table.Name = "CopiedTable" + tableId;
                tableDefPart.Table.Save();
            }
        }
        /// <summary>
        /// Переименовать вкладку
        /// </summary>
        /// <param name="file">отчет</param>
        /// <param name="nameFrom">оригинальное наименование</param>
        /// <param name="nameResult">результирующее наименование</param>
        public static void RenameSheet(Stream file, string nameFrom, string nameResult)
        {
            using (var mySpreadsheet = SpreadsheetDocument.Open(file, true))
            {
                var workbookPart = mySpreadsheet.WorkbookPart;
                var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(x => x.Name == nameFrom);
                if (sheet != null)
                    sheet.Name = nameResult;
            }
        }
    }
}
