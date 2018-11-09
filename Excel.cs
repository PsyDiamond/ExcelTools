using System;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using LexTalionis.StreamTools;

namespace LexTalionis.ExcelTools
{
    /// <summary>
    /// Класс для работы с Excel
    /// </summary>
    public class Excel
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
    }
}
