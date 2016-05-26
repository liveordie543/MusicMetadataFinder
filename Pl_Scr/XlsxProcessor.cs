using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pl_Scr
{
    public class XlsxProcessor
    {
        private Tuple<SheetData, SharedStringTablePart> _sheetInfo;

        private int _currentRowNumber = 1;

        private string FilePath { get; }

        private List<XlsxRow> Rows { get; set; }

        public XlsxProcessor(string filePath)
        {
            FilePath = filePath;
        }

        public void Process()
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(FilePath, true))
            {
                PopulateRows(document);
                InititalizeSheetInfo(document);

                SharedStringTablePart stringTablePart = _sheetInfo.Item2;
                foreach (XlsxRow row in Rows.Skip(1))
                {
                    _currentRowNumber++;
                    Dictionary<int, string> songInfo = GetSongInfo(row, stringTablePart);
                    string releaseDate = LastFmApiHelper.GetReleaseDateBySong(songInfo[0], songInfo[1]);
                    row.Row.Append(SetCellValue(3, (uint)_currentRowNumber, releaseDate));
                }
            }
        }

        private void PopulateRows(SpreadsheetDocument document)
        {
            InititalizeSheetInfo(document);
            SheetData sheetData = _sheetInfo.Item1;
            Rows = new List<XlsxRow>();
            foreach (Row row in sheetData.Elements<Row>())
            {
                Rows.Add(new XlsxRow(row));
            }
        }

        private void InititalizeSheetInfo(SpreadsheetDocument document)
        {
            if (_sheetInfo != null)
            {
                return;
            }
            WorkbookPart workbookPart = document.WorkbookPart;
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            SharedStringTablePart stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            _sheetInfo = new Tuple<SheetData, SharedStringTablePart>(sheetData, stringTable);
        }

        private static Dictionary<int, string> GetSongInfo(XlsxRow row, SharedStringTablePart stringTablePart)
        {
            int i = 0;
            Dictionary<int, string> songInfo = new Dictionary<int, string>();
            foreach (Cell cell in row)
            {
                string cellValue = GetCellValue(cell, stringTablePart);
                if (!String.IsNullOrWhiteSpace(cellValue))
                {
                    songInfo.Add(i, cellValue);
                }
                i++;
            }
            return songInfo;
        }

        private static string GetCellValue(CellType cell, SharedStringTablePart stringTablePart)
        {
            string result = String.Empty;
            if (cell != null)
            {
                result = cell.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && stringTablePart != null)
                {
                    result = stringTablePart.SharedStringTable.ElementAt(int.Parse(result)).InnerText;
                }
            }
            return result;
        }

        public Cell SetCellValue(int columnIndex, uint rowIndex, string text)
        {
            string column = GetColumnName(columnIndex);
            string row = rowIndex.ToString();
            return new Cell(new InlineString(new List<Text> { new Text(text) }))
            {
                DataType = CellValues.InlineString,
                CellReference = column + row
            };
        }

        public string GetColumnName(int columnIndex)
        {
            string columnName = String.Empty;
            while (columnIndex > 0)
            {
                int modulo = (columnIndex - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                columnIndex = (columnIndex - modulo) / 26;
            }
            return columnName;
        }
    }
}
