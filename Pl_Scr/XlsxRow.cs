using System;
using System.Collections;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pl_Scr
{
    public class XlsxRow : IEnumerable
    {
        public Row Row { get; set; }

        public XlsxRow(Row row)
        {
            Row = row;
        }

        public IEnumerator GetEnumerator()
        {
            int index = 0;
            foreach (Cell cell in Row.Descendants<Cell>())
            {
                int columnIndex = ConvertColumnNameToIndex(GetColumnName(cell.CellReference));
                for (; index < columnIndex; index++)
                {
                    yield return new Cell();
                }
                yield return cell;
                index++;
            }
        }

        public static string GetColumnName(string cellReference)
        {
            return Regex.Match(cellReference, "[A-Za-z]+").Value;
        }

        public static int ConvertColumnNameToIndex(string columnName)
        {
            char[] colLetters = columnName.ToCharArray();
            Array.Reverse(colLetters);
            return colLetters
                .Select((letter, i) => i == 0 ? letter - 65 : letter - 64)
                .Select((current, i) => current*(int) Math.Pow(26, i))
                .Sum();
        }
    }
}
