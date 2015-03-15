/* File: ExcelWriter.cs
 * Author: Daniel Teodoro Gonçalves Mariano
 * Email: dtgmariano@gmail.com
 * Date: 2015/03/11
 * 
 * Comments: CSharp class to create Excel spreadsheets
 */
namespace Biolab
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Data;
    using Excel = Microsoft.Office.Interop.Excel;
    using System.Reflection;
    using System.Diagnostics;

    public class ExcelWriter
    {
        string filename { get; set; }
        Excel.Application xlApp { get; set; }
        Excel.Workbook wb { get; set; }

        public ExcelWriter(string _filename)
        {
            filename = _filename;
            xlApp = new Excel.Application(filename);
            wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
        }

        public void DatatableToExcel(DataTable _table, string _sheetName)
        {
            wb.Sheets.Add();
            Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[1];
            ws.Name = _sheetName;

            for (int i = 0; i < _table.Columns.Count; i++)
            {
                string rowIndex = "1";
                string columnIndex = NumberToLetter(i + 1);
                string header = _table.Columns[i].Caption;

                Excel.Range aRange = ws.get_Range(columnIndex + rowIndex, columnIndex + rowIndex);
                object[] args = new object[1];
                args[0] = header;
                aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);
            }

            for (int j = 0; j < _table.Rows.Count; j++)
            {
                {
                    for (int i = 0; i < _table.Columns.Count; i++)
                    {
                        string rowIndex = (j + 2).ToString();
                        string columnIndex = NumberToLetter(i + 1);
                        Excel.Range aRange = ws.get_Range(columnIndex + rowIndex);
                        object[] args = new object[1];
                        args[0] = _table.Rows[j][i];
                        aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);
                    }
                }
            }

        }

        public void saveExcel()
        {
            wb.SaveAs(filename);
            wb.Close();
        }

        private string NumberToLetter(int _num)
        {
            #region switch cases
            switch (_num)
            {
                case 1:
                    return "A";

                case 2:
                    return "B";

                case 3:
                    return "C";

                case 4:
                    return "D";

                case 5:
                    return "E";


                case 6:
                    return "F";

                case 7:
                    return "G";

                case 8:
                    return "H";

                case 9:
                    return "I";

                case 10:
                    return "J";


                case 11:
                    return "K";

                case 12:
                    return "L";

                case 13:
                    return "M";

                case 14:
                    return "N";

                case 15:
                    return "O";



                case 16:
                    return "P";

                case 17:
                    return "Q";

                case 18:
                    return "R";

                case 19:
                    return "S";

                case 20:
                    return "T";



                case 21:
                    return "U";

                case 22:
                    return "V";

                case 23:
                    return "W";

                case 24:
                    return "X";

                case 25:
                    return "Y";



                case 26:
                    return "Z";
            }
            #endregion

            return "A";
        }
    }
}
