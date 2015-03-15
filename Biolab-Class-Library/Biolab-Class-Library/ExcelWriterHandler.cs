/* File: ExcelWriterHandler.cs
 * Author: Daniel Teodoro Gonçalves Mariano
 * Email: dtgmariano@gmail.com
 * Date: 2015/03/11
 * 
 * Comments: CSharp static class to handle ExcelWriter class
 */
namespace Biolab
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Data;
    using System.Reflection;
    using System.Diagnostics;

    public static class ExcelWriterHandler
    {
        public static void createExcel_OneSheet(string filename, string sheetname, List<object> columns, List<object[]> rows)
        {
            ExcelWriter writer = new ExcelWriter(filename + ".xlsx");
            DataTable sheet1 = createDataTable(columns, rows);
            writer.DatatableToExcel(sheet1, sheetname);
            writer.saveExcel();
        }

        public static void createExcel_MultipleSheets(string filename, List<string>sheetNames, List<DataTable> tables)
        {
            ExcelWriter writer = new ExcelWriter(filename + ".xlsx");
            int nsheets = 0;
            if (sheetNames.Count == tables.Count)
                nsheets = sheetNames.Count;
            for (int i = 0; i < nsheets; i++)
                writer.DatatableToExcel(tables[i], sheetNames[i]);
            writer.saveExcel();
        }

        static DataTable createDataTable(List<object> columns, List<object[]> rows)
        {
            DataTable dt = new DataTable();
            columns.ForEach(x => dt.Columns.Add(x.ToString()));
            rows.ForEach(x => dt.Rows.Add(x));
            return dt;
        }
    }
}
