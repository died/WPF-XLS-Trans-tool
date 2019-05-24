using System;
using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xls_Trans_Tool_Wpf
{
    public class ExcelTool
    {
        public static InnerResult GetExcelFile(string file)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            var xlApp = new Excel.Application();
            var workbooks = xlApp.Workbooks;
            Excel.Workbook xlWorkbook;
            try
            {
                xlWorkbook = workbooks.Open(file, null, ReadOnly: true);
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception Message: " + e.Message);
                if (e.InnerException != null) MessageBox.Show("InnerException Message: " + e.InnerException.Message);
                MessageBox.Show("Exception Trace : " + e.StackTrace);
                return new InnerResult { Success = false, Message = "Error on open file."};
            }

            var sheets = xlWorkbook.Sheets;
            Console.WriteLine(sheets.Count);
            Excel._Worksheet xlWorksheet = null;
            if (sheets.Count == 1) xlWorksheet = xlWorkbook.Sheets[1];
            else
            {
                for (int i = 1; i <= sheets.Count; i++)
                {
                    Excel._Worksheet tmp = sheets[i];
                    if (tmp.Name != "PO Upload Download Template-APP") continue;
                    xlWorksheet = tmp;
                    break;
                }
            }
            //check if sheet null
            if (xlWorksheet == null && sheets.Count>0)
            {
                //try first one
                xlWorksheet = xlWorkbook.Sheets[1];
                
                //forgot quit
                //xlWorkbook.Close(false);
                //xlApp.Quit();
                //return new InnerResult { Success = false, Message = Wording.CantFindRightColumns };
            }

            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            Debug.WriteLine($"row={rowCount} col={colCount}");

            #region get skip
            var target = xlWorksheet.UsedRange.Find("Issue Date:");
            if (target == null)
            {
                xlWorkbook.Close(false);
                xlApp.Quit();
                return new InnerResult { Success = false, Message = Wording.CantFindRightColumns };
            }
            var skip = target.Row;
            #endregion

            #region check header
            //Check header
            var heads = xlWorksheet.Range[target, xlWorksheet.Cells[skip, Models.AdidasHeader.Length]];
            var check = CheckHeader(heads, Models.AdidasHeader);
            if (!check.Success)
            {
                xlWorkbook.Close(false);
                xlApp.Quit();
                return new InnerResult { Success = false, Message = check.Message };
            }
            #endregion

            #region toObject
            object[,] data = xlRange.Value2;
            DataTable dt = new DataTable();
            // Create new Column in DataTable
            for (int cCnt = 1; cCnt <= colCount; cCnt++)
            {
                var column = new DataColumn
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = cCnt.ToString()
                };
                dt.Columns.Add(column);

                // Create row for Data Table
                for (int rCnt = skip + 1; rCnt <= rowCount; rCnt++)
                {
                    string cellVal;
                    if (Models.DateTimeColume.Contains(cCnt))
                    {
                        try
                        {
                            cellVal = data[rCnt, cCnt] != null
                                ? DateTime.FromOADate(Convert.ToDouble(data[rCnt, cCnt])).ToShortDateString()
                                : null;
                        }
                        catch (Exception)
                        {
                            cellVal = data[rCnt, cCnt]?.ToString();
                        }
                    }
                    else
                    {
                        cellVal = data[rCnt, cCnt]?.ToString(); 
                    }

                    DataRow row;

                    // Add to the DataTable
                    if (cCnt == 1)
                    {
                        row = dt.NewRow();
                        row[cCnt.ToString()] = cellVal;
                        dt.Rows.Add(row);
                    }
                    else
                    {
                        row = dt.Rows[rCnt - 1 - skip];
                        row[cCnt.ToString()] = cellVal;
                    }
                }
            }
            #endregion

            //close and release
            
            xlWorkbook.Close(false);
            workbooks.Close();
            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine(@"Load Success");
            return new InnerResult { Success = true, Data = dt };
        }

        /// <summary>
        /// Check Head title match
        /// </summary>
        /// <param name="row"></param>
        /// <param name="heads"></param>
        /// <returns></returns>
        public static InnerResult CheckHeader(Excel.Range row, string[] heads)
        {

            //check count
            if (row.Columns.Count != heads.Length)
            {
                var msg = string.Format(Wording.WrongHeaderNumbers, row.Columns.Count, heads.Length); 
                Debug.WriteLine(msg);
                return new InnerResult { Success = false, Message = msg };
            }

            //check content
            for (var i = 0; i < row.Columns.Count; i++)
            {
                var s1 = (string)row.Cells[i + 1].Value2;
                var s2 = heads[i];

                var editLimit = s2.Length > 10 ? 2 : 1;
                
                //try use edit distance
                if(CalcLevenshteinDistance(Regex.Replace(s1, @"\s+|\r\n", string.Empty),Regex.Replace(s2, @"\s+|\r\n", string.Empty))>editLimit)
                //remove line break and space before compare
                //if (Regex.Replace(s1, @"\s+|\r\n", string.Empty) != Regex.Replace(s2, @"\s+|\r\n", string.Empty))
                {
                    var msg = string.Format(Wording.WrongHeader, s1, s2); 
                    Debug.WriteLine(msg);
                    //TODO log
                    return new InnerResult { Success = false, Message = msg };
                }
            }
            return new InnerResult { Success = true };
        }

        /// <summary>
        /// Columns num to ABC for excel
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string IntToLetters(int value)
        {
            string result = string.Empty;
            while (--value >= 0)
            {
                result = (char)('A' + value % 26) + result;
                value /= 26;
            }
            return result;
        }

        /// <summary>
        /// Columns ABC to num for excel
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static int LettersToInt(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) return 0;

            columnName = columnName.ToUpperInvariant();
            var sum = 0;
            foreach (char t in columnName)
            {
                sum *= 26;
                sum += (t - 'A' + 1);
            }

            return sum;
        }

        /// <summary>
        /// Return min distance
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        private static int CalcLevenshteinDistance(string a, string b)
        {
            if (String.IsNullOrEmpty(a) || String.IsNullOrEmpty(b)) return 0;

            int lengthA = a.Length;
            int lengthB = b.Length;
            var distances = new int[lengthA + 1, lengthB + 1];

            for (int i = 1; i <= lengthA; i++)
            for (int j = 1; j <= lengthB; j++)
            {
                int cost = b[j - 1] == a[i - 1] ? 0 : 1;
                distances[i, j] = Math.Min
                (
                    Math.Min(distances[i - 1, j] + 1, distances[i, j - 1] + 1),
                    distances[i - 1, j - 1] + cost
                );
            }
            return distances[lengthA, lengthB];
        }
    }
}
