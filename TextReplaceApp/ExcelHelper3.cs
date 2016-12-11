using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;

namespace TextReplaceApp
{
    class ExcelHelper3
    {
        private Microsoft.Office.Interop.Excel.Application excelApp;
        private string newPath;
        string dateString = "未知";

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public ExcelHelper3()
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            //设置不可见
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
        }
        //
        public string ReplaceInExcel(string text, string newText, string excelPath)
        {
            //打开Eecel文件
            Workbooks workbooks = excelApp.Workbooks;
           
            //Workbook workbook = workbooks.Add(excelPath);//这个方式打开的不能保存成功
            //这个能保存成功
            Workbook workbook = excelApp.Workbooks.Open(excelPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing
             , Type.Missing, Type.Missing, Type.Missing, Type.Missing
             , Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //取得sheets
            int strCount = 0;
            Sheets sheets = workbook.Sheets;
            
            foreach (Worksheet sheet in sheets)
            {
                //    表页下使用区域的行数、列数  
               int iRowCount = sheet.UsedRange.Cells.Rows.Count;
               int  iColCount = sheet.UsedRange.Cells.Columns.Count;
               //    表页下使用区域的起始行列号  
               int  iBeginRow = sheet.UsedRange.Cells.Row;
               int iBeginCol = sheet.UsedRange.Cells.Column;
                for (int Row = iBeginRow; Row < iRowCount + iBeginRow; Row++)
                {
                    for (int Col = iBeginCol; Col < iColCount + iBeginCol; Col++)
                    {
                        var range = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[Row, Col];
                        var _text = range.Text.ToString();
                        if (!string.IsNullOrEmpty(_text))
                        {
                            if (_text.IndexOf(text) >= 0)
                            {
                                Regex regex = new Regex(text);
                                var matches = regex.Matches(_text);
                                strCount += matches.Count;
                                string newStr = range.Value2.ToString().Replace(text, newText);
                                range.Value =  newStr;
                         
                            }
                        }

                    }
                }
            }
            workbook.Save();
            workbook.Close(false, Type.Missing, Type.Missing);
            string fileName = Path.GetFileName(excelPath);
            string reslut = string.Format("在文件{0}-----替换了{1}个{2}", fileName, strCount, text);
            ColseExcel();
            return reslut;
        }

        public string FindInExcel(string text, string excelPath)
        {
            //打开Eecel文件
            Workbooks workbooks = excelApp.Workbooks;
            Workbook workbook = workbooks.Add(excelPath);
            //取得sheets
            int strCount = 0;
            Sheets sheets = workbook.Sheets;
            foreach (Worksheet sheet in sheets)
            {
                //    表页下使用区域的行数、列数  
                int iRowCount = sheet.UsedRange.Cells.Rows.Count;
                int iColCount = sheet.UsedRange.Cells.Columns.Count;
                //    表页下使用区域的起始行列号  
                int iBeginRow = sheet.UsedRange.Cells.Row;
                int iBeginCol = sheet.UsedRange.Cells.Column;
                for (int Row = iBeginRow; Row < iRowCount + iBeginRow; Row++)
                {
                    for (int Col = iBeginCol; Col < iColCount + iBeginCol; Col++)
                    {
                        var range = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[Row, Col];
                        var _text = range.Text.ToString();
                        if (!string.IsNullOrEmpty(_text))
                        {
                            if (_text.IndexOf(_text) >= 0)
                            {
                                Regex regex = new Regex(text);
                                var matches = regex.Matches(_text);
                                strCount  += matches.Count;
                            }
                        }
                    }
                }
            }
            workbook.Close(false, Type.Missing, Type.Missing);
            string fileName = Path.GetFileName(excelPath);
            ColseExcel();
            return string.Format("在文件：{0}中---- - 找到{1}个\"{2}\"", fileName, strCount, text);
        }

        //关闭的新新方法
        private void ColseExcel()
        {
            //关闭
            IntPtr t = new IntPtr(excelApp.Hwnd);
            int k = 0;
            GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();
        }

    }
}
