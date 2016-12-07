using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.IO;
using Office = Microsoft.Office.Core;
using VBIDE = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace TextReplaceApp
{

    /// <summary>
    /// 对excel 进行 加入宏，然后运行宏操作[可以运行]，
    /// </summary>
    public class ExcelHelper
    {
        private Microsoft.Office.Interop.Excel.Application excelApp;
        private string newPath;
        string dateString = "未知";

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public ExcelHelper()
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
        }
        //
        public string RepaceForExcel(string text, string newText, string excelPath)
        {
            //打开Eecel文件
            Workbooks workbooks = excelApp.Workbooks;
            Workbook workbook = workbooks.Add(excelPath);
            //取得sheets
            int strCount = 0;
            Sheets sheets = workbook.Sheets;
            foreach (Worksheet sheet in sheets)
            {
                //最大行，最大列2003为准
                Range sourceRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1000, 50]];
                for (int i = 1; i < 1000; i++)
                {
                    for (int j = 1; j < 50; j++)
                    {
                        var range = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[i, j];
                        var _text = range.Text.ToString();
                        if (!string.IsNullOrEmpty(_text))
                        {
                            if (_text.Contains(text))
                            {
                                string newStr = range.Value2.ToString().Replace(text, newText);
                                range.Value2 = newStr;
                                strCount++;
                            }
                        }

                    }
                }
            }
            workbook.Close(false, Type.Missing, Type.Missing);
            string fileName = Path.GetFileName(excelPath);
            string reslut = string.Format("在文件{0}-----替换了{1}个{2}", fileName, strCount, text);
            return reslut;
        }

        public string FindForExcel(string text, string excelPath)
        {
            //打开Eecel文件
            Workbooks workbooks = excelApp.Workbooks;
            Workbook workbook = workbooks.Add(excelPath);
            //取得sheets
            int strCount = 0;
            Sheets sheets = workbook.Sheets;
            foreach (Worksheet sheet in sheets)
            {
                //最大行，最大列2003为准
                Range sourceRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[65536, 256]];
                for (int i = 1; i < 65536; i++)
                {
                    for (int j = 1; j < 256; j++)
                    {
                        var range = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[i, j];
                        var _text = range.Text.ToString();
                        if (!string.IsNullOrEmpty(_text))
                        {
                            if (_text.Contains(text))
                            {
                                strCount++;
                            }
                        }
                    }
                }
            }
            workbook.Close(false, Type.Missing, Type.Missing);
            string fileName = Path.GetFileName(excelPath);
            return string.Format("在文件：{0}中---- - 找到{1}个\"{2}\"", fileName, strCount, text);
        }


        public void AddVBAForExcel(string text, string newText, string excelPath)
        {

            VBIDE.VBComponent oModule;
            Office.CommandBar oCommandBar;
            Office.CommandBarButton oCommandBarButton;
            String sCode;
            Object oMissing = System.Reflection.Missing.Value;

            Workbooks workbooks = excelApp.Workbooks;
            Workbook workbook = excelApp.Workbooks.Open(excelPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing
              , Type.Missing, Type.Missing, Type.Missing, Type.Missing
              , Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Create a new VBA code module.
            oModule = workbook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);

            sCode =
                "sub VBAMacro()\r\n" +
                " Dim Cz As String\r\n" +
                " Dim Th As String\r\n" +
                " Cz = \"" + text + "\"\r\n" +
                " Th = \"" + newText + "\"\r\n" +
                "  Cells.Replace What:=Cz, Replacement:=Th, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False\r\n" +
                "end sub";
            // Add the VBA macro to the new code module.
            oModule.CodeModule.AddFromString(sCode);

            try
            {
                // Create a new toolbar and show it to the user.
                oCommandBar = excelApp.CommandBars.Add("VBAMacroCommandBar", oMissing, oMissing);
                oCommandBar.Visible = false;
                // Create a new button on the toolbar.
                oCommandBarButton = (Office.CommandBarButton)oCommandBar.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    oMissing, oMissing, oMissing, oMissing);
                // Assign a macro to the button.
                oCommandBarButton.OnAction = "VBAMacro";
                // Set the caption of the button.
                oCommandBarButton.Caption = "Call VBAMacro";
                // Set the icon on the button to a picture.
                oCommandBarButton.FaceId = 2151;
            }
            catch (Exception eCBError)
            {
                //已经存在了
            }

            excelApp.UserControl = true;

            int strCount = 0;
            Sheets sheets = workbook.Sheets;
            ExcelMacroHelper ex = new ExcelMacroHelper(excelApp, workbook);
            object outStr = null;
            object[] objs = new object[0];
            foreach (Worksheet sheet in sheets)
            {
                string sheetName = sheet.Name;
                sheet.Activate();
                string vbaName = "VBAMacro";
                ex.RunExcelMacro(excelPath, vbaName, objs, outStr, false);
            }
            workbook.Save();
            workbook.Close(false, Type.Missing, Type.Missing);

            // Release the variables.
            oCommandBarButton = null;
            oCommandBar = null;
            oModule = null;
            workbook = null;
            excelApp = null;
            // Collect garbage.
            GC.Collect();
        }

    }

    public class ExcelMacroHelper
    {

        Excel.Application oExcel;
        private Excel._Workbook oBook;
        public ExcelMacroHelper(Excel.Application excel, Excel._Workbook oBook)
        {
            this.oExcel = excel;
            this.oBook = oBook;
        }
        /// <summary>  
        /// 执行Excel中的宏  
        /// </summary>  
        /// <param name="excelFilePath">Excel文件路径</param>  
        /// <param name="macroName">宏名称</param>  
        /// <param name="parameters">宏参数组</param>  
        /// <param name="rtnValue">宏返回值</param>  
        /// <param name="isShowExcel">执行时是否显示Excel</param> 
        //public void RunExcelMacro(string excelFilePath, string macroName, object[] parameters, out object rtnValue, bool isShowExcel)
        public void RunExcelMacro(string excelFilePath, string macroName, object[] parameters, object rtnValue, bool isShowExcel)
        {
            try
            {
                #region 检查入参
                //检查文件是否存在  
                if (!File.Exists(excelFilePath))
                {
                    MessageBox.Show(excelFilePath + " 文件不存在");
                    //return;
                }
                // 检查是否输入宏名称  
                if (string.IsNullOrEmpty(macroName))
                {
                    MessageBox.Show("请输入宏的名称");
                    //return;
                }
                #endregion
                #region 调用宏处理
                // 准备打开Excel文件时的缺省参数对象  
                object oMissing = System.Reflection.Missing.Value;
                // 根据参数组是否为空，准备参数组对象  
                object[] paraObjects;
                if (parameters == null)
                {
                    paraObjects = new object[] { macroName };
                }
                else
                {
                    // 宏参数组长度  
                    int paraLength = parameters.Length;
                    paraObjects = new object[paraLength + 1];
                    paraObjects[0] = macroName;
                    for (int i = 0; i < paraLength; i++)
                    {
                        paraObjects[i + 1] = parameters[i];
                    }
                }
                // 创建Excel对象示例  
                //Excel.ApplicationClass oExcel = new Excel.ApplicationClass();
                // 判断是否要求执行时Excel可见  
                if (isShowExcel)
                {
                    // 使创建的对象可见  
                    oExcel.Visible = false;
                }

                rtnValue = this.RunMacro(oExcel, paraObjects);
                // 保存更改  
                oBook.Save();

            }
            catch (Exception)
            {
                throw;
            }
        }
        /// <summary>  
        /// 执行宏  
        /// </summary>  
        /// <param name="oApp">Excel对象</param>  
        /// <param name="oRunArgs">参数（第一个参数为指定宏名称，后面为指定宏的参数值）</param>  
        /// <returns>宏返回值</returns>  
        private object RunMacro(object oApp, object[] oRunArgs)
        {
            try
            {
                // 声明一个返回对象  
                object objRtn;
                // 反射方式执行宏  
                objRtn = oApp.GetType().InvokeMember(
                                                        "Run",
                                                        System.Reflection.BindingFlags.Default |
                                                        System.Reflection.BindingFlags.InvokeMethod,
                                                        null,
                                                        oApp,
                                                        oRunArgs
                                                     );
                // 返回值  
                return objRtn;
            }
            catch (Exception ex)
            {
                // 如果有底层异常，抛出底层异常  
                if (ex.InnerException.Message.ToString().Length > 0)
                {
                    throw ex.InnerException;
                }
                else
                {
                    throw ex;
                }
            }
        }

        internal void RunExcelMacro()
        {
            throw new Exception("The method or operation is not implemented.");
        }
    }
}
#endregion
