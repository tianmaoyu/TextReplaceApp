using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Spire.Xls;
using System.IO;

namespace TextReplaceApp
{
    /// <summary>
    ///
    /// </summary>
    public class ExcelHelper2
    {
        Workbook workbook;
        public ExcelHelper2()
        {
           workbook = new Workbook();
          
        }

        /// <summary>
        /// 再excel中进行查找替换
        /// </summary>
        /// <param name="oldText"></param>
        /// <param name="newText"></param>
        /// <param name="fliePath"></param>
        /// <returns></returns>
        public string ReplaceInExcel(string oldText, string newText, string fliePath)
        {
            workbook.LoadFromFile(fliePath);
            var worksheets = workbook.Worksheets;
            int totall = 0;
            foreach(Worksheet worksheet in worksheets)
            {
                CellRange[] ranges = worksheet.FindAllString(oldText, false, false);
                totall += ranges.Length;
                foreach (CellRange range in ranges)
                {
                    range.Text = newText;
                }
            }
            workbook.SaveToFile(fliePath);
            string fileName = Path.GetFileName(fliePath);
            String result = string.Format("在文件：{0}中-----替换了{1}个\"{2}\"", fileName, totall, oldText);
            return result;
        }

        public string FindInExcel(string oldText, string fliePath)
        {
            workbook.LoadFromFile(fliePath);
            var worksheets = workbook.Worksheets;
            int total = 0;
            foreach (Worksheet worksheet in worksheets)
            {
                CellRange[] ranges = worksheet.FindAllString(oldText, false, false);
                total += ranges.Length;
            }
            string fileName = Path.GetFileName(fliePath);
            string reslut = string.Format("在文件:{0} -----找到 {1} 个:\"{2}\"", fileName, total, oldText);
            return reslut;
        }
    }
}
