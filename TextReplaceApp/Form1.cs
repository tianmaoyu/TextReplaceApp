using pptWrite;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace TextReplaceApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
      


        private void cob_path_MouseDown_1(object sender, MouseEventArgs e)
        {
            string path = null;
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowNewFolderButton = false;
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                path = folderBrowserDialog.SelectedPath;
            }
            else
            {
            }
            this.cob_path.Text = path;
        }
        /// <summary>
        /// 查找操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_find_Click(object sender, EventArgs e)
        {
            if (IsOpenedPPT())
            {
                MessageBox.Show("请先关闭PPT!");
                return;
            }
            string directoyPath = this.cob_path.Text;
            var IsWaring = directoyPath.Length == 3 && (directoyPath.ToUpper().Contains("C")
                || directoyPath.ToUpper().Contains("D") || directoyPath.ToUpper().Contains("E"));
            if (IsWaring)
            {
                MessageBox.Show("当前选择的是一个磁盘，太危险！禁止操作，请选择磁盘中的一个文件");
                return;
            }
            string fileText = this.tb_findText.Text;
            FindEnum where = GetWhere();
            if (where.Equals(FindEnum.InContent))
            {
                var filePaths = GetFlies(directoyPath);
                string message = "";
                foreach (string fliePath in filePaths)//\n
                {
                    //HTML.XML
                    if (Path.GetExtension(fliePath).ToUpper().Equals(".HTML") || Path.GetExtension(fliePath).ToUpper().Equals(".XML"))
                    {
                        message += FindInHTMLOrXML(fileText, fliePath) + "\n";
                    }
                    if (Path.GetExtension(fliePath).ToUpper().Equals(".TXT"))
                    {
                        message += FindInTXT(fileText, fliePath) + "\n";

                    }
                    if (Path.GetExtension(fliePath).ToUpper().Equals(".XLSX") || Path.GetExtension(fliePath).ToUpper().Equals(".XLS"))
                    {
                        ExcelHelper2 exceHelper = new ExcelHelper2();
                        message += exceHelper.FindInExcel(fileText, fliePath) + "\n";
                    }
                    if (Path.GetExtension(fliePath).ToUpper().Equals(".DOCX") || Path.GetExtension(fliePath).ToUpper().Equals(".RTF"))
                    {
                        WordHelper wordHelper = new WordHelper();
                        message += wordHelper.FindInWord(fileText, fliePath) + "\n"; ;
                    }
                    if (Path.GetExtension(fliePath).ToUpper().Equals(".PPTX") || Path.GetExtension(fliePath).ToUpper().Equals(".PPT"))
                    {
                        OperatePPT ppt = new OperatePPT();
                        ppt.PPTOpen(fliePath);
                        message += ppt.FindInPPT(fileText, fliePath) + "\n"; ;
                        ppt.PPTClose();
                    }
                }
                MessageBox.Show(message);
            }

        }
        /// <summary>
        /// 替换操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_replace_Click(object sender, EventArgs e)
        {
            if (IsOpenedPPT())
            {
                MessageBox.Show("请先关闭PPT!");
                return;
            }
            string directoyPath = this.cob_path.Text;
            string fileText = this.tb_findText.Text;
            string repalceText = this.tb_replaceText.Text.Trim();
            var IsWaring = directoyPath.Length == 3 && (directoyPath.ToUpper().Contains("C")
                || directoyPath.ToUpper().Contains("D") || directoyPath.ToUpper().Contains("E"));
            if (IsWaring)
            {
                MessageBox.Show("当前选择的是一个磁盘，太危险！禁止操作，请选择磁盘中的一个文件");
                return;
            }
            if (string.IsNullOrEmpty(repalceText))
            {
                MessageBox.Show("当前要替换的文字不能为空！");
                return;
            }

            FindEnum where = GetWhere();
            if (where.Equals(FindEnum.InContent))
            {
                var filePaths = GetFlies(directoyPath);
                string message = "";
                foreach (string fliePath in filePaths)//\n
                {
                    //HTML.XML
                    if (Path.GetExtension(fliePath).ToUpper().Equals(".HTML") || Path.GetExtension(fliePath).ToUpper().Equals(".XML"))
                    {
                        message += RelaceInHTMLOrXML(fileText, repalceText, fliePath) + "\n";
                    }
                    //txt
                    if (Path.GetExtension(fliePath).ToUpper().Equals(".TXT"))
                    {
                        message += ReplaceInTXT(fileText, repalceText, fliePath) + "\n";
                    }
                    //excel
                    if (Path.GetExtension(fliePath).ToUpper().Equals(".XLSX") || Path.GetExtension(fliePath).ToUpper().Equals(".XLS"))
                    {
                        ExcelHelper2 exceHelper = new ExcelHelper2();
                        message += exceHelper.ReplaceInExcel(fileText, repalceText, fliePath) + "\n"; 
                    }
                    //word,rft
                    if (Path.GetExtension(fliePath).ToUpper().Equals(".DOCX") || Path.GetExtension(fliePath).ToUpper().Equals(".RTF"))
                    {
                        WordHelper wordHelper = new WordHelper();
                        message += wordHelper.ReplaceInWord(fileText, repalceText, fliePath) + "\n"; ;
                    }
                    //ppt
                    if (Path.GetExtension(fliePath).ToUpper().Equals(".PPTX") || Path.GetExtension(fliePath).ToUpper().Equals(".PPT"))
                    {
                        OperatePPT ppt = new OperatePPT();
                        ppt.PPTOpen(fliePath);
                        message += ppt.ReplaceAll(fileText, repalceText, fliePath) + "\n"; ;
                        ppt.PPTClose();
                    }
                }
                MessageBox.Show(message);
            }
        }
        /// <summary>
        ///在txt 文本中查找
        /// </summary>
        /// <param name="text"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public string FindInTXT(string text, string filePath)
        {
            if (!File.Exists(filePath))
            {
                return "没找到" + filePath;
            }

            StreamReader sr = new StreamReader(filePath, Encoding.Default);
            string line;
            int total = 0;
            while ((line = sr.ReadLine()) != null)
            {
                Regex regex = new Regex(text);
                var matches = regex.Matches(line);
                total += matches.Count;
            }
            string fileName = Path.GetFileName(filePath);
            sr.Close();
            string reslut = string.Format("在文件:{0} -----找到 {1} 个:\"{2}\"", fileName, total, text);
            return reslut;
        }

        /// <summary>
        /// 的查找内容的个数
        /// </summary>
        /// <param name="text"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public string FindTextCountInContent(string text, string filePath)
        {
            if (!File.Exists(filePath))
            {
                return "没找到" + filePath;
            }
            StreamReader sr = new StreamReader(filePath, Encoding.Default);
            string line;
            int total = 0;
            while ((line = sr.ReadLine()) != null)
            {
               
                Regex regex = new Regex(text);
                var matches = regex.Matches(line);
                total += matches.Count;
            }
            sr.Close();
            return total.ToString();
        }


        public string ReplaceInTXT(string text, string newText, string filePath)
        {
            if (!File.Exists(filePath))
            {
                return "没找到" + filePath + "或者无效文件";
            }

            //得到为替换前的个数
            string countStr = FindTextCountInContent(newText, filePath);
            int oldCount = 0;
            int newCount = 0;
            Int32.TryParse(countStr, out oldCount);


            string con = "";
            StreamReader sr = new StreamReader(filePath, Encoding.Default);
            con = sr.ReadToEnd();
            con = con.Replace(text, newText);
            sr.Close();
            StreamWriter sw = new StreamWriter(filePath, false, Encoding.Default);
            sw.WriteLine(con);
            sw.Close();

            countStr = FindTextCountInContent(newText, filePath);
            Int32.TryParse(countStr, out newCount);

            string fileName = Path.GetFileName(filePath);
            string reslut = string.Format("在文件{0}-----替换了{1}个{2}", fileName, newCount - oldCount, text);
            return reslut;
        }

        public String FindInHTMLOrXML(string text, string filePath)
        {
            int replaceCount = 0;
            string content = "";
            StreamReader sr = new StreamReader(filePath, Encoding.Default);
            content = sr.ReadToEnd();
            Regex regex = new Regex(@"(?<=>).*?(?=<)");
            var matches = regex.Matches(content);
            foreach (Match match in matches)
            {

                if (match.Value.Contains(text))
                {
                    replaceCount++;
                }
            }
            sr.Close();
            string fileName = Path.GetFileName(filePath);
            String result = string.Format("在文件：{0}中-----找到{1}个\"{2}\"", fileName, replaceCount, text);
            return result;
        }
        /// <summary>
        /// 替换在html 中，xml 中的
        /// </summary>
        /// <param name="text"></param>
        /// <param name="newText"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public String RelaceInHTMLOrXML(string text, string newText, string filePath)
        {
            List<string> oldStrings = new List<string>();
            List<string> newStrings = new List<string>();
            int replaceCount = 0;
            string content = "";
            StreamReader sr = new StreamReader(filePath, Encoding.GetEncoding("gb2312"));
            content = sr.ReadToEnd();
            Regex regex = new Regex(@"(?<=>).*?(?=<)");
            var matches = regex.Matches(content);
            foreach (Match match in matches)
            {

                if (match.Value.Contains(text))
                {
                    oldStrings.Add(match.Value);
                    newStrings.Add(match.Value.Replace(text, newText));
                    replaceCount++;
                }
            }
            sr.Close();
            //对html .xml 进行替换
            for (int i = 0; i < oldStrings.Count; i++)
            {
                content = content.Replace(oldStrings[i], newStrings[i]);
            }

            StreamWriter sw = new StreamWriter(filePath, false, Encoding.GetEncoding("gb2312"));
            sw.WriteLine(content);
            sw.Close();
            string fileName = Path.GetFileName(filePath);
            String result = string.Format("在文件：{0}中-----替换了{1}个\"{2}\"", fileName, replaceCount, text);
            return result;
        }
        /// <summary>
        /// 得到符合规则的文件
        /// </summary>
        /// <returns></returns>
        public string[] GetFlies(string fileDirectory)
        {
            if (!Directory.Exists(fileDirectory))
            {
                return null;
            }
            String[] files = Directory.GetFiles(fileDirectory, "*.*", SearchOption.TopDirectoryOnly);
            return files;
        }



        /// <summary>
        /// 勾选那一个
        /// </summary>
        /// <returns></returns>
        public FindEnum GetWhere()
        {
            if (this.check_incontent.Checked)
            {
                return FindEnum.InContent;
            }
            return FindEnum.No;
        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private bool IsOpenedPPT()
        {
            Process[] ps = Process.GetProcesses();
            foreach (Process item in ps)
            {
                
                if (item.ProcessName == "POWERPNT")
                {
                    return true;
                }
            }
            return false;
        }
    }
    /// <summary>
    /// 在内容中还是在题目中
    /// </summary>
    public enum FindEnum
    {
        No = 0,
        InContent = 1,
    };
}
