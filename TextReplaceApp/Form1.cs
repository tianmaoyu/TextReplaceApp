using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
        #region 没有用的事件
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cob_path_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }
        #endregion


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

        private void btn_find_Click(object sender, EventArgs e)
        {
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
                    message += FindTextInContent(fileText, fliePath) + "\n";
                }
                MessageBox.Show(message);
            }

        }

        private void btn_replace_Click(object sender, EventArgs e)
        {
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
                    message += ReplaceTextInContent(fileText, repalceText, fliePath) + "\n";
                }
                MessageBox.Show(message);
            }
        }
        /// <summary>
        /// 的到查找内容的详情
        /// </summary>
        /// <param name="text"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public string FindTextInContent(string text, string filePath)
        {
            if (!File.Exists(filePath))
            {
                return "没找到" + filePath;
            }

            StreamReader sr = new StreamReader(filePath, Encoding.UTF8);
            string line;
            int i = 0;
            while ((line = sr.ReadLine()) != null)
            {
                if (line.Contains(text))
                {
                    i++;

                }
            }
            string fileName = Path.GetFileName(filePath);
            sr.Close();
            string reslut = string.Format("内容中:{0} 找到 {1} 个:\"{2}\"", fileName, i, text);
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
            StreamReader sr = new StreamReader(filePath, Encoding.UTF8);
            string line;
            int i = 0;
            while ((line = sr.ReadLine()) != null)
            {
                if (line.Contains(text))
                {
                    i++;
                }
            }
            sr.Close();
            return i.ToString();
        }


        public string ReplaceTextInContent(string text, string newText, string filePath)
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
            StreamReader sr = new StreamReader(filePath, Encoding.UTF8);
            con = sr.ReadToEnd();
            con = con.Replace(text, newText);
            sr.Close();
            StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8);
            sw.WriteLine(con);
            sw.Close();

            countStr = FindTextCountInContent(newText, filePath);
            Int32.TryParse(countStr, out newCount);

            string fileName = Path.GetFileName(filePath);
            string reslut = string.Format("在文件{0}中一共替换了{1}个{2}", fileName, newCount - oldCount, text);
            return reslut;
        }


        /// <summary>
        /// 得到html.xml 这种文件要替换的字符串，和新的字符串
        /// </summary>
        /// <param name="oldStrings"></param>
        /// <param name="newStrings"></param>
        public String GRelaceForHTMLOrXML(string text, string newText, string filePath)
        {
            List<string> oldStrings = new List<string>();
            List<string> newStrings = new List<string>();
            int replaceCount = 0;
            string content = "";
            StreamReader sr = new StreamReader(filePath, Encoding.UTF8);
            content = sr.ReadToEnd();
            Regex regex = new Regex(@"(?<=>).*?(?=<)");
            var matches = regex.Matches(content);
            foreach(Match match in matches)
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
            for(int i=0;i< oldStrings.Count; i++)
            {
                content.Replace(oldStrings[i], newStrings[i]);
            }
            StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8);
            sw.WriteLine(content);
            sw.Close();
            string fileName = Path.GetFileName(filePath);
            String result = string.Format("标题中：{0}的题目中找到{1}个\"{2}\"", fileName, replaceCount, text);
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
            String[] files = Directory.GetFiles(fileDirectory, "*.txt", SearchOption.TopDirectoryOnly);//SearchOption.AllDirectories
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
