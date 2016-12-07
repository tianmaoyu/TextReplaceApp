﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
namespace TextReplaceApp
{
    public class WordHelper
    {
        Word.Application word;
        public WordHelper()
        {
            word = new Microsoft.Office.Interop.Word.Application();
        }
        /// <summary>
        /// 在word 中查找替换操作
        /// </summary>
        /// <param name="fileText"></param>
        /// <param name="repalceText"></param>
        /// <param name="fliePath"></param>
        /// <returns></returns>
        public string ReplaceInWord(string fileText, string repalceText, string fliePath)
        {
            object unknow = Type.Missing;
            Word.Document doc = null;
            word.Visible = true;
            object file = fliePath;
            doc = word.Documents.Open(ref file,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow);

            word.Selection.Find.Replacement.ClearFormatting();
            word.Selection.Find.ClearFormatting();
            word.Selection.Find.Text = fileText;//需要被替换的文本
            word.Selection.Find.Replacement.Text = repalceText;//替换文本 
            object oMissing = System.Reflection.Missing.Value;
            object replace = Word.WdReplace.wdReplaceAll;

            string contenText = doc.Content.Text;
            Regex regex = new Regex(fileText);
            var matches = regex.Matches(contenText);
          
            //执行替换操作
            word.Selection.Find.Execute(
            ref oMissing, ref oMissing,
            ref oMissing, ref oMissing,
            ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref replace,
            ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            //doc.Content.Text = doc.Content.Text.Replace("第","hello");
            //foreach()
            doc.Save();
            doc.Close();
            word.Quit();
            string fileName = Path.GetFileName(fliePath);
            String result = string.Format("在文件：{0}中-----找到{1}个\"{2}\"", fileName, matches.Count, fileText);
            return result;
        }


        /// <summary>
        /// 查找在word 中
        /// </summary>
        /// <param name="text"></param>
        /// <param name="wordPath"></param>
        /// <returns></returns>
        public string FindInWord(string text, string wordPath)
        {

            object unknow = Type.Missing;
            Word.Document doc = null;
            word.Visible = true;
            object file = wordPath;
            doc = word.Documents.Open(ref file,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow);
            doc.Close();
            word.Quit();
            string contenText = doc.Content.Text;
            Regex regex = new Regex(text);
            var matches = regex.Matches(contenText);
            string fileName = Path.GetFileName(wordPath);
            String result = string.Format("在文件：{0}中-----找到{1}个\"{2}\"", fileName, matches.Count, text);
            return result;
        }
    }
}
