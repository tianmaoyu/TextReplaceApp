using System;
using System.Collections.Generic;
using System.Text;
using OFFICECORE = Microsoft.Office.Core;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;
using System.Windows;
using System.Collections;
using System.Windows.Forms;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace pptWrite
{

    /// <summary>
    /// PPT文档操作实现类.
    /// </summary>
    public class OperatePPT
    {
        #region=========基本的参数信息=======
        POWERPOINT.Application objApp = null;
        POWERPOINT.Presentation objPresSet = null;
        POWERPOINT.SlideShowWindows objSSWs;
        POWERPOINT.SlideShowTransition objSST;
        POWERPOINT.SlideShowSettings objSSS;
        POWERPOINT.SlideRange objSldRng;
        bool bAssistantOn;

        #endregion
        #region===========操作方法==============
        /// <summary>
        /// 打开PPT文档并播放显示。
        /// </summary>
        /// <param name="filePath">PPT文件路径</param>
        public void PPTOpen(string filePath)
        {
            //防止连续打开多个PPT程序.
            if (this.objApp != null) { return; }
            try
            {
                objApp = new POWERPOINT.Application();
                objApp.Visible = OFFICECORE.MsoTriState.msoTrue;
                
                //以非只读方式打开,方便操作结束后保存.
                objPresSet = objApp.Presentations.Open(filePath, OFFICECORE.MsoTriState.msoFalse);
                //假装隐藏
             
            }
            catch (Exception ex)
            {
                MessageBox.Show("错误:" + ex.Message.ToString());
                this.objApp.Quit();

            }
        }
        ///// <summary>
        ///// 自动播放PPT文档.
        ///// </summary>
        ///// <param name="filePath">PPTy文件路径.</param>
        ///// <param name="playTime">翻页的时间间隔.【以秒为单位】</param>
        //public void PPTAuto(string filePath, int playTime)
        //{
        //    //防止连续打开多个PPT程序.
        //    if (this.objApp != null) { return; }
        //    objApp = new POWERPOINT.Application();
        //    objPresSet = objApp.Presentations.Open(filePath, OFFICECORE.MsoTriState.msoCTrue, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse);
        //    // objPresSet = objApp.Presentations.Add(filePath, OFFICECORE.MsoTriState.msoCTrue, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse);
        //    // 自动播放的代码（开始）


        //    int Slides = objPresSet.Slides.Count;
        //    int[] SlideIdx = new int[Slides];
        //    for (int i = 0; i < Slides; i++) { SlideIdx[i] = i + 1; };
        //    objSldRng = objPresSet.Slides.Range(SlideIdx);
        //    objSST = objSldRng.SlideShowTransition;
        //    //设置翻页的时间.
        //    objSST.AdvanceOnTime = OFFICECORE.MsoTriState.msoCTrue;
        //    objSST.AdvanceTime = playTime;
        //    //翻页时的特效!
        //    objSST.EntryEffect = POWERPOINT.PpEntryEffect.ppEffectCircleOut;
        //    //Prevent Office Assistant from displaying alert messages:
        //    bAssistantOn = objApp.Assistant.On;
        //    objApp.Assistant.On = false;
        //    //Run the Slide show from slides 1 thru 3.
        //    objSSS = objPresSet.SlideShowSettings;
        //    objSSS.StartingSlide = 1;
        //    objSSS.EndingSlide = Slides;
        //    objSSS.Run();
        //    //Wait for the slide show to end.
        //    objSSWs = objApp.SlideShowWindows;
        //    while (objSSWs.Count >= 1) System.Threading.Thread.Sleep(playTime * 100);
        //    this.objPresSet.Close();
        //    this.objApp.Quit();
        //}
        /// <summary>
        /// PPT下一页。
        /// </summary>
        public void NextSlide()
        {
            if (this.objApp != null)
                try
                {
                    this.objPresSet.SlideShowWindow.View.Next();
                }
                catch
                { }
        }
        /// <summary>
        /// PPT上一页。
        /// </summary>
        public void PreviousSlide()
        {
            if (this.objApp != null)
                this.objPresSet.SlideShowWindow.View.Previous();
        }

        private int PageNum()
        {
            return objPresSet.Slides.Count;

        }

        public void SetLine()
        {
            int num = PageNum();
            for (int i = 0; i < num; i++)
            {
                if (i > 2)
                {
                    objSldRng = objPresSet.Slides.Range(i);
                    objSldRng.Select();
                    try
                    {
                        objSldRng.Application.ActiveWindow.Selection.SlideRange.Shapes.SelectAll();
                        objSldRng.Application.ActiveWindow.Selection.ShapeRange.Line.Visible = OFFICECORE.MsoTriState.msoFalse;
                    }
                    catch
                    { }
                    //MessageBox.Show("" + i.ToString());


                    //NextSlide();
                }

            }

        }


        /// <summary>
        /// 关闭PPT文档。
        /// </summary>
        public void PPTClose()
        {
            //装备PPT程序。
            if (this.objPresSet != null)
            {
                //判断是否退出程序,可以不使用。
                //objSSWs = objApp.SlideShowWindows;
                //if (objSSWs.Count >= 1)
                //{
                //if (MessageBox.Show("是否保存修改的笔迹!", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
                this.objPresSet.Save();
                //}
                //this.objPresSet.Close();
            }
            if (this.objApp != null)
            {
                objApp.Quit();
            }
            GC.Collect();
            Process[] ps = Process.GetProcesses();
            foreach (Process item in ps)
            {
               
                {
                    var mae = item.MainWindowTitle;
                    if (item.MainWindowTitle.Contains("PowerPoint"))
                    {
                        ;
                    }
                }
                if (item.ProcessName == "POWERPNT")
                {
                    item.Kill();
                }
            }
         
           
        }
        #endregion

        public void ReplaceAll(string OldText, string NewText)
        {
            int num = PageNum();
            for (int j = 1; j <= num; j++)
            {
                POWERPOINT.Slide slide = objPresSet.Slides[j];
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    POWERPOINT.Shape shape = slide.Shapes[i];
                    if (shape.TextFrame != null)
                    {
                        POWERPOINT.TextFrame textFrame = shape.TextFrame;
                        try
                        {
                            if (textFrame.TextRange != null)
                            {
                                string text = textFrame.TextRange.Text;
                                Regex regex = new Regex(NewText);
                                var matches = regex.Matches(text);
                                //需求替换的次数
                                 foreach(var macth in matches)
                                {
                                    textFrame.TextRange.Replace(OldText, NewText);
                                }
                           
                            }
                        }
                        catch
                        { }
                    }
                }
            }
        }

        public string FindInPPT(string oldText, string filePath)
        {
            int num = PageNum();
            string text = "";
            for (int j = 1; j <= num; j++)
            {
                POWERPOINT.Slide slide = objPresSet.Slides[j];
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    POWERPOINT.Shape shape = slide.Shapes[i];
                    if (shape.TextFrame != null)
                    {
                        POWERPOINT.TextFrame textFrame = shape.TextFrame;
                        try
                        {
                            if (textFrame.TextRange != null)
                            {
                                text += textFrame.TextRange.Text;
                            }
                        }
                        catch
                        { }
                    }
                }
            }
            Regex regex = new Regex(oldText);
            var matches = regex.Matches(text);
            string fileName = Path.GetFileName(filePath);
            String result = string.Format("在文件：{0}中-----找到{1}个\"{2}\"", fileName, matches.Count, oldText);
            return result;
        }
    }
}