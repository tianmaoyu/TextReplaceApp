
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

namespace TextReplaceApp
{

    //public class PPTHelper
    //{
    //    public string Replace(string filePyth)
    //    {
    //        Presentation pres = new Presentation(filePyth);
    //        for (int j = 0; j < pres.Slides.Count; j++)
    //        {
    //            Slide slide = (Slide)pres.Slides[j];
    //            for (int i = 0; i < slide.Shapes.Count; i++)
    //            {
    //                Aspose.Slides.Shape shape = (Shape)slide.Shapes[i];
    //                if (shape.TextFrame != null)
    //                {
    //                    TextFrame textFrame = shape.TextFrame;
    //                    for (int par = 0; par < textFrame.Paragraphs.Count; par++)
    //                    {
    //                        Paragraph paragraph = textFrame.Paragraphs[par];
    //                        for (int por = 0; por < paragraph.Portions.Count; por++)
    //                        {
    //                            Portion portion = paragraph.Portions[por];
    //                            portion.Text = portion.Text.Replace("你好", "hello!");
    //                        }
    //                    }
    //                }
    //            }
    //        }

    //        pres.Save(filePyth, SaveFormat.Ppt);
    //        return null;
    //    }

    //}
}

//public class OperatePPT
//{
//    #region=========基本的参数信息=======
//    POWERPOINT.Application objApp = null;
//    POWERPOINT.Presentation objPresSet = null;
//    POWERPOINT.SlideShowWindows objSSWs;
//    POWERPOINT.SlideShowTransition objSST;
//    POWERPOINT.SlideShowSettings objSSS;
//    POWERPOINT.SlideRange objSldRng;
//    bool bAssistantOn;

//    #endregion
//    #region===========操作方法==============
//    /// <summary>
//    /// 打开PPT文档并播放显示。
//    /// </summary>
//    /// <param name="filePath">PPT文件路径</param>
//    public void PPTOpen(string filePath)
//    {
//        //防止连续打开多个PPT程序.
//        if (this.objApp != null) { return; }
//        try
//        {
//            objApp = new POWERPOINT.Application();
//            objApp.Visible = OFFICECORE.MsoTriState.msoTrue;

//            //以非只读方式打开,方便操作结束后保存.
//            objPresSet = objApp.Presentations.Open(filePath, OFFICECORE.MsoTriState.msoFalse);

//        }
//        catch (Exception ex)
//        {
//            MessageBox.Show("错误:" + ex.Message.ToString());
//            this.objApp.Quit();

//        }
//    }
//    /// <summary>
//    /// 自动播放PPT文档.
//    /// </summary>
//    /// <param name="filePath">PPTy文件路径.</param>
//    /// <param name="playTime">翻页的时间间隔.【以秒为单位】</param>
//    public void PPTAuto(string filePath, int playTime)
//    {
//        //防止连续打开多个PPT程序.
//        if (this.objApp != null) { return; }
//        objApp = new POWERPOINT.Application();
//        objPresSet = objApp.Presentations.Open(filePath, OFFICECORE.MsoTriState.msoCTrue, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse);
//        // 自动播放的代码（开始）
//        int Slides = objPresSet.Slides.Count;
//        int[] SlideIdx = new int[Slides];
//        for (int i = 0; i < Slides; i++) { SlideIdx[i] = i + 1; };
//        objSldRng = objPresSet.Slides.Range(SlideIdx);
//        objSST = objSldRng.SlideShowTransition;
//        //设置翻页的时间.
//        objSST.AdvanceOnTime = OFFICECORE.MsoTriState.msoCTrue;
//        objSST.AdvanceTime = playTime;
//        //翻页时的特效!
//        objSST.EntryEffect = POWERPOINT.PpEntryEffect.ppEffectCircleOut;
//        //Prevent Office Assistant from displaying alert messages:
//        bAssistantOn = objApp.Assistant.On;
//        objApp.Assistant.On = false;
//        //Run the Slide show from slides 1 thru 3.
//        objSSS = objPresSet.SlideShowSettings;
//        objSSS.StartingSlide = 1;
//        objSSS.EndingSlide = Slides;
//        objSSS.Run();
//        //Wait for the slide show to end.
//        objSSWs = objApp.SlideShowWindows;
//        while (objSSWs.Count >= 1) System.Threading.Thread.Sleep(playTime * 100);
//        this.objPresSet.Close();
//        this.objApp.Quit();
//    }
//    /// <summary>
//    /// PPT下一页。
//    /// </summary>
//    public void NextSlide()
//    {
//        if (this.objApp != null)
//            try
//            {
//                this.objPresSet.SlideShowWindow.View.Next();
//            }
//            catch
//            { }
//    }
//    /// <summary>
//    /// PPT上一页。
//    /// </summary>
//    public void PreviousSlide()
//    {
//        if (this.objApp != null)
//            this.objPresSet.SlideShowWindow.View.Previous();
//    }

//    private int PageNum()
//    {
//        return objPresSet.Slides.Count;

//    }

//    public void SetLine()
//    {
//        int num = PageNum();
//        for (int i = 0; i < num; i++)
//        {
//            if (i > 2)
//            {
//                objSldRng = objPresSet.Slides.Range(i);
//                objSldRng.Select();
//                try
//                {
//                    objSldRng.Application.ActiveWindow.Selection.SlideRange.Shapes.SelectAll();
//                    objSldRng.Application.ActiveWindow.Selection.ShapeRange.Line.Visible = OFFICECORE.MsoTriState.msoFalse;
//                }
//                catch
//                { }
//                //MessageBox.Show("" + i.ToString());


//                //NextSlide();
//            }

//        }

//    }


//    /// <summary>
//    /// 关闭PPT文档。
//    /// </summary>
//    public void PPTClose()
//    {
//        //装备PPT程序。
//        if (this.objPresSet != null)
//        {
//            //判断是否退出程序,可以不使用。
//            //objSSWs = objApp.SlideShowWindows;
//            //if (objSSWs.Count >= 1)
//            //{
//            //if (MessageBox.Show("是否保存修改的笔迹!", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
//            this.objPresSet.Save();
//            //}
//            //this.objPresSet.Close();
//        }
//        if (this.objApp != null)
//            this.objApp.Quit();
//        GC.Collect();
//    }
//    #endregion
//}
//}
