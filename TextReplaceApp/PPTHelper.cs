
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


