using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.Drawing;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties;

namespace OpenXmlPptxCreator
{

    public static class Program
    {
        static void Main(string[] args)
        {
            string filepath = @"e:\temp\my1.pptx";

            using (PpPresentation presentation = new PpPresentation(filepath, 612 * 10000, 792 * 10000))
            {
                presentation.AddSlideWithImage(@"e:\temp\images\page1.png");
                presentation.AddSlideWithImage(@"e:\temp\images\page2.png");
                presentation.AddSlideWithImage(@"e:\temp\images\page3.png");
            }
        }
    }
}
