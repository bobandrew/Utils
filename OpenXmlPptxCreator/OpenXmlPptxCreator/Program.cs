
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
