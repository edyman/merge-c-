using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;

namespace merge
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            Console.WriteLine(File.Exists("ppt/file1.pptx"));
            Console.Read();

            int id = 0;
            using (PresentationDocument myDestDeck = PresentationDocument.Open("ppt/file1.pptx", true)) {
                PresentationPart destPresPart = myDestDeck.PresentationPart;

                if (destPresPart.Presentation.SlideIdList == null)
                    destPresPart.Presentation.SlideIdList = new SlideIdList();
            }
        }
    }
}
