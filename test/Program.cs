using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace ConsoleApplication1
{
    class Program
    {
        public static string DirectoryPath = @"C:\projekty\";
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            ApplicationClass pptApplication = new ApplicationClass();
            Presentation pptPresentation = pptApplication.Presentations.Open(DirectoryPath + "testfile.pptx", MsoTriState.msoFalse,
                MsoTriState.msoFalse, MsoTriState.msoFalse);
            try
            {
                int i = 1; //slides index starts from 1
                List<string> fileNames = new List<string>();
                foreach (var slide in pptPresentation.Slides) 
                {
                    var slideFileName = $"slide{i}.jpg";
                    pptPresentation.Slides[i].Export(slideFileName, "jpg"); //exports to directory where .pptx file exists

                    fileNames.Add(DirectoryPath + slideFileName);
                    i++;
                }

                PdfHelper.Instance.SaveImagesAsPdf(fileNames, DirectoryPath + "prezentacja.pdf");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                pptPresentation.Close();
            }
        }
    }
}
