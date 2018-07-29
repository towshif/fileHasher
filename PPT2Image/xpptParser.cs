using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GhostscriptSharp;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using iTextSharp.text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Graph; 
using ch = Microsoft.Office.Interop.Graph.Chart;
using System.Text.RegularExpressions;
using System.Data;
// define custom keyValue pair for IDs
using DataPair = System.Collections.Generic.KeyValuePair<string, int>;
using System.Reflection;

namespace fileHasherConverter
{
    public class pptParser
    {

        /* flags*/
        private bool ERROR_PPT2PDF_FULL = false;
        private bool ERROR_PPT2IMG = false;
        private bool ERROR_PPT2TEST = false;

        /* IDs Identifiers */
        private string MOTHER_ID = null;
        private string FILE_ID = null;
        private string FILE_PATH = null;
        private string FILE_TYPE = null;

        /* Paths */
        private string DATA_ROOT = null; 
        private string PDF_ROOT = null;
        private string IMG_ROOT = null;
        private string IMG_CONTENT_ROOT = null;
        private string EXPORT_PATH = null;

        /* Constructor for globals */
        public pptParser()
        {
            ERROR_PPT2PDF_FULL = false;
            ERROR_PPT2IMG = false;
            ERROR_PPT2TEST = false;
            MOTHER_ID = null;
            FILE_ID = null;
            FILE_PATH = null;
            FILE_TYPE = null;

        }
        public pptParser(string motherid, string fileid)
        {
            MOTHER_ID = motherid;
            FILE_ID = fileid;

        }

        public void GetMetaData(string pptfile)
        {
            Console.WriteLine("PPT File Location:" + pptfile);
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptPresentation = pptApplication.Presentations
            .Open(pptfile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            listallbuiltinPropertiesfromPowerpoint(pptPresentation, pptPresentation.BuiltInDocumentProperties);
            pptPresentation.Close();

        }


        public void listallbuiltinPropertiesfromPowerpoint(PowerPoint.Presentation presentation, object builtInProps)
        {
            Console.WriteLine("Writing List of MetaData\n-----------------------");
            try
            {
                var builtinProps = presentation.BuiltInDocumentProperties; // don't strong cast this or you will get null
                //var builtinProps = presentation.CustomDocumentProperties; // don't strong cast this or you will get null
                GetBuiltInProperty(builtinProps);
            }
            catch (Exception e)
            {
                // Ignorer l'erreur
                //Log.Warn("Erreur inattendue à la lecture des propriétés internes du document", e);
            }
            Console.WriteLine("-------------------\nDone. Writing List");
        }

        internal void GetBuiltInProperty(dynamic builtInProps)
        {
            if (builtInProps != null)
            {
                try
                {
                    foreach (var p in builtInProps)
                    {
                        try
                        {
                            Console.Write(p.Name + "(" + p.Type + ") \t:");
                            Console.Write(p.Value + "\n");
                            ;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("ERROR");
                        }
                    }
                }
                catch (Exception e)
                {
                    // Property is missing
                    Console.WriteLine(e.ToString());
                }

            }
        }




        public void pdf2Image(string pdfFile, string exportPath)
        //static void pdf2Image(string pptfile, string prefix)
        {

            // Do not forget the %d in the output file name  @"Example%d.jpg"
            GhostscriptWrapper.GeneratePageThumbs(pdfFile, exportPath, 1, 15, 100, 100);

            // for a single page [you have to know the page number -- no function in ghostscript]
            // GhostscriptWrapper.GeneratePageThumb(exeBase + @"\" + @"d5000.pdf", exeBase + @"\" + @"Example1.jpg", 1, 100, 100);            
        }

        public void pdf2ImageByPage(string pdfFile, string exportPath, int pagenumber)
        //static void pdf2Image(string pptfile, string prefix)
        {
            // Do not forget the %d in the output file name  @"Example%d.jpg"
            GhostscriptWrapper.GeneratePageThumbs(pdfFile, exportPath, pagenumber, pagenumber, 100, 100);
            // for a single page [you have to know the page number -- no function in ghostscript]
            // GhostscriptWrapper.GeneratePageThumb(exeBase + @"\" + @"d5000.pdf", exeBase + @"\" + @"Example1.jpg", 1, 100, 100);            
        }


        /**************************           FINAL        *****************************/
        public string pdf2Text(string pdfFile)
        {
            Console.WriteLine("PDF PATH=" + pdfFile);
            PdfReader reader = new PdfReader(pdfFile);
            //ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();
            ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();

            string txt = "";
            int pages = reader.NumberOfPages;

            for (int page = 1; page <= reader.NumberOfPages; page++)
            {

                txt = PdfTextExtractor.GetTextFromPage(reader, page, its);

                txt = txt.Replace("NOTITLE", "").Replace("\v", " ").Replace("\f", " ") + "\n";
                txt = Regex.Replace(txt, @"[^\t\r\n\u0020-\u007E]+", string.Empty);

                Console.WriteLine("Slide #" + page + "\n-----------------------\n" + txt + "\n-----------------------\n");
            }
            // txt = txt.Replace("\n", " ").Replace("  ", " "); ;


            return txt;
        }

        /**************************           FINAL        *****************************/
        public string pdf2TextByPage(string pdfFile, int pagenumber)
        {
            Console.WriteLine("PDF PATH=" + pdfFile);
            PdfReader reader = new PdfReader(pdfFile);
            //ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();
            ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();

            string txt = PdfTextExtractor.GetTextFromPage(reader, pagenumber, its);
            //txt = txt.Replace("\n", " ");
            // Replace multi spaces with single 
            RegexOptions options = RegexOptions.None;
            Regex regex = new Regex("[ ]{2,}", options);

            txt = regex.Replace(txt, " ");
            Console.WriteLine("Slide #" + pagenumber + "\n-----------------------\n" + txt + "\n-----------------------\n");
            // txt = txt.Replace("\n", " ").Replace("  ", " "); ;            
            return txt;
        }


        // PDF Split 
        public void pdfSplit(string pdffile, string prefix)
        {
        }

        // PDF Merge 
        public void pdfMerge(string pptfile, string prefix)
        {
        }

        public void ppt2pdf(string pptfile, string pdfname, string pdfpath, string prefix)
        {
            Console.WriteLine("PPT File Location:" + pptfile);
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptPresentation = pptApplication.Presentations
            .Open(pptfile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);


            //Save as PDF for text Extraction
            pdfpath = pdfpath + @"\" + pdfname + ".pdf";

            //Publish PPT - to - PDF
            try
            {
                // publishes hidden slide to pdf 
                // pptPresentation.ExportAsFixedFormat(pdfpath, PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint, MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst, PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoTrue);
                object unknownType = Type.Missing;
                if (pptPresentation != null)
                {
                    pptPresentation.ExportAsFixedFormat((string)pdfpath,
                                                         PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                                         PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                                                         MsoTriState.msoFalse,
                                                         PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                                                         PowerPoint.PpPrintOutputType.ppPrintOutputSlides,
                                                         MsoTriState.msoTrue, null,
                                                         PowerPoint.PpPrintRangeType.ppPrintAll, string.Empty,
                                                         true, true, true, true, false, unknownType);
                }
                // do not publish hidden slides to pdf - messes up slide # b/w ppt and pdf 
                // pptPresentation.ExportAsFixedFormat(pdfpath, PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint, MsoTriState.msoFalse);                 
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
                Console.ReadLine();
                return;
            }
        }

        public void ppt2pdfByPage(string pptfile, string path, string prefix, int slideNumber)
        {
            // see Overloaded function for multiple slides. 
            ppt2pdfByPage(pptfile, path, prefix, slideNumber, slideNumber);
        }

        /* PPT 2 PDF by Slide Range : ASP/PowerShell example.
         * https://stackoverflow.com/questions/4086541/powerpoint-exportasfixedformat-in-powershell
         */
        public void ppt2pdfByPage(string pptfile, string path, string prefix, int slideStart, int slideEnd)
        {
            // see Overloaded function for custom slide start - end Ranges. 
            Console.WriteLine("PPT File Location:" + pptfile);
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptPresentation = pptApplication.Presentations
            .Open(pptfile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);


            //Save as PDF for text Extraction
            string pdfpath = pptfile.Replace(".ppt", "") + ".pdf";

            //Publish PPT - to - PDF
            try
            {
                // publishes hidden slide to pdf 
                // pptPresentation.ExportAsFixedFormat(pdfpath, PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint, MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst, PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoTrue);
                object unknownType = Type.Missing;
                if (pptPresentation != null)
                {
                    // define the slide ranges for pdf printing. 
                    PowerPoint.PrintRanges ranges = pptPresentation.PrintOptions.Ranges;
                    PowerPoint.PrintRange range = ranges.Add(slideStart, slideEnd);

                    // note: PowerPoint.PpPrintRangeType.ppPrintSlideRange not PrintAll.

                    pptPresentation.ExportAsFixedFormat((string)pdfpath,
                                                         PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                                         PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                                                         MsoTriState.msoFalse,
                                                         PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                                                         PowerPoint.PpPrintOutputType.ppPrintOutputSlides,
                                                         MsoTriState.msoTrue, range /*null*/,
                                                         PowerPoint.PpPrintRangeType.ppPrintSlideRange, string.Empty,
                                                         true, true, true, true, false, unknownType);
                }
                // do not publish hidden slides to pdf - messes up slide # b/w ppt and pdf 
                // pptPresentation.ExportAsFixedFormat(pdfpath, PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint, MsoTriState.msoFalse);                 
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
                Console.ReadLine();
                return;
            }
        }


        /* PPT 2 PDF by Slide Range : ASP/PowerShell example.
         * https://stackoverflow.com/questions/4086541/powerpoint-exportasfixedformat-in-powershell
         */

        public void ppt2pdfByFilesSeparatePages(string pptfile, string path, List<DataPair> slides)
        {
            // see Overloaded function for custom slide start - end Ranges. 
            Console.WriteLine("PPT File Location:" + pptfile);
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptPresentation = pptApplication.Presentations
            .Open(pptfile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);


            //Save as PDF for text Extraction
            string pdfpath = pptfile.Replace(".ppt", "") + ".pdf";
            //string pdfpath = @path+"\\"+ id + ".pdf";

            //Publish PPT - to - PDF

            // publishes hidden slide to pdf 
            // pptPresentation.ExportAsFixedFormat(pdfpath, PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint, MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst, PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoTrue);
            object unknownType = Type.Missing;
            if (pptPresentation != null)
            {
                foreach (DataPair slide in slides)
                {
                    pdfpath = @path + @"\\" + slide.Key + ".pdf";
                    // define the slide ranges for pdf printing. 

                    try
                    {
                        PowerPoint.PrintRanges ranges = pptPresentation.PrintOptions.Ranges;
                        PowerPoint.PrintRange range = ranges.Add(slide.Value, slide.Value);

                        // note: PowerPoint.PpPrintRangeType.ppPrintSlideRange not PrintAll.

                        pptPresentation.ExportAsFixedFormat((string)pdfpath,
                                                             PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                                             PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                                                             MsoTriState.msoFalse,
                                                             PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                                                             PowerPoint.PpPrintOutputType.ppPrintOutputSlides,
                                                             MsoTriState.msoTrue, range /*null*/,
                                                             PowerPoint.PpPrintRangeType.ppPrintSlideRange, string.Empty,
                                                             true, true, true, true, false, unknownType);
                    }
                    // do not publish hidden slides to pdf - messes up slide # b/w ppt and pdf 
                    // pptPresentation.ExportAsFixedFormat(pdfpath, PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint, MsoTriState.msoFalse);                 
                    catch (Exception ex)
                    {
                        Console.Write(ex.ToString());
                        Console.ReadLine();
                        return;
                    }
                }
            }

        }



        /**************************           FINAL PPT-to-Image       *****************************
         * http://www.free-power-point-templates.com/articles/c-code-to-convert-powerpoint-to-image/         
         */

        public void ppt2Image(string pptfilePath, string exportPath, string prefix)
        {
            Console.WriteLine("PPT File Location:" + pptfilePath);
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptPresentation = pptApplication.Presentations
            .Open(pptfilePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            int pageWidth = 800;
            int pageHeight = (int)(pageWidth / pptPresentation.PageSetup.SlideWidth * pptPresentation.PageSetup.SlideHeight);

            Console.WriteLine("w/H" + pageWidth + ":" + pageHeight);

            //slides to image
            int slide_count = pptPresentation.Slides.Count;
            Console.Write("count=" + slide_count);
            for (int i = 1; i <= slide_count; ++i)
            {
                /* full HD -  Statistical Size: GIF < JPG < PNG */
                pptPresentation.Slides[i].Export(exportPath + @"\" + "slide" + i + ".gif", "gif", pageWidth, pageHeight);
                /* Thumbnail 3x*/
                pptPresentation.Slides[i].Export(exportPath + @"\" + "thumb.slide" + i + "_3x.gif", "gif", (int)(pageWidth / 3.2), (int)(pageHeight / 3.2));
                /* Thumbnail 6x*/
                pptPresentation.Slides[i].Export(exportPath + @"\" + "thumb.slide" + i + "_6x.gif", "gif", (int)(pageWidth / 6), (int)(pageHeight / 6));

            }

            //ppt2ContentImages(pptPresentation, exportPath.Replace(@"\ppt-img", @"\" + "content-img"));

            pptPresentation.Close();
        }


        void ppt2ContentImages(Presentation pptPresentation, string exportPath)
        {
            //slide extract image content
            // https://stackoverflow.com/questions/4990825/export-movies-from-powerpoint-to-file-in-c-sharp
            //https://stackoverflow.com/questions/42442659/c-sharp-save-ppt-shape-msopicture-as-image-w-office-interop

            foreach (PowerPoint.Slide slide in pptPresentation.Slides)
            {
                PowerPoint.Shapes slideShapes = slide.Shapes;
                int count = 0;
                try
                {
                    foreach (PowerPoint.Shape shape in slideShapes)
                    {
                        if (shape.Type == MsoShapeType.msoPicture)
                        {
                            //LinkFormat.SourceFullName contains the movie path 
                            //get the path like this
                            shape.Export(exportPath + @"\" + "content" + slide.SlideNumber + "_" + count++ + ".png", Microsoft.Office.Interop.PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                            Console.WriteLine("Exported" + exportPath + @"\" + "content" + slide.SlideNumber + "_" + count++ + ".png");
                            //System.IO.File.Copy(shape.LinkFormat.SourceFullName, path + imageBase + @"\" + "content" + slide.SlideNumber + "_"+ count++ + ".png"); 
                        }
                    }
                }
                catch (Exception e)
                {
                    // TODO
                }
            }
        }


        /**************************           FINAL PPT- Content Images - Extract        *****************************
        //slide extract image content
        // https://stackoverflow.com/questions/4990825/export-movies-from-powerpoint-to-file-in-c-sharp
        //https://stackoverflow.com/questions/42442659/c-sharp-save-ppt-shape-msopicture-as-image-w-office-interop         
        */

        public void pptContentExtract(string src_pptfilePath, string exportPath, string prefix)
        {
            Console.WriteLine("PPT File Location:" + src_pptfilePath);
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptPresentation = pptApplication.Presentations
            .Open(src_pptfilePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            foreach (PowerPoint.Slide slide in pptPresentation.Slides)
            {
                PowerPoint.Shapes slideShapes = slide.Shapes;
                int count = 0;
                foreach (PowerPoint.Shape shape in slideShapes)
                {
                    if (shape.Type == MsoShapeType.msoPicture)
                    {
                        //LinkFormat.SourceFullName contains the movie path 
                        //get the path like this
                        shape.Export(exportPath + @"\" + "content" + slide.SlideNumber + "_" + count++ + ".png", Microsoft.Office.Interop.PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                        Console.WriteLine("Exported" + exportPath + @"\" + "content" + slide.SlideNumber + "_" + count++ + ".png");
                        //System.IO.File.Copy(shape.LinkFormat.SourceFullName, path + imageBase + @"\" + "content" + slide.SlideNumber + "_"+ count++ + ".png"); 
                    }
                }
            }
            pptPresentation.Close();
        }



        /* Version not working accurately [DO NOT USE]
         * PPT to Text 
         */

        public void ppt2text(string pptfile)
        {

            Console.WriteLine("PPT File Location:" + pptfile);
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptPresentation = pptApplication.Presentations
            .Open(pptfile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            System.IO.File.WriteAllText(@".\OutText.md", "#Summary of slides:\n");


            foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in pptPresentation.Slides)
            {
                //if (slide.SlideNumber > 20) return;

                string pps = "";

                string slide_title = "NOTITLE"; //  slide.Shapes.Title.TextFrame.TextRange.Text;


                //string slide_title = slide.Shapes.Title.TextFrame.TextRange.Text;

                try
                {
                    if (slide.Shapes.Title.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        var textFrame = slide.Shapes.Title.TextFrame;
                        if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            var textRange = textFrame.TextRange;
                            var paragraphs = textRange.Paragraphs(-1, -1);
                            foreach (PowerPoint.TextRange paragraph in paragraphs)
                            {
                                var text = paragraph.Text;
                                text = text.Replace("\r", "");
                                text = Regex.Replace(text, @"[^\t\r\n\u0020-\u007E]+", string.Empty);
                                text = Regex.Replace(text, @"[ ]{ 2,}", string.Empty);
                                if (text.Length > 2)
                                {
                                    slide_title = slide_title.Replace("NOTITLE", "").Replace("\v", " ").Replace("\f", " ") + text + "\n";
                                    slide_title = Regex.Replace(slide_title, @"[^\t\r\n\u0020-\u007E]+", string.Empty);
                                }

                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }


                Console.WriteLine("@" + slide_title);

                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    pps += readShape(shape);
                }

                pps = Regex.Replace(pps, @"[^\t\r\n\u0020-\u007E]+", string.Empty);


                Console.WriteLine("Slide #" + slide.SlideNumber + "\n-----------------------\n" + pps + "\n-----------------------\n");

                System.IO.File.AppendAllText(@".\OutText.md", "---\nSlide #" + slide.SlideNumber + "\n# " + slide_title);
                System.IO.File.AppendAllText(@".\OutText.md", "\n" + pps + "\n");
            }
            pptPresentation.Close();
        }

        private string readShape(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            string str = "";

            // extract text 
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                var textFrame = shape.TextFrame;
                if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    var textRange = textFrame.TextRange;
                    var paragraphs = textRange.Paragraphs(-1, -1);
                    foreach (PowerPoint.TextRange paragraph in paragraphs)
                    {
                        var text = paragraph.Text;
                        text = text.Replace("\r", "").Replace("\v", " ").Replace("\f", " ").Trim();
                        text = Regex.Replace(text, @"[^\t\r\n\u0020-\u007E]+", string.Empty);
                        text = Regex.Replace(text, @"[ ]{ 2,}", " ");
                        if (text.Length > 2)
                            str += "* " + text + "\n";
                    }
                }
            }
            // read groups -> ungroup and recursively iterate through this function
            else if (shape.Type == MsoShapeType.msoGroup)
            {
                var p = shape.Ungroup();
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shp in p)
                {
                    str += readShape(shp);
                }
            }
            //read tables in ppt
            else if (shape.HasTable == MsoTriState.msoTrue)
            {
                str += "\n";
                var t = shape.Table;

                for (int j = 1; j <= t.Rows.Count; ++j)
                {
                    // create the table header for the .md format 
                    if (j == 2)
                    {
                        str += "| ";
                        for (int k = 1; k <= t.Columns.Count; ++k)
                            str += "---| ";
                        str += "\n";
                    }
                    // for every other row enter |---|---| notation for cells.
                    str += "| ";
                    for (int k = 1; k <= t.Columns.Count; ++k)
                    {
                        var textFrame = t.Cell(j, k).Shape.TextFrame;
                        var textRange = textFrame.TextRange;
                        var paragraphs = textRange.Paragraphs(-1, -1);
                        foreach (PowerPoint.TextRange paragraph in paragraphs)
                        {
                            var text = paragraph.Text;
                            text = text.Replace("\r", "").Replace("\v", " ").Replace("\f", " ");
                            text = Regex.Replace(text, @"[^\t\r\n\u0020-\u007E]+", string.Empty);
                            text = Regex.Replace(text, @"[ ]{ 2,}", string.Empty);
                            str += " " + text;
                        }
                        str += " | ";
                    }
                    str += "\n";
                }
                str += "\n~\n";
            }
            else if (shape.HasChart == MsoTriState.msoTrue)
            {
                Console.WriteLine("Has Chart: True");
                Microsoft.Office.Interop.PowerPoint.Chart t = shape.Chart;
                //string text = ""; 
                if (t.HasTitle)
                {
                    Console.WriteLine("Has Title: True");
                    Console.WriteLine("Title:" + t.ChartTitle.Text.ToString());
                    str += t.ChartTitle.Text.ToString() + " ";
                }

                if (t.HasDataTable)
                {
                    Console.WriteLine("Has DataTable: True");
                    System.Data.DataTable dp = (System.Data.DataTable)shape.Chart.DataTable;
                    string strRowCommaSeparated = "";
                    foreach (DataRow dr in dp.Rows)
                    {
                        for (int i = 0; i < dr.ItemArray.Length; i++)
                        {
                            strRowCommaSeparated += strRowCommaSeparated == "" ? dr.ItemArray[i].ToString() : ("," + strRowCommaSeparated);
                        }
                    }
                    Console.WriteLine("\n\n\t\tOur DataTable : " + strRowCommaSeparated);

                    var p = t.DataTable;
                    str += t.DataTable.ToString(); //.DataTable.ToString();
                }

                // Extracting series labels and count 
                Microsoft.Office.Interop.PowerPoint.SeriesCollection tmp = (Microsoft.Office.Interop.PowerPoint.SeriesCollection)t.SeriesCollection();
                Console.WriteLine("Series Count:" + tmp.Count);
                t.ApplyDataLabels();
                for (int j = 1; j <= tmp.Count; ++j)
                {
                    Microsoft.Office.Interop.PowerPoint.Series aSeries = tmp.Item(j);
                    Console.WriteLine("Series Name: " + aSeries.Name);
                    //Microsoft.Office.Interop.PowerPoint.Points pts = tmp.Item(j).Points(); // # point in the chart per series - meaning X axis for a pareto chart
                    //Console.WriteLine("Points #:" + pts.Count);
                    //Console.WriteLine(pts.Item(2).Name);
                }

                //Excel Route for chart series data extraction 
                PowerPoint.ChartData pChartData = t.ChartData;
                //Console.WriteLine("Chart Title:" + t.Title); // No need - shape.title takes care of this 
                if (!t.ChartData.IsLinked)
                {
                    Console.WriteLine("Has Embedded Excel: True");
                    Excel.Workbook eWorkbook = (Excel.Workbook)pChartData.Workbook;
                    Excel.Worksheet eWorksheet = (Excel.Worksheet)eWorkbook.Worksheets[1];
                    var columnsRange = eWorksheet.UsedRange.Columns;
                    var rowsRange = eWorksheet.UsedRange.Rows;
                    var columnCount = columnsRange.Columns.Count;
                    var rowCount = rowsRange.Rows.Count;
                    //Console.WriteLine("r#, c# :  " + rowCount + ":" + columnCount);

                    foreach (Excel.Range c in eWorksheet.UsedRange)
                    {
                        string changedCell = c.get_Address(Type.Missing, Type.Missing, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                        Console.Write(" | " + c.Value2);
                    }
                    eWorkbook.Close();
                }
                else if (t.ChartData.IsLinked)
                {
                    // add a note for PDF extraction and flagging
                }
                
            }

            //else if (shape.Type == MsoShapeType.msoTextBox || shape.Type == MsoShapeType.msoAutoShape)
            //{
            //    var textFrame = shape.TextFrame;
            //    if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            //    {
            //        var textRange = textFrame.TextRange;
            //        var paragraphs = textRange.Paragraphs(-1, -1);
            //        foreach (PowerPoint.TextRange paragraph in paragraphs)
            //        {
            //            var text = paragraph.Text;
            //            text = text.Replace("\r", "");
            //            str += text + "\n";
            //        }
            //    }
            //}

            str = Regex.Replace(str, @"[^\t\r\n\u0020-\u007E]+", string.Empty);
            return str; //.Replace("\n\n", "\n"); //.Replace("\r",""); 
        }


        /**************************           FINAL        *****************************/
        void mergePPTs(string pptfile2, string pptfile1)
        {
            Console.WriteLine("PPT File Location:" + pptfile1 + ", " + pptfile2);

            PowerPoint.Application pptApplication1 = new Microsoft.Office.Interop.PowerPoint.Application();
            PowerPoint.Application pptApplication2 = new Microsoft.Office.Interop.PowerPoint.Application();

            PowerPoint.Application app = new PowerPoint.Application();

            Presentation pptPresentation1 = pptApplication1.Presentations.Open(pptfile1, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            Presentation pptPresentation2 = pptApplication1.Presentations.Open(pptfile2, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            //Take the first PPT and and merge rest.
            pptPresentation1.SaveAs(EXPORT_PATH + @"\" + @"temp.pptx", PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            pptPresentation1.Close();

            PowerPoint.Presentation mergedPPT = app.Presentations.Open(EXPORT_PATH + @"\" + @"temp.pptx", MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);

            int slide_count = 5; // pptPresentation1.Slides.Count;
            Console.Write("count=" + slide_count);

            //mergedPPT.Slides.InsertFromFile(pptfile1, 1, -1);

            mergedPPT.Slides.InsertFromFile(pptfile2, 0, 1, -1);

            mergedPPT.SaveAs(EXPORT_PATH + @"\" + @"merged.pptx", PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            mergedPPT.Close();
            app.Quit();

            for (int i = 1; i <= slide_count; ++i)
            {
                /* full HD*/
                //pptPresentation1.Slides[i].Export(imageBase + @"\" + prefix + "slide" + i + ".png", "png", 800, 600);

                /* Thumbnail*/
                //pptPresentation1.Slides[i].Export(imageBase + @"\" + prefix + "thumb.slide" + i + ".png", "png", 320, 240);

            }
        }
        
    }
}

