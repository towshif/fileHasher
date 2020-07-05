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
//using ch = Microsoft.Office.Interop.Graph.Chart;
using System.Text.RegularExpressions;
using System.Data;
// define custom keyValue pair for IDs
using DataPair = System.Collections.Generic.KeyValuePair<string, int>;


namespace fileHasherConverter
{
    public class __BACKUP__FileHash
    {
        //public string imageBase = @"H:\output";
        private  string imageBase = ".";
        private static string exeBase = ".";
        //public DBConnect db;
        private static string pa = System.Reflection.Assembly.GetExecutingAssembly().Location;
        private static string dd = System.IO.Path.GetDirectoryName(pa);
        //exeBase = dd.ToString();
            
       // For testing only. 
        public void hashMe()
        {
            string pa = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var dd = System.IO.Path.GetDirectoryName(pa);
            exeBase = dd.ToString();
            //string ppt = exeBase + @"\" + @"d5000.pdf";
            string ppt = exeBase + @"\" + @"ppt1.pdf";
            //Console.WriteLine(ppt);

            Console.WriteLine("writing PDF: page 4");
            Console.WriteLine(pdf2TextByPage(ppt, 4));
            Console.WriteLine("----------------------");
            //pdf2Image(ppt);
            //Console.ReadLine();
            //return;

            //Console.WriteLine(args[0]); 

            string prefix = "";
            string pptfile = "";

            /*  GET PATH OF EXE  */
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            //To get the location the assembly normally resides on disk or the install directory
            //string path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            //once you have the path you get the directory with:
            var directory = System.IO.Path.GetDirectoryName(pa);
            exeBase = dd.ToString();
            /* END PATH */
            imageBase = exeBase;

            /* 
             * For commandline only run: 
             * $> PPT2Image sample.pptx
             * OR
             * 
            */
            /*
            if (args.Length != 0)  // while using commandline <this>.exe <filename>.pptx
            {                
                ppt = exeBase + @"\" + args[0].ToString().Trim();
                prefix = args[0].ToString().Trim();
            }

            else // while running tests
            {
                //String pptfile = @"G:\LS_CAT_Apps_Weekly_7-17-15_Final.pptx";
                //string pptfile = @"H:\output\FULL_Q1FY17_LS-SWIFT_DivisionReview_Draft_Final.pptx";
                ppt = exeBase + @"\" + @"ppt1.pptx";
                
            }
            */
            pptfile = exeBase + @"\" + @"ppt1.pptx";

            // remove when standalone application 
            //string[] filePaths = System.IO.Directory.GetFiles(imagebase + @"\", "*.pptx");
            //pptfile = filePaths[0];
            // END 



            //db = new DBConnect();
            //db.Insert("insert into weekly (filename, hashtext, imgthumb, imglarge) values ('LS_WEEKLY5', 'PO1 ET', '/img/weekly7/thumb.png', '/img/weekly7/HD.png') "); 



            Console.WriteLine("Exe Base Dir = " + exeBase);
            Console.WriteLine("Image Base Dir = " + imageBase);
            Console.WriteLine("File = " + pptfile);
            ppt2Image(pptfile, imageBase, prefix);


            readPPTText(pptfile);

            //mergePPTs(exeBase + @"\" + @"ppt1.pptx", exeBase + @"\" + @"ppt2.pptx"); 

        }


        public void pdf2Image(string pdfFile)
        //static void pdf2Image(string pptfile, string prefix)
        {

            // Do not forget the %d in the output file name  @"Example%d.jpg"
            GhostscriptWrapper.GeneratePageThumbs(exeBase + @"\" + @"d5000.pdf", exeBase + @"\" + @"Example%d.jpg", 1, 15, 100, 100);

            // for a single page [you have to know the page number -- no function in ghostscript]
            // GhostscriptWrapper.GeneratePageThumb(exeBase + @"\" + @"d5000.pdf", exeBase + @"\" + @"Example1.jpg", 1, 100, 100);            
        }

        public void pdf2ImageByPage(string pdfFile, int pagenumber)
        //static void pdf2Image(string pptfile, string prefix)
        {
            // Do not forget the %d in the output file name  @"Example%d.jpg"
            GhostscriptWrapper.GeneratePageThumbs(exeBase + @"\" + @"d5000.pdf", exeBase + @"\" + @"Example%d.jpg", 1, 15, 100, 100);
            // for a single page [you have to know the page number -- no function in ghostscript]
            // GhostscriptWrapper.GeneratePageThumb(exeBase + @"\" + @"d5000.pdf", exeBase + @"\" + @"Example1.jpg", 1, 100, 100);            
        }


        /**************************           FINAL        *****************************/
        public string pdf2Text(string pdfFile)
        {
            PdfReader reader = new PdfReader(pdfFile);
            //ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();
            ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();
            
            string txt = "";
            int pages = reader.NumberOfPages;

            for (int page = 1; page <= reader.NumberOfPages; page++)
            {

                txt = PdfTextExtractor.GetTextFromPage(reader, page, its);
                txt = txt.Replace("\n", " ");
                // Replace multi spaces with single 
                RegexOptions options = RegexOptions.None;
                Regex regex = new Regex("[ ]{2,}", options);

                txt = regex.Replace(txt, " ");
                Console.WriteLine("Slide #" + page + "\n-----------------------\n" + txt + "\n-----------------------\n");
            }
            // txt = txt.Replace("\n", " ").Replace("  ", " "); ;


            return txt;
        }

        /**************************           FINAL        *****************************/
        public string pdf2TextByPage(string pdfFile, int pagenumber)
        {
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
        public void pdfSplit (string pdffile, string prefix)
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
            pdfpath = pdfpath +@"\"+ pdfname+ ".pdf";

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
                /* full HD*/
                pptPresentation.Slides[i].Export(exportPath + @"\" + "slide" + i + ".png", "png", pageWidth, pageHeight);

                /* Thumbnail*/
                pptPresentation.Slides[i].Export(exportPath + @"\" + "thumb.slide" + i + ".png", "png", pageWidth / 2, pageHeight / 2);

            }

            ppt2ContentImages(pptPresentation, exportPath.Replace(@"\ppt-img", @"\" + "content-img"));

            pptPresentation.Close();
        }

        void ppt2ContentImages( Presentation pptPresentation, string exportPath)
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
                                text = Regex.Replace(text, @"[^\t\r\n\u0020-\u007E]+", " ");
                                text = Regex.Replace(text, @"[ ]{ 2,}", string.Empty);
                                if (text.Length > 2)
                                {
                                    slide_title = slide_title.Replace("NOTITLE", "").Replace("\v"," ").Replace("\f", " ") + text + "\n";
                                    slide_title = Regex.Replace(slide_title, @"[^\t\r\n\u0020-\u007E]+", string.Empty);
                                }

                            }
                        }
                    }
                } catch (Exception ex)
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

                System.IO.File.AppendAllText(@".\OutText.md", "---\nSlide #" + slide.SlideNumber +"\n# " + slide_title );
                System.IO.File.AppendAllText(@".\OutText.md", "\n" + pps + "\n");
            }
            pptPresentation.Close();
        }

        private string readShape(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            string str = "";
            // Just Checking this ! //comment it (below) out 
            //if (shape.Type == MsoShapeType.msoEmbeddedOLEObject)
            //{
            //    Console.WriteLine("Excel Table = true");
            //} else
            //{
            //    Console.WriteLine("Excel Table = false");
            //}

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

                //Microsoft.Office.Interop.PowerPoint.SeriesCollection chartSeriesA = (Microsoft.Office.Interop.PowerPoint.SeriesCollection)t.SeriesCollection();
                //foreach (Microsoft.Office.Interop.PowerPoint.Series Srs in chartSeriesA)
                //{
                //    System.Array a = (System.Array)((object)Srs.Values); 
                //    //var XV = Srs.XValues;
                //    //var V = Srs.Values;
                //    //str += Srs.ToString();
                //}

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

            //str = CleanInput(str.Replace("\n\n", "\n"));
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
            pptPresentation1.SaveAs(exeBase + @"\" + @"temp.pptx", PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            pptPresentation1.Close();

            PowerPoint.Presentation mergedPPT = app.Presentations.Open(exeBase + @"\" + @"temp.pptx", MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);

            int slide_count = 5; // pptPresentation1.Slides.Count;
            Console.Write("count=" + slide_count);

            //mergedPPT.Slides.InsertFromFile(pptfile1, 1, -1);

            mergedPPT.Slides.InsertFromFile(pptfile2, 0, 1, -1);

            mergedPPT.SaveAs(exeBase + @"\" + @"merged.pptx", PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
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
   

        /* 
         * http://mantascode.com/c-get-text-content-from-microsoft-powerpoint-file/ 
         */
        // under construction - still buggy on: shape.Chart to string 
        public void readPPTText(string pptfile)
        {
            Microsoft.Office.Interop.PowerPoint.Application PowerPoint_App = new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;
            Microsoft.Office.Interop.PowerPoint.Presentation presentation = multi_presentations.Open(pptfile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            /*  MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse 
             required to not open the file in a separate process.
             */
            string presentation_text = "";
            string fulltext = "";
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                //fulltext += "\n\n";
                //fulltext += "Slide:" + (i + 1) + " || ";
                Console.WriteLine("______Slide # " + (i + 1) + "________");

                foreach (var item in presentation.Slides[i + 1].Shapes)
                {
                    var shape = (Microsoft.Office.Interop.PowerPoint.Shape)item;

                    /* if shape object is a group of shapes
                     http://www.pptfaq.com/FAQ00600_Changing_shapes_within_groups_-without_ungrouping-.htm                     
                     */
                    if (shape.Type == MsoShapeType.msoGroup)
                    {
                        Console.WriteLine(shape.GroupItems.Count);
                        for (int j = 1; j <= shape.GroupItems.Count; ++j)
                        {
                            string temp = getShapeText((Microsoft.Office.Interop.PowerPoint.Shape)shape.GroupItems[j]);
                            fulltext += temp;
                            presentation_text += temp;
                        }
                    }
                    /* else if shape object is NOT group of shapes and just a Shape*/
                    else
                    {
                        string temp = getShapeText(shape);
                        fulltext += temp;
                        presentation_text += temp;
                    }

                }
                presentation_text = presentation_text.Replace("\"", " ").Replace("\'", " ").Replace("\n", " ").Replace("\r", " ");
                //presentation_text= presentation_text.Replace("\'", " ");
                //Console.WriteLine("insert into weekly (filename, hashtext, imgthumb, imglarge) values ('" + pptfile + "', '" + presentation_text + "', '/img/weekly/" + pptfile + "thumb_slide" + (i + 1) + ".png', '/img/weekly/" + pptfile + "slide" + (i + 1) + ".png') ");
                //db.Insert("insert into weekly (filename, hashtext, imgthumb, imglarge) values ('" + pptfile + "', '" + presentation_text + "', '/img/weekly/" + pptfile + "_slide" + (i + 1) + ".png', '/img/weekly/" + pptfile + "_slide" + (i + 1) + ".png') "); 


                Console.Write("Slide:" + (i + 1) + ":\n" + presentation_text);
                Console.Write("\n ");

                Console.ReadKey();

                presentation_text = "";


            }
            //System.IO.File.WriteAllText("G:\\text.txt",fulltext);
            System.IO.File.WriteAllText("text.txt", fulltext);            
            PowerPoint_App.Quit();
            Console.WriteLine(presentation_text);
            Console.ReadLine();
        }

        public string getShapeText(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            string presentation_text = "";
            string textString = "";


            if (shape.HasTextFrame == MsoTriState.msoTrue)
            {
                if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                {

                    var textFrame = shape.TextFrame;
                    var textRange = textFrame.TextRange;
                    var paragraphs = textRange.Paragraphs(-1, -1);
                    foreach (Microsoft.Office.Interop.PowerPoint.TextRange paragraph in paragraphs)
                    {
                        var text = paragraph.Text;
                        text = text.Replace("\r", "");
                        text = text.Replace("\n", " ");
                        presentation_text += text + " ";
                        textString += text + " ";
                    }
                }
            }

            /*
            if (shape.HasTable == MsoTriState.msoTrue)
            {
                var t = shape.Table;

                for (int j = 1; j <= t.Rows.Count; ++j)
                    for (int k = 1; k <= t.Columns.Count; ++k)
                    {
                        //if (shape.HasTextFrame == MsoTriState.msoTrue)
                        //{
                        //    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                        //    {

                        var textFrame = t.Cell(j, k).Shape.TextFrame;
                        var textRange = textFrame.TextRange;

                        presentation_text += textRange.Text + " ";
                        textString += textRange.Text + " ";

                        //    }
                        //}
                    }
            }
            */

            //if (shape.HasChart == MsoTriState.msoTrue)
            //{
            //    Console.WriteLine("Has Chart: True");
            //    Microsoft.Office.Interop.PowerPoint.Chart t = shape.Chart;


            //    if (t.HasTitle)
            //    {
            //        Console.WriteLine("Has DataTable: True");
            //        Console.WriteLine("Title:" + t.ChartTitle.Text.ToString()); textString += t.ChartTitle.Text.ToString() + " ";
            //    }

            //    if (t.HasDataTable)
            //    {
            //        Console.WriteLine("Has DataTable: True");
            //        var p = t.DataTable;
            //        //textString += t.DataTable.ToString();
            //    }

            //    textString += shape.Chart.ToString();

            //    Console.WriteLine("Shape.Chart.tostring()" + shape.Chart.ToString());

            //    Microsoft.Office.Interop.PowerPoint.SeriesCollection tmp = (Microsoft.Office.Interop.PowerPoint.SeriesCollection)t.SeriesCollection();
            //    Console.WriteLine("Series Count:" + tmp.Count);



            //    /*

            //    for (int j = 1; j <= tmp.Count; ++j)
            //    {
            //        Microsoft.Office.Interop.PowerPoint.Series aSeries = (Microsoft.Office.Interop.PowerPoint.Series)tmp.Item(j);

            //        foreach (object v in (Array)aSeries.XValues)
            //        {
            //            if (v != null) { Console.WriteLine(v.ToString()); textString += v.ToString() +" "; }
            //        }
            //        foreach (object v in (Array)aSeries.Values)
            //        {
            //            if (v != null) { Console.WriteLine(v.ToString()); textString += v.ToString()+ " "; }
            //        }
            //    }               
            //    */
            //    /*
            //    foreach (Microsoft.Office.Interop.PowerPoint.Series aSeries in tmp ) {
            //        foreach (object v in aSeries.XValues)
            //        {

            //        }
            //        foreach (object v in aSeries.Values as Array)
            //        {
                        
            //        }
            //        var p = aSeries.XValues;

            //        ;
            //        try
            //        {
            //            Console.WriteLine(ArrayToStringGeneric(p, " "));
            //        }
            //        catch (Exception e) { Console.WriteLine(e.ToString()); }
            //    }

            //    */
            //}
            textString = textString.Replace("\r", "");
            textString = textString.Replace("\n", " ");
            return textString;

        }


        public string ArrayToStringGeneric<T>(IList<T> array, string delimeter)
        {
            string outputString = "";

            for (int i = 0; i < array.Count; i++)
            {
                if (array[i] is IList<T>)
                {
                    //Recursively convert nested arrays to string
                    outputString += ArrayToStringGeneric<T>((IList<T>)array[i], delimeter);
                }
                else
                {
                    outputString += array[i];
                }

                if (i != array.Count - 1)
                    outputString += delimeter;
            }

            return outputString;
        }


    }
}




// failed junk with series.Values and series.XValues for non-embedded excel sheets in charts. 
/*
                    //object [] V = (object[])aSeries.Values;
                    //string[] P = (string[])aSeries.DataLabels();
                    
                    //var XV = aSeries.XValues; 

                    //Microsoft.Office.Interop.PowerPoint.Points pts = tmp.Item(j).Points();
                    //Console.WriteLine("Points #:" + pts.Count);
                    //Console.WriteLine(pts.Item(2).Name);

                    //Does not work : Invalid Case Ex. 
                    //foreach (PowerPoint.Point p in pts) {
                    //    Console.WriteLine(p.Name);
                    //}


                    // Working till this point 
                    // now lets extract (XValues , Values) -- How ? I dont know. :P 


                    //Console.Write(aSeries.Points());
                    //foreach (PowerPoint.Point p in aSeries.Points() as PowerPoint.Points)
                    //{
                    //    Console.WriteLine((Array)p);
                    //}

                    ////var p = aSeries.XValues();
                    //dynamic p = aSeries
                    //try
                    //{
                    //    Console.WriteLine(ArrayToStringGeneric(p, " "));
                    //}
                    //catch (Exception e) { Console.WriteLine(e.ToString()); } 



 */
