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
using System.Text.RegularExpressions;

namespace fileHasherConverter
{
    public class FileHash
    {
        //public string imageBase = @"H:\output";
        public string imageBase = ".";
        public string exeBase = ".";
        //public DBConnect db;


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

        public void pdf2Image(string pdfFile, int pagenumber)
        //static void pdf2Image(string pptfile, string prefix)
        {

            // Do not forget the %d in the output file name  @"Example%d.jpg"
            GhostscriptWrapper.GeneratePageThumbs(exeBase + @"\" + @"d5000.pdf", exeBase + @"\" + @"Example%d.jpg", 1, 15, 100, 100);

            // for a single page [you have to know the page number -- no function in ghostscript]
            // GhostscriptWrapper.GeneratePageThumb(exeBase + @"\" + @"d5000.pdf", exeBase + @"\" + @"Example1.jpg", 1, 100, 100);            
        }

        public string pdf2TextByPage(string pdfFile, int pagenumber)
        {
            PdfReader reader = new PdfReader(pdfFile);
            ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();

            string txt = PdfTextExtractor.GetTextFromPage(reader, pagenumber, new SimpleTextExtractionStrategy());
            txt = txt.Replace("\n", " ");
            // Replace multi spaces with single 
            RegexOptions options = RegexOptions.None;
            Regex regex = new Regex("[ ]{2,}", options);

            txt = regex.Replace(txt, " ");

            // txt = txt.Replace("\n", " ").Replace("  ", " "); ;


            return txt;
        }

        public void pdfMerge(string pptfile, string prefix)
        {

        }



        /*
         * http://www.free-power-point-templates.com/articles/c-code-to-convert-powerpoint-to-image/         
         */

        public void ppt2Image(string pptfile, string path, string prefix)
        {
            Console.WriteLine("PPT File Location:" + pptfile);
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptPresentation = pptApplication.Presentations
            .Open(pptfile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);


            //Save as PDF for text Extraction
            //string pdfpath = pptfile.Replace(".ppt", "") + ".pdf";
            //pptPresentation.ExportAsFixedFormat(pdfpath, PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint, MsoTriState.msoFalse);


            int slide_count = pptPresentation.Slides.Count;
            Console.Write("count=" + slide_count);

            for (int i = 1; i <= slide_count; ++i)
            {
                /* full HD*/
                pptPresentation.Slides[i].Export(imageBase + @"\" + path + "slide" + i + ".png", "png", 800, 600);

                /* Thumbnail*/
                pptPresentation.Slides[i].Export(imageBase + @"\" + path + "thumb.slide" + i + ".png", "png", 320, 240);

            }

        }



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


            if (shape.HasChart == MsoTriState.msoTrue)
            {
                Console.WriteLine("Has Chart: True");
                Microsoft.Office.Interop.PowerPoint.Chart t = shape.Chart;


                if (t.HasTitle)
                {
                    Console.WriteLine("Has DataTable: True");
                    Console.WriteLine("Title:" + t.ChartTitle.Text.ToString()); textString += t.ChartTitle.Text.ToString() + " ";
                }

                if (t.HasDataTable)
                {
                    Console.WriteLine("Has DataTable: True");
                    var p = t.DataTable;
                    //textString += t.DataTable.ToString();
                }

                textString += shape.Chart.ToString();

                Console.WriteLine("Shape.Chart.tostring()" + shape.Chart.ToString());

                Microsoft.Office.Interop.PowerPoint.SeriesCollection tmp = (Microsoft.Office.Interop.PowerPoint.SeriesCollection)t.SeriesCollection();
                Console.WriteLine("Series Count:" + tmp.Count);



                /*

                for (int j = 1; j <= tmp.Count; ++j)
                {
                    Microsoft.Office.Interop.PowerPoint.Series aSeries = (Microsoft.Office.Interop.PowerPoint.Series)tmp.Item(j);

                    foreach (object v in (Array)aSeries.XValues)
                    {
                        if (v != null) { Console.WriteLine(v.ToString()); textString += v.ToString() +" "; }
                    }
                    foreach (object v in (Array)aSeries.Values)
                    {
                        if (v != null) { Console.WriteLine(v.ToString()); textString += v.ToString()+ " "; }
                    }
                }               
                */
                /*
                foreach (Microsoft.Office.Interop.PowerPoint.Series aSeries in tmp ) {
                    foreach (object v in aSeries.XValues)
                    {

                    }
                    foreach (object v in aSeries.Values as Array)
                    {
                        
                    }
                    var p = aSeries.XValues;

                    ;
                    try
                    {
                        Console.WriteLine(ArrayToStringGeneric(p, " "));
                    }
                    catch (Exception e) { Console.WriteLine(e.ToString()); }
                }

                */
            }
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
