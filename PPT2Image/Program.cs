﻿using System;
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

//using Microsoft.Office.Interop.Excel;

namespace fileHasherConverter
{
    class Program
    {
        //static string imageBase = @"H:\output";
        static string imageBase = ".";
        static string exeBase = ".";
        //static DBConnect db;


        static void Main(string[] args)
        {
            //MongoDBConnect newConn = new MongoDBConnect();
            //newConn.Connect();
            ////newConn.getDoc(); 
            //newConn.getDocuments().Wait();
            //newConn.getFilteredDocuments().Wait();
            //try
            //{
            //    newConn.testConnection().Wait();
            //}
            //catch (Exception e) { Console.Write(e.StackTrace); }

            runHash();

            Console.WriteLine("All Tasks Completed.");
            Console.ReadKey();
        }

        static void runHash()
        {
            FileHash newHash = new FileHash();

            string pa = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var dd = System.IO.Path.GetDirectoryName(pa);
            exeBase = dd.ToString();
            string ppt = exeBase + @"\" + @"d5000.pdf";
            //Console.WriteLine(ppt);
            //Console.WriteLine(pdf2Text(ppt));
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

            newHash.ppt2Image(pptfile, imageBase, prefix);

            //newHash.readPPTText(pptfile);

            //newHash.mergePPTs(exeBase + @"\" + @"ppt1.pptx", exeBase + @"\" + @"ppt2.pptx");
        }

    }

}