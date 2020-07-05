using System;
using MongoDB.Driver;
using MongoDB.Bson;
using System.Threading.Tasks;
using System.Collections.Generic;
//using Microsoft.Office.Interop.Excel;

namespace fileHasherConverter
{
    class Program
    {
        //static string imageBase = @"H:\output";
        static string ppt_img_root = ".";
        static string exe_root = ".";
        static string content_img_root = ".";
        static string pdf_root = ".";

        /* Database */
        private static string connectionString;
        private static MongoClient client;
        public static bool status = false;
        private static IMongoDatabase db;
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

            //runHash();
            test_PPT_Parser();

            /*
             *  READ FILESTOR for new unprocessed files: isProcessed = false; 
             */

            // connect fileStor DB 

            // read collection 

            // for entries in collection (files) loop 

            // if type is ppt. pptx - switch function 

            // call file type --> Controller --> processing file data  

            // update fileStor DB as         'proocessed'


            Console.WriteLine("All Tasks Completed.");
            Console.ReadKey();
        }


        static void test_PPT_Parser()
        {

            /* read file store database */
            /* connect database */
            //md = new MongoDBConnect();
            //md.Connect("result_database");
            //md.getDocuments("fileStor");

            //var filter = "{ filename : 'Nanopoint_Training_Agenda.pptx'}"; 
            //var filter = "{ isProcessed : false, filetype : { $nin: ['.pdf', '.doc', '.docx'] } }";
            var filter = "{ isProcessed : false, filetype : '.pptx'}";

            getFilteredDocuments("fileStor", filter);
            
            /*

            pptController pt = new pptController();

            // sample path example 
            string pa = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var dd = System.IO.Path.GetDirectoryName(pa);            
            string exe_root = dd.ToString();
            string ppt_img_root = exe_root;
            string pptfile = exe_root + @"\" + @"LS-Char-Extraction.pptx";

            pt.runPptParserFlow("Hello"); 

            */

        }


        public static async Task getFilteredDocuments(string collection_name, FilterDefinition<BsonDocument> filter)
        {
            connectionString = "mongodb://localhost:27017";
            client = new MongoClient(connectionString);
            IMongoDatabase db = client.GetDatabase("result_database");

            var collection = db.GetCollection<BsonDocument>(collection_name);

            //FilterDefinition<BsonDocument> filter = FilterDefinition<BsonDocument>.Empty;
            FindOptions<BsonDocument> options = new FindOptions<BsonDocument>
            {
                BatchSize = 2,
                NoCursorTimeout = false
            };
            
            
            using (IAsyncCursor<BsonDocument> cursor = await collection.FindAsync(filter, options))
            {                
                var batch = 0;
                var countRecord = 0; 
                while (await cursor.MoveNextAsync() /*&& countRecord <2*/)
                {
                    IEnumerable<BsonDocument> documents = cursor.Current;

                    batch++;

                    Console.WriteLine($"\nBatch: {batch}");
                    

                    foreach (BsonDocument document in documents)
                    {
                        //Console.WriteLine("Processing document: "+ document["_id"] + "  Path:"+ document["path"]);
                        //Console.WriteLine();


                        // sample path example 
                        //string pa = System.Reflection.Assembly.GetExecutingAssembly().Location;
                        //var dd = System.IO.Path.GetDirectoryName(pa);
                        //string exe_root = dd.ToString();
                        //string ppt_img_root = exe_root;
                        //string pptfile = exe_root + @"\" + @"LS-Char-Extraction.pptx";

                        //pt.runPptParserFlow(document["rawPath"].ToString(), document.AsObjectId);
                        //pt.runPptParserFlow(document["rawPath"].ToString(), document["_id"].ToString());

                        //Console.WriteLine("Starting to run ppt parser flow..");


                        // check filetype 
                        string filetype = document["filetype"].ToString();
                        Console.WriteLine("FileType " + filetype);
                        var motherID = ObjectId.Empty;
                        bool success = false;

                        // TODO task if '*.PPT , *.PPTX                    
                        if (filetype.EndsWith(".ppt") || filetype.EndsWith(".pptx"))
                        {
                            pptController pt = new pptController();

                            var ret = pt.runPptParserFlow(document);
                            motherID = ret.Item1;
                            success = ret.Item2;
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("[ERROR] File Type " + filetype + " not recognized");
                            success = false;
                            Console.ForegroundColor = ConsoleColor.White;

                        }

                        // TODO For PDF 

                        // TODO for DOC, DOCX


                        //Console.WriteLine("Mother ID extracted.");
                        /* Update FileStor {isProcessed : true } */
                        // create new Class entity for post 

                        //var fileCollection = db.GetCollection<BsonDocument>("fileStor");

                        //find the MasterID with 1130 and replace it with 1120
                        if (success)
                        {
                            var result = await collection.FindOneAndUpdateAsync(
                                            Builders<BsonDocument>.Filter.Eq("_id", document["_id"]),
                                            Builders<BsonDocument>.Update.Set("isProcessed", false)
                                            .Set("MotherID", motherID));
                        }
                        else
                        {
                            var result = await collection.FindOneAndUpdateAsync(
                                            Builders<BsonDocument>.Filter.Eq("_id", document["_id"]),
                                            Builders<BsonDocument>.Update.Set("isProcessed", false)
                                            .Set("MotherID", motherID)
                                            .Set("logtext", "ERROR"));
                        }

                            //retrive the data from collection
                            //await fileCollection.Find(new BsonDocument())
                            // .ForEachAsync(x => Console.WriteLine(x));

                            countRecord++;
                    }
                }

                Console.WriteLine($"Total Batch: { batch}");
                Console.WriteLine($"Total Records Processed: { countRecord}");
                Console.ReadLine();

            }

        }


        public static async Task getDocuments(string collection_name)
        {
            connectionString = "mongodb://localhost:27017";
            client = new MongoClient(connectionString);
            IMongoDatabase db = client.GetDatabase("result_database");

            var collection = db.GetCollection<BsonDocument>(collection_name);

            using (IAsyncCursor<BsonDocument> cursor = await collection.FindAsync(new BsonDocument()))
            {
                while (await cursor.MoveNextAsync())
                {
                    IEnumerable<BsonDocument> batch = cursor.Current;
                    foreach (BsonDocument document in batch)
                    {
                        Console.WriteLine(document);
                        Console.WriteLine();
                    }
                }
            }
        }



        static void runHash()
        {
            __BACKUP__FileHash newHash = new __BACKUP__FileHash();

            string pa = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var dd = System.IO.Path.GetDirectoryName(pa);
            exe_root = dd.ToString();
            string ppt = exe_root + @"\" + @"d5000.pdf";
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
            exe_root = dd.ToString();
            /* END PATH */
            ppt_img_root = exe_root;

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
            //pptfile = exe_root + @"\" + @"ppt1.pptx";
            //pptfile = exe_root + @"\" + @"LS-SWIFT.pptx";
            //pptfile = exe_root + @"\" + @"VoyagerWeeklyUpdate_18-07-27.pptx";



            pptfile = exe_root + @"\" + @"LS-Char-Extraction.pptx";
            //pptfile = exe_root + @"\" + @"Flash_SJ_Pitch.ppt";
            //pptfile = exe_root + @"\" + @"9650_Intro_111110.ppt";


            //pptfile = exeBase + @"\" + @"LS-SWIFT_Product_Apps_Weekly_2018-07-06.pptx";
            // remove when standalone application 
            //string[] filePaths = System.IO.Directory.GetFiles(imagebase + @"\", "*.pptx");
            //pptfile = filePaths[0];
            // END 



            //db = new DBConnect();
            //db.Insert("insert into weekly (filename, hashtext, imgthumb, imglarge) values ('LS_WEEKLY5', 'PO1 ET', '/img/weekly7/thumb.png', '/img/weekly7/HD.png') "); 
            //MongoDBConnect mg = new MongoDBConnect();
            //mg.Connect();
            //mg.getDocuments();
            

            ////print all ascii chars

            //for (int i = 0; i < 256; ++i)
            //{
            //    Console.Write(i + ":" + (char)i + "| ");
            //}

            Console.WriteLine("Exe Base Dir = " + exe_root);
            Console.WriteLine("Image Base Dir = " + ppt_img_root);
            ppt_img_root = exe_root + @"\ppt-img";
            pdf_root = exe_root + @"\pdf";
            content_img_root = exe_root + @"\content-img";

            /* PPT Operations */
            pptParser fe = new pptParser();
            fe.GetMetaData();
            //fe.ppt2Image(pptfile, ppt_img_root, prefix);
            //fe.ppt2ImageAlt(pptfile, ppt_img_root, prefix);

            //fe.ppt2pdf(pptfile, "8ih423yu673", pdf_root, "");
            fe.ppt2text(pptfile);
            //fe.ppt2pdfByPage(pptfile, null, null, 1);

            /* PDF Operations */
            //string pdfFile = pdf_root + @"\8ih423yu673.pdf";
            //fe.pdf2Text(pdfFile);
            //fe.pdf2TextByPage(pdfFile, 2);

            //newHash.readPPTText(pptfile);
            //newHash.mergePPTs(exeBase + @"\" + @"ppt1.pptx", exeBase + @"\" + @"ppt2.pptx");

            pptController pc = new pptController();


        }

    }

}