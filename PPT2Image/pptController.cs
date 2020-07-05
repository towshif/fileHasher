using System;
using System.Collections.Generic;
using MongoDB.Bson;

namespace fileHasherConverter
{
    public class pptController
    {

        /* Database */
        MongoDBConnect md;
        
        public Tuple <ObjectId, bool> runPptParserFlow(BsonDocument document)
        {
            string pptfile = document["rawPath"].ToString();
            string fid = document["_id"].ToString(); 
            string fname = document["filename"].ToString();



            //Console.WriteLine(pptfile + " : id = " + fid);
            ObjectId file_ID = new ObjectId(fid); 

            /* connect database */
            md = new MongoDBConnect();
            md.Connect("result_database");

            //read raw ppt Path from file ID 

            // insert entry ppt : 'slide collection' into Mother 
            var mother_ID = md.newEmptyRecord("mother");

            // get iD from mother for entry 
            

            /*  TESTING CODE */
            string pa = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var dd = System.IO.Path.GetDirectoryName(pa);
            string binary_root = dd.ToString();
            string ppt_img_root = binary_root + @"\ppt-img";
            string content_img_root = binary_root + @"\content-img";
            //string pptfile = binary_root + @"\" + @"LS-Char-Extraction.pptx";

            /*  CLOSE INITIALIZING CODE */

            pptParser my_Ppt_Parser = new pptParser(mother_ID, file_ID, pptfile, ppt_img_root, content_img_root, document);


            bool success = my_Ppt_Parser.runPptParserFlow();

            if (success)
            {
                /* Update MOTHER DB with <collection> {slide ID, meta , author } */
                // create new Class entity for post 
                mongoData mg = new mongoData();
                mg.Id = mother_ID;
                mg.Data = new BsonDocument().AddRange(my_Ppt_Parser.get_PPT_METADATA());
                mg.Collection = new BsonDocument().AddRange(my_Ppt_Parser.get_Slide_Collection());


                /*      Addition Fields for Mother      */
                Dictionary<string, Object> fd = new Dictionary<string, Object>();
                //List<Object> lt = new List<Object>();

                object lt = document["tags"];

                // link parent 
                fd.Add("sourceID", file_ID);
                fd.Add("source", "fileStor");
                fd.Add("filename", fname);
                fd.Add("filePath", pptfile);
                fd.Add("fileType", "*.pptx");
                // link child <collection>
                fd.Add("collection_source", "dataSlides");
                fd.Add("tags", lt);
                mg.Data.AddRange(fd);


                md.updateEntireRecord("mother", mother_ID, mg);

                return Tuple.Create(mother_ID, true);
            }

            else
                return Tuple.Create( ObjectId.Empty , false) ;

        }

    }
}
