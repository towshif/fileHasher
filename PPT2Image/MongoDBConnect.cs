using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Driver;
using MongoDB.Bson;

namespace fileHasherConverter
{
    public class Entity
    {
        public ObjectId Id { get; set; }
        public bool isProcessed { get; set; }
    }
    class MongoDBConnect
    {
        private string connectionString;
        private MongoClient client;
        public bool status = false;
        private IMongoDatabase db;

        public void Connect(string databaseName)
        {
            connectionString = "mongodb://localhost:27017";
            try
            {
                client = new MongoClient(connectionString);
                //Console.WriteLine("Connected");
                status = true;
                //Console.ReadKey();
                db = client.GetDatabase(databaseName);
            }
            catch (Exception e)
            {
                status = false;
                Console.WriteLine("MongoDB Connection Failed. " + e.StackTrace);
            }
        }

        public ObjectId newEmptyRecord(string collectionName)
        {
            var collection = db.GetCollection<mongoData>(collectionName);

            var doc = new mongoData { }; /* empty document*/
            collection.InsertOne(doc);
            var id = doc.Id;
            return id; 
        }


        // /to be updated 
        public ObjectId postRecord(string collectionName)
        {
            var collection = db.GetCollection<mongoData>(collectionName);

            var doc = new mongoData { }; /* empty document*/
            collection.InsertOne(doc);
            var id = doc.Id;
            return id;
        }


        public void updateEntireRecord (string collectionName, ObjectId ID, mongoData dt )
        {
            var collection = db.GetCollection<BsonDocument>(collectionName);
            //var query = Query<mongoData>.EQ(p => p.Id, 10);

            var filter = Builders<BsonDocument>.Filter.Eq("_id", ID);
            //var update = Builders<BsonDocument>.Update.

            var replaceOneResult = collection.ReplaceOne(filter, dt.ToBsonDocument());

            //Console.WriteLine("replacement result: " + replaceOneResult); 

        }

        public ObjectId postError(string collectionName, mongoData errorData)
        {
            Console.WriteLine(errorData);
            var collection = db.GetCollection<mongoData>(collectionName);
            //var doc = new mongoData { }; /* empty document*/
            collection.InsertOne(errorData);
            var id = errorData.Id;
            return id;

        }

        public async void getDoc()
        {
            //await getDocuments();
        }

        public async Task getFilteredDocuments( string collection_name)
        {

            var collection = db.GetCollection<BsonDocument>(collection_name);

            FilterDefinition<BsonDocument> filter = FilterDefinition<BsonDocument>.Empty;
            FindOptions<BsonDocument> options = new FindOptions<BsonDocument>
            {
                BatchSize = 2,
                NoCursorTimeout = false
            };

            using (IAsyncCursor<BsonDocument> cursor = await collection.FindAsync(filter, options))
            {

                var batch = 0;
                while (await cursor.MoveNextAsync())
                {
                    IEnumerable<BsonDocument> documents = cursor.Current;
                    batch++;

                    Console.WriteLine($"Batch: {batch}");

                    foreach (BsonDocument document in documents)
                    {
                        Console.WriteLine(document);
                        Console.WriteLine();
                    }
                }

                Console.WriteLine($"Total Batch: { batch}");
            }

        }

        public async Task getDocuments(string collection_name)
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

        
    }
}
