using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Driver;
using MongoDB.Bson;

namespace fileHasherConverter
{
    class MongoDBConnect
    {
        private static string connectionString;
        private static MongoClient client;

        public void Connect()
        {
            connectionString = "mongodb://localhost:27017";
            try
            {
                client = new MongoClient(connectionString);
                Console.WriteLine("Connected");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine("MongoDB Connection Faile. " + e.StackTrace);
            }           
        }

        public async void getDoc()
        {
            await getDocuments();
        }

        public async Task getFilteredDocuments()
        {
            connectionString = "mongodb://localhost:27017";
            client = new MongoClient(connectionString);
            IMongoDatabase db = client.GetDatabase("test_database");

            var collection = db.GetCollection<BsonDocument>("posts");

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

        public async Task getDocuments()
        {
            connectionString = "mongodb://localhost:27017";
            client = new MongoClient(connectionString);
            IMongoDatabase db = client.GetDatabase("test_database"); 

            var collection = db.GetCollection<BsonDocument>("posts");
            
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
        
        internal class posts
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string Class { get; set; }
            public int Age { get; set; }
            public IEnumerable<string> Subjects { get; set; }
        }

    }
}
