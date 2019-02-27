using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Driver;
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace fileHasherConverter
{
    public class mongoData
    {
        public ObjectId Id { get; set; }

        [BsonIgnoreIfNull]
        public BsonDocument Collection { get; set; }

        [BsonExtraElements]
        public BsonDocument Data { get; set; }


    }
}
