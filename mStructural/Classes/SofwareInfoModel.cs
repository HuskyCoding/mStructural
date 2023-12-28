using MongoDB.Bson;

namespace mStructural.Classes
{
    public class SofwareInfoModel
    {
        public ObjectId _id  { get; set; }
        public string version { get; set; }
    }
}
