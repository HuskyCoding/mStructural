using MongoDB.Bson;
using System;

namespace mStructuralInstaller
{
    public class ActivationModel
    {
        public ObjectId Id { get; set; }
        public string PcName { get; set; }
        public string MacAddress { get; set; }
        public string License { get; set; }
        public bool Status { get; set; }
        public DateTime TimeStamp { get; set; }
    }
}
