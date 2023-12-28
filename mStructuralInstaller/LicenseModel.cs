using MongoDB.Bson;
using System;

namespace mStructuralInstaller
{
    public class LicenseModel
    {
        public ObjectId Id { get; set; }
        public string SerialNo { get; set; }
        public DateTime ExpiryDate { get; set; }
    }
}
