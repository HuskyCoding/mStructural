using MongoDB.Bson;
using MongoDB.Driver;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Windows.Forms;

namespace mStructural.Function
{
    public class Update
    {
        #region Private Variables
        #endregion

        // Constructor
        public Update()
        {
        }

        // Main Method
        public async void Run()
        {
            // mongodb connection string
            const string connectionUri = "mongodb+srv://maxwell:Mwellbend92@mstructural.zdeasje.mongodb.net/?retryWrites=true&w=majority";
            var settings = MongoClientSettings.FromConnectionString(connectionUri);
            // Set the ServerApi field of the settings object to Stable API version 1
            settings.ServerApi = new ServerApi(ServerApiVersion.V1);
            // Create a new client and connect to the server
            var client = new MongoClient(settings);
            // get assembly version
            Version version = Assembly.GetExecutingAssembly().GetName().Version;

            // Get version info
            var db = client.GetDatabase("mStructural");
            var collection = db.GetCollection<Classes.SofwareInfoModel>("SoftwareInfo");
            ObjectId objectId;
            bool bRet = ObjectId.TryParse("6493a6cfbd0e75acc60ec3e7", out objectId);
            var filter = Builders<Classes.SofwareInfoModel>.Filter.Eq(x => x._id, objectId);
            var result = await ((await collection.FindAsync(filter)).ToListAsync());
            Version mVersion = new Version(result.First().version);
            if(mVersion > version)
            {
                if(MessageBox.Show("Latest Version is " + mVersion.ToString() + "\r\n" + "Current version is "+ version.ToString() +"\r\n" 
                    + "Have you saved all your SolidWorks Files? If Yes then you may proceed to upgrade by pressing the OK button below.", "Upgrade?", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    try
                    {
                        // get my document location
                        string myDocPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                        // create download path
                        string downloadPath = myDocPath + "\\mStructural Download";
                        Directory.CreateDirectory(downloadPath);

                        // get next version
                        string majorVersion = mVersion.Major.ToString();
                        string minorVersion = mVersion.Minor.ToString();

                        // create location string
                        string myUri = "https://raw.githubusercontent.com/MaxwellBTG/mStructuralPublish/main/mStructuralInstallerV" + majorVersion+"."+minorVersion+".zip";
                        string zipFile = downloadPath + "\\mStructuralInstallerV" + majorVersion + "." + minorVersion + ".zip";
                        string extractDir = downloadPath + "\\mStructuralInstallerV" + majorVersion + "." + minorVersion;

                        // download and save file to here
                        using (WebClient wc = new WebClient())
                        {
                            ServicePointManager.Expect100Continue = true;
                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                            wc.Headers.Add("a", "a");
                            await wc.DownloadFileTaskAsync(myUri, zipFile);
                        }

                        // check directory for extractdir
                        Directory.CreateDirectory(extractDir);

                        // unzip file
                        ZipFile.ExtractToDirectory(zipFile, extractDir);

                        // check directory
                        if (!Directory.Exists(extractDir))
                        {
                            MessageBox.Show("Download failed.");
                            return;
                        }

                        // run installer
                        Process process = new Process();
                        process.StartInfo.FileName = extractDir + "\\mStructuralInstaller.exe";
                        process.StartInfo.Verb = "runas";
                        process.Start();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("No Update available");
            }
        }
    }
}
