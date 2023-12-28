using Microsoft.VisualBasic;
using Microsoft.Win32;
using MongoDB.Driver;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace mStructuralInstaller
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static async Task Main()
        {
            // change this to false for actual release
            bool isDebugging = false;

            var assembly = typeof(Program).Assembly;
            GuidAttribute guidAtt = (GuidAttribute)assembly.GetCustomAttributes(typeof(GuidAttribute),true)[0];
            string guid = guidAtt.Value;
            Version curVersion = Assembly.GetExecutingAssembly().GetName().Version;
            Version existVersion = new Version();
            string installedLoc = "";
            string Status = "";

            string masterPlatformTarDir = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.System)) + @"\mStructural Data";

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // check if run as admin
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            bool IsAdmin = false;
            IsAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
            if (!IsAdmin)
            {
                MessageBox.Show("Please Run the Installer as Admin.");
                return;
            }

            // Check if SolidWorks is Running
            Process[] procs = Process.GetProcessesByName("SLDWORKS");
            bool bRet;
            if (procs.Length > 0)
            {
                if(MessageBox.Show("There are still some SolidWorks running in the background, you want to close it now?","SolidWorks Session detected!", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    foreach(Process proc in procs)
                    {
                        proc.Kill();
                        bRet = proc.WaitForExit(3000);
                        while (!bRet)
                        {
                            proc.Kill();
                            bRet = proc.WaitForExit(3000);
                        }
                    }
                }
            }

            // check if old registry exist
            bool IsInstalled = false;
            bool oldVersion = false;
            try
            {
                // check for old version
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\mStructural"))
                {
                    if (key != null)
                    {
                        if (MessageBox.Show("Old version found, do you want to upgrade to latest version?", "Upgrade", MessageBoxButtons.OKCancel) == DialogResult.OK)
                        {
                            IsInstalled = true;
                            oldVersion = true;
                            object o = "1.3.0.0";
                            existVersion = new Version(o as string);

                            object o2 = key.GetValue("ProgramLoc");
                            if (o2 != null)
                            {
                                installedLoc = (string)o2;
                            }
                        }
                        else
                        {
                            return;
                        }
                    }
                }

                if (oldVersion)
                {
                    // delete old registry
                    RegistryKey key = Registry.LocalMachine.OpenSubKey("Software", true);
                    key.DeleteSubKeyTree("mStructural");
                    key.Close();
                }
                else
                {
                    // read control panel registry
                    using (RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{" + guid.ToUpper() + "}"))
                    {
                        if (key != null)
                        {
                            IsInstalled = true;
                            object o = key.GetValue("DisplayVersion");
                            if (o != null)
                            {
                                existVersion = new Version(o as string);
                            }

                            object o2 = key.GetValue("InstallLocation");
                            if (o2 != null)
                            {
                                installedLoc = (string)o2;
                            }
                        }
                    }
                }

                if (isDebugging)
                {
                    await DebugDeploy(guid, Status, installedLoc, masterPlatformTarDir);
                }
                else
                {
                    // if dont have registry then proceed installation
                    if (!IsInstalled)
                    {
                        Status = "Install";
                        await Install(guid, Status, masterPlatformTarDir);
                    }
                    else
                    {
                        var result = curVersion.CompareTo(existVersion);

                        // if have current version then prompt for uninstallation
                        if (result > 0)
                        {
                            // Installer is newer
                            if (MessageBox.Show("Newer vesion will be install on this PC, press ok to proceed", "Upgrade", MessageBoxButtons.OKCancel) == DialogResult.OK)
                            {
                                Status = "Upgrade";

                                // Uninstall
                                Uninstall(guid,Status, installedLoc, masterPlatformTarDir);

                                // Install
                                await Install(guid, Status, masterPlatformTarDir);
                            }
                        }
                        else
                        {
                            // Already Installed
                            if (MessageBox.Show("Latest version of the software is already installed on this PC, do you want to uninstall?", "Uninstall", MessageBoxButtons.OKCancel) == DialogResult.OK)
                            {
                                Status = "Uninstall";

                                // Uninstall
                                Uninstall(guid, Status, installedLoc, masterPlatformTarDir);
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to read old Registry with error: " + ex.ToString());
                return;
            }
        }

        static async Task Install(string guid, string status, string masterplatloc)
        {
            if(status == "Install")
            {
                DateTime expirydate;
                string serialNo;

                // get mac address
                var macAddr =
                    (
                        from nic in NetworkInterface.GetAllNetworkInterfaces()
                        where nic.OperationalStatus == OperationalStatus.Up
                        select nic.GetPhysicalAddress().ToString()
                    ).FirstOrDefault();

                // get pc name
                string pcName = Environment.MachineName;

                // mongodb connection string
                const string connectionUri = "";
                var settings = MongoClientSettings.FromConnectionString(connectionUri);
                // Set the ServerApi field of the settings object to Stable API version 1
                settings.ServerApi = new ServerApi(ServerApiVersion.V1);
                // Create a new client and connect to the server
                var client = new MongoClient(settings);

                try
                {
                    // Get license expiry info
                    var db = client.GetDatabase("mStructural");
                    var collection = db.GetCollection<LicenseModel>("Licenses");
                    serialNo = Interaction.InputBox("Please enter your serial no");
                    var filter = Builders<LicenseModel>.Filter.Eq(x => x.SerialNo, serialNo);
                    var result = await ((await collection.FindAsync(filter)).ToListAsync());
                    if (result.First().ExpiryDate < DateTime.Now)
                    {
                        MessageBox.Show("License expired, please request for new one");
                        return;
                    }
                    else
                    {
                        expirydate = result.First().ExpiryDate;
                        // log an entry to database
                        var actCollect = db.GetCollection<ActivationModel>("Activation");
                        ActivationModel activation = new ActivationModel { PcName = pcName, MacAddress = macAddr.ToString(), Status = true, License = serialNo, TimeStamp = DateTime.Now};
                        await actCollect.InsertOneAsync(activation);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    return;
                }

                string targetDir = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + @"\mStructural";

                // Create directory
                Directory.CreateDirectory(targetDir);

                // Get Assembly Location (Name)
                string assemLoc = Assembly.GetExecutingAssembly().Location;

                // copy files from installer
                string sourcePath = Path.GetDirectoryName(assemLoc) + @"\source";
                copydir(sourcePath, targetDir);

                // Run Bat file
                ProcessStartInfo processInfo;
                Process process;
                bool rResult = false;

                processInfo = new ProcessStartInfo(targetDir + @"\Install.bat");
                processInfo.CreateNoWindow = true;
                processInfo.UseShellExecute = false;
                processInfo.RedirectStandardInput = true;
                processInfo.Verb = "runas";

                try
                {
                    process = Process.Start(processInfo);
                    rResult = true;

                    // if success, write to registry
                    if (rResult)
                    {
                        RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", true);

                        RegistryKey newkey = key.CreateSubKey("{"+guid.ToUpper()+"}");
                        newkey.SetValue("DisplayName", "mStructural");
                        newkey.SetValue("Publisher", "BTG");
                        newkey.SetValue("DisplayVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString());
                        newkey.SetValue("InstallLocation", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + @"\mStructural");
                        newkey.SetValue("UninstallString", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + @"\mStructural\mStructuralUninstaller.exe");
                        newkey.Close();
                        key.Close();
                    }

                    // create license string
                    string licenseInfo = "";
                    licenseInfo += macAddr.ToString();
                    licenseInfo += expirydate.ToString("yyyyMMdd");
                    licenseInfo += serialNo;

                    // Declare CspParameters and RsaCryptoServiceProvider
                    // objects with global scope of your Form class.
                    CspParameters _cspp = new CspParameters();
                    RSACryptoServiceProvider _rsa;

                    _cspp.KeyContainerName = "";
                    _rsa = new RSACryptoServiceProvider(_cspp)
                    {
                        PersistKeyInCsp = true
                    };

                    //FileInfo fileInfo = new FileInfo(targetDir + @"\mStructural.txt");
                    EncryptFile(licenseInfo, _rsa, targetDir + @"\");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    return;
                }

                MessageBox.Show("Installation Completed!");
            }

            if (status == "Upgrade")
            {
                DateTime expirydate;
                string serialNo;

                // get mac address
                var macAddr =
                    (
                        from nic in NetworkInterface.GetAllNetworkInterfaces()
                        where nic.OperationalStatus == OperationalStatus.Up
                        select nic.GetPhysicalAddress().ToString()
                    ).FirstOrDefault();

                // get pc name
                string pcName = Environment.MachineName;

                // mongodb connection string
                const string connectionUri = "mongodb+srv://maxwell:Mwellbend92@mstructural.zdeasje.mongodb.net/?retryWrites=true&w=majority";
                var settings = MongoClientSettings.FromConnectionString(connectionUri);
                // Set the ServerApi field of the settings object to Stable API version 1
                settings.ServerApi = new ServerApi(ServerApiVersion.V1);
                // Create a new client and connect to the server
                var client = new MongoClient(settings);

                try
                {
                    // Get license expiry info
                    var db = client.GetDatabase("mStructural");
                    var collection = db.GetCollection<LicenseModel>("Licenses");
                    serialNo = Interaction.InputBox("Please enter your serial no");
                    var filter = Builders<LicenseModel>.Filter.Eq(x => x.SerialNo, serialNo);
                    var result = await ((await collection.FindAsync(filter)).ToListAsync());
                    if (result.First().ExpiryDate < DateTime.Now)
                    {
                        MessageBox.Show("License expired, please request for new one");
                        return;
                    }
                    else
                    {
                        expirydate = result.First().ExpiryDate;
                        // log an entry to database
                        var actCollect = db.GetCollection<ActivationModel>("Activation");
                        ActivationModel activation = new ActivationModel { PcName = pcName, MacAddress = macAddr.ToString(), Status = true, License = serialNo, TimeStamp = DateTime.Now };
                        await actCollect.InsertOneAsync(activation);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    return;
                }

                string targetDir = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + @"\mStructural";

                // Create directory
                Directory.CreateDirectory(targetDir);

                // Get Assembly Location (Name)
                string assemLoc = Assembly.GetExecutingAssembly().Location;

                // copy files from installer
                string sourcePath = Path.GetDirectoryName(assemLoc) + @"\source";
                copydir(sourcePath, targetDir);

                // Run Bat file
                ProcessStartInfo processInfo;
                Process process;
                bool rResult = false;

                processInfo = new ProcessStartInfo(targetDir + @"\Install.bat");
                processInfo.CreateNoWindow = true;
                processInfo.UseShellExecute = false;
                processInfo.RedirectStandardInput = true;
                processInfo.Verb = "runas";

                try
                {
                    process = Process.Start(processInfo);
                    rResult = true;

                    // if success, write to registry
                    if (rResult)
                    {
                        RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", true);

                        RegistryKey newkey = key.CreateSubKey("{" + guid.ToUpper() + "}");
                        newkey.SetValue("DisplayName", "mStructural");
                        newkey.SetValue("Publisher", "BTG");
                        newkey.SetValue("DisplayVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString());
                        newkey.SetValue("InstallLocation", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + @"\mStructural");
                        newkey.SetValue("UninstallString", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + @"\mStructural\mStructuralUninstaller.exe");
                        newkey.Close();
                        key.Close();
                    }

                    // create license string
                    string licenseInfo = "";
                    licenseInfo += macAddr.ToString();
                    licenseInfo += expirydate.ToString("yyyyMMdd");
                    licenseInfo += serialNo;

                    // Declare CspParameters and RsaCryptoServiceProvider
                    // objects with global scope of your Form class.
                    CspParameters _cspp = new CspParameters();
                    RSACryptoServiceProvider _rsa;

                    _cspp.KeyContainerName = "";
                    _rsa = new RSACryptoServiceProvider(_cspp)
                    {
                        PersistKeyInCsp = true
                    };

                    //FileInfo fileInfo = new FileInfo(targetDir + @"\mStructural.txt");
                    EncryptFile(licenseInfo, _rsa, targetDir + @"\");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    return;
                }

                MessageBox.Show("Upgrade Completed!");
            }
        }

        static void Uninstall(string guid, string status, string installloc, string masterplatloc)
        {
            // Run Bat file
            ProcessStartInfo processInfo;
            Process process;
            bool rResult = false;

            processInfo = new ProcessStartInfo(installloc + @"\Uninstall.bat");
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = false;
            processInfo.RedirectStandardInput = true;
            processInfo.Verb = "runas";
            try
            {
                process = Process.Start(processInfo);
                int count = 0;

                do
                {
                    if (!process.HasExited)
                    {
                        if(count > 10)
                        {
                            MessageBox.Show("Uninstallation failed, please contact admin");
                            return;
                        }
                        count++;
                    }
                }
                while(!process.WaitForExit(500));

                /*
                // get my document location
                string myDocPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                // create download path
                string backupPath = myDocPath + "\\mStructural Download\\Backup\\" + DateTime.Now.ToString("yyMMddHHmm");
                
                // create backup directory
                Directory.CreateDirectory(backupPath);

                // destination file loaction
                string backupFile = backupPath + "\\dtlbstring.txt";

                // original location
                string fileToBackup = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\\mStructural\\Settings\\dtlbstring.txt";

                // move the file
                File.Copy(fileToBackup, backupFile);
                */

                // delete directory
                Directory.Delete(installloc, true);

                rResult = true;

                // if success, delete registry
                if (rResult)
                {
                    RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", true);
                    key.DeleteSubKeyTree("{"+guid.ToUpper()+"}");
                    key.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }

            if (status == "Uninstall")
                MessageBox.Show("Uninstallation completed!");
        }

        static void copydir(string sourcedir, string destdir)
        {
            try
            {
                foreach (string dirPath in Directory.GetDirectories(sourcedir, "*.*", SearchOption.AllDirectories))
                {
                    Directory.CreateDirectory(dirPath.Replace(sourcedir, destdir));
                }

                foreach (var file in Directory.GetFiles(sourcedir, "*.*", SearchOption.AllDirectories))
                {
                    File.Copy(file, file.Replace(sourcedir, destdir), true);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        static async Task DebugDeploy(string guid, string status, string installloc, string masterplatloc)
        {
            // uninstall first
            // Run Bat file
            ProcessStartInfo processInfo;
            Process process;
            bool rResult = false;

            processInfo = new ProcessStartInfo(installloc + @"\Uninstall.bat");
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = false;
            processInfo.RedirectStandardInput = true;
            processInfo.Verb = "runas";
            try
            {
                process = Process.Start(processInfo);
                int count = 0;

                do
                {
                    if (!process.HasExited)
                    {
                        if (count > 10)
                        {
                            MessageBox.Show("Uninstallation failed, please contact admin");
                            return;
                        }
                        count++;
                    }
                }
                while (!process.WaitForExit(500));

                // delete directory
                Directory.Delete(installloc, true);

                rResult = true;

                // if success, delete registry
                if (rResult)
                {
                    RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", true);
                    key.DeleteSubKeyTree("{" + guid.ToUpper() + "}");
                    key.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
            // then reinstall here


            DateTime expirydate;

            // get mac address
            var macAddr =
                (
                    from nic in NetworkInterface.GetAllNetworkInterfaces()
                    where nic.OperationalStatus == OperationalStatus.Up
                    select nic.GetPhysicalAddress().ToString()
                ).FirstOrDefault();

            // get pc name
            string pcName = Environment.MachineName;

            // mongodb connection string
            const string connectionUri = "mongodb+srv://maxwell:Mwellbend92@mstructural.zdeasje.mongodb.net/?retryWrites=true&w=majority";
            var settings = MongoClientSettings.FromConnectionString(connectionUri);
            // Set the ServerApi field of the settings object to Stable API version 1
            settings.ServerApi = new ServerApi(ServerApiVersion.V1);
            // Create a new client and connect to the server
            var client = new MongoClient(settings);

            try
            {
                // Get license expiry info
                var db = client.GetDatabase("mStructural");
                var collection = db.GetCollection<LicenseModel>("Licenses");
                string serialNo = "6db4bb24-2816-4839-929e-cc6a45af28dd";
                var filter = Builders<LicenseModel>.Filter.Eq(x => x.SerialNo, serialNo);
                var result = await ((await collection.FindAsync(filter)).ToListAsync());
                if (result.First().ExpiryDate < DateTime.Now)
                {
                    MessageBox.Show("License expired, please request for new one");
                    return;
                }
                else
                {
                    expirydate = result.First().ExpiryDate;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }

            string targetDir = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + @"\mStructural";

            // Create directory
            Directory.CreateDirectory(targetDir);

            // Get Assembly Location (Name)
            string assemLoc = Assembly.GetExecutingAssembly().Location;

            // copy files from installer
            string sourcePath = Path.GetDirectoryName(assemLoc) + @"\source";
            copydir(sourcePath, targetDir);

            // Run Bat file
            //ProcessStartInfo processInfo;
            //Process process;
            rResult = false;

            processInfo = new ProcessStartInfo(targetDir + @"\Install.bat");
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = false;
            processInfo.RedirectStandardInput = true;
            processInfo.Verb = "runas";

            try
            {
                process = Process.Start(processInfo);
                rResult = true;

                // if success, write to registry
                if (rResult)
                {
                    RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", true);

                    RegistryKey newkey = key.CreateSubKey("{" + guid.ToUpper() + "}");
                    newkey.SetValue("DisplayName", "mStructural");
                    newkey.SetValue("Publisher", "BTG");
                    newkey.SetValue("DisplayVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString());
                    newkey.SetValue("InstallLocation", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + @"\mStructural");
                    newkey.SetValue("UninstallString", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + @"\mStructural\mStructuralUninstaller.exe");
                    newkey.Close();
                    key.Close();
                }

                // create license string
                string licenseInfo = "";
                licenseInfo += macAddr.ToString();
                licenseInfo += expirydate.ToString("yyyyMMdd");

                // Declare CspParameters and RsaCryptoServiceProvider
                // objects with global scope of your Form class.
                CspParameters _cspp = new CspParameters();
                RSACryptoServiceProvider _rsa;

                _cspp.KeyContainerName = "";
                _rsa = new RSACryptoServiceProvider(_cspp)
                {
                    PersistKeyInCsp = true
                };

                //FileInfo fileInfo = new FileInfo(targetDir + @"\mStructural.txt");
                EncryptFile(licenseInfo, _rsa, targetDir + @"\");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        static void EncryptFile(string myStr, RSACryptoServiceProvider _rsa, string EncrFolder)
        {
            // Create instance of Aes for
            // symmetric encryption of the data.
            Aes aes = Aes.Create();
            aes.Padding = PaddingMode.Zeros;
            ICryptoTransform transform = aes.CreateEncryptor();

            // Use RSACryptoServiceProvider to
            // encrypt the AES key.
            // rsa is previously instantiated:
            //    rsa = new RSACryptoServiceProvider(cspp);
            byte[] keyEncrypted = _rsa.Encrypt(aes.Key, false);

            // Create byte arrays to contain
            // the length values of the key and IV.
            int lKey = keyEncrypted.Length;
            byte[] LenK = BitConverter.GetBytes(lKey);
            int lIV = aes.IV.Length;
            byte[] LenIV = BitConverter.GetBytes(lIV);

            // Write the following to the FileStream
            // for the encrypted file (outFs):
            // - length of the key
            // - length of the IV
            // - encrypted key
            // - the IV
            // - the encrypted cipher content

            // Change the file's extension to ".lic"
            string outFile =
                Path.Combine(EncrFolder, "mStructural.lic");

            using (var outFs = new FileStream(outFile, FileMode.Create))
            {
                outFs.Write(LenK, 0, 4);
                outFs.Write(LenIV, 0, 4);
                outFs.Write(keyEncrypted, 0, lKey);
                outFs.Write(aes.IV, 0, lIV);

                // Now write the cipher text using
                // a CryptoStream for encrypting.
                using (var outStreamEncrypted =
                    new CryptoStream(outFs, transform, CryptoStreamMode.Write))
                {
                    Encoding ascii = Encoding.ASCII;
                    byte[] myStrByte = ascii.GetBytes(myStr);

                    outStreamEncrypted.Write(myStrByte, 0, myStrByte.Length);
                }
            }
        }

        static string DecryptFile(FileInfo file)
        {
            // encrypt license file
            CspParameters _cspp = new CspParameters();
            RSACryptoServiceProvider _rsa;
            _cspp.KeyContainerName = "";
            _rsa = new RSACryptoServiceProvider(_cspp)
            {
                PersistKeyInCsp = true
            };

            // Create instance of Aes for
            // symmetric decryption of the data.
            Aes aes = Aes.Create();
            aes.Padding = PaddingMode.Zeros;

            // Create byte arrays to get the length of
            // the encrypted key and IV.
            // These values were stored as 4 bytes each
            // at the beginning of the encrypted package.
            byte[] LenK = new byte[4];
            byte[] LenIV = new byte[4];

            // Construct the file name for the decrypted file.
            // string outFile = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop) + @"\mStructural\license.txt";

            // Use FileStream objects to read the encrypted
            // file (inFs) and save the decrypted file (outFs).
            using (var inFs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read))
            {
                inFs.Seek(0, SeekOrigin.Begin);
                inFs.Read(LenK, 0, 3);
                inFs.Seek(4, SeekOrigin.Begin);
                inFs.Read(LenIV, 0, 3);

                // Convert the lengths to integer values.
                int lenK = BitConverter.ToInt32(LenK, 0);
                int lenIV = BitConverter.ToInt32(LenIV, 0);

                // Determine the start position of
                // the cipher text (startC)
                // and its length(lenC).
                int startC = lenK + lenIV + 8;
                int lenC = (int)inFs.Length - startC;

                // Create the byte arrays for
                // the encrypted Aes key,
                // the IV, and the cipher text.
                byte[] KeyEncrypted = new byte[lenK];
                byte[] IV = new byte[lenIV];

                // Extract the key and IV
                // starting from index 8
                // after the length values.
                inFs.Seek(8, SeekOrigin.Begin);
                inFs.Read(KeyEncrypted, 0, lenK);
                inFs.Seek(8 + lenK, SeekOrigin.Begin);
                inFs.Read(IV, 0, lenIV);

                // Use RSACryptoServiceProvider
                // to decrypt the AES key.
                byte[] KeyDecrypted = _rsa.Decrypt(KeyEncrypted, false);

                // Decrypt the key.
                ICryptoTransform transform = aes.CreateDecryptor(KeyDecrypted, IV);

                // Decrypt the cipher text from
                // from the FileSteam of the encrypted
                // file (inFs) into the FileStream
                // for the decrypted file (outFs).
                // int count = 0;
                // int offset = 0;

                // blockSizeBytes can be any arbitrary size.
                int blockSizeBytes = aes.BlockSize / 8;
                byte[] data = new byte[blockSizeBytes];

                // By decrypting a chunk a time,
                // you can save memory and
                // accommodate large files.

                // Start at the beginning
                // of the cipher text.
                inFs.Seek(startC, SeekOrigin.Begin);
                using (var outStreamDecrypted =
                    new CryptoStream(inFs, transform, CryptoStreamMode.Read))
                {
                    using (var streamReader = new StreamReader(outStreamDecrypted))
                    {
                        return streamReader.ReadToEnd();
                    }
                }
            }
        }
    }
}
