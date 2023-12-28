using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Windows;

namespace mStructural
{
    /// <summary>
    /// SolidWorks Taskpane Add-in
    /// </summary>

    [Guid("494DC02B-31E7-4F92-BB90-3FD185A60F35")]
    public class SWIntegration : ISwAddin
    {
        #region Private Members
        // taskpane view for our add-in
        private TaskpaneView mTaskpaneView;

        // UI in taskpane view
        private TaskpaneHostUIWin mTaskpaneHostUI;
        #endregion

        #region Public Members
        // unique progid for the addin registration in COM
        public const string SWTASKPANE_PROGID = "mStructural.SolidWorks.Taskpane.Addin";

        // Current solidworks instance
        public SldWorks SwApp;

        // Current solidworks cookie id
        public int CookieId;

        // Expiry Date
        public string ExpiryDateStr;

        // Mac Address String
        // Removed in v1.6
        // public string MacAddr;
        #endregion

        #region SolidWorks Add-in callbacks

        /// <summary>
        /// Called when solidworks has loaded our add-in and wants us to do our connection logic
        /// </summary>
        /// <param name="ThisSW"> The current SolidWorks instance</param>
        /// <param name="Cookie"> The current SolidWorks cookie Id</param>
        /// <returns></returns>
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            // Store reference to the current solidworks instance.
            SwApp = (SldWorks)ThisSW; 

            // Store reference to the current solidworks cookie id.
            CookieId = Cookie;

            // Setup callback info
            var ok = SwApp.SetAddinCallbackInfo2(0, this, CookieId);

            string licenseInfo = "";
            try
            {
                // Check License
                string licensePath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles) + @"\mStructural\mStructural.lic";
                FileInfo fileInfo= new FileInfo(licensePath);
                licenseInfo =  DecryptFile(fileInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }

            if(licenseInfo == "")
            {
                MessageBox.Show("Failed to obtain license info");
                return false;
            }
            else
            {
                // v1.6 removed macaddr
                // MacAddr = licenseInfo.Substring(0, 12);
                ExpiryDateStr = licenseInfo.Substring(12, 8);
                DateTime expiryDate = DateTime.ParseExact(ExpiryDateStr, "yyyyMMdd", CultureInfo.InvariantCulture);

                if (expiryDate < DateTime.Now)
                {
                    MessageBox.Show("License expired");
                    return false;
                }

                // get mac address
                // v1.6 removed checking for mac address
                /*var thisMacAddr =
                    (
                        from nic in NetworkInterface.GetAllNetworkInterfaces()
                        where nic.OperationalStatus == OperationalStatus.Up
                        select nic.GetPhysicalAddress().ToString()
                    ).FirstOrDefault();

                if (thisMacAddr.ToString() != MacAddr)
                {
                    MessageBox.Show("Unmatched PC Info");
                    return false;
                }*/

                // load ui
                LoadUI();
            }

            // return true when completed loading taskpane
            return true;
        }

        /// <summary>
        /// Called when solidworks has unloaded our add-in and wants us to do our disconnection logic
        /// </summary>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public bool DisconnectFromSW()
        {
            // clean up ui
            UnloadUI();

            // return true when completed unloading taskpane
            return true;
        }

        #endregion

        #region Create UI
        // function to load ui
        private void LoadUI()
        {
            // path to taskpane icon
            string[] bitmap = new string[1];
            bitmap[0] = Path.Combine(Path.GetDirectoryName(typeof(SWIntegration).Assembly.CodeBase).Replace(@"file:\", ""), "MainIcon.bmp");

            // Create taskpane
            mTaskpaneView = SwApp.CreateTaskpaneView2(bitmap[0], "mStructural Add-in");

            // Create new Winform UI
            mTaskpaneHostUI = new TaskpaneHostUIWin(this);

            // Add winform into taskpaneview
            mTaskpaneView.DisplayWindowFromHandlex64(mTaskpaneHostUI.Handle.ToInt64());
        }

        // function to clean up ui
        private void UnloadUI()
        {
            // set host ui as null
            mTaskpaneHostUI = null;

            // remove taskpane view
            mTaskpaneView.DeleteView();

            // Release COM reference
            Marshal.ReleaseComObject(mTaskpaneView);

            // set taskpane view as null
            mTaskpaneView = null;
        }

        #endregion

        #region COM Registration
        
        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            // key path for registry
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            using (var rk = Registry.LocalMachine.CreateSubKey(keyPath))
            {
                // Load add-in when solidworks opens
                rk.SetValue(null, 1);

                // Set Title and description
                rk.SetValue("Title", "mStructural");
                rk.SetValue("Description", "mStructural customize solidworks add-in ");
            }
        }

        [ComUnregisterFunction()]
        private static void ComUnregister(Type t)
        {
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\SolidWorks\AddIns", true);
            key.DeleteSubKeyTree("{" + t.GUID + "}");
            key.Close();
        }

        #endregion

        #region cryptography
        private string DecryptFile(FileInfo file)
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
                    using(var streamReader = new StreamReader(outStreamDecrypted))
                    {
                        return streamReader.ReadToEnd();
                    }
                }
            }
        }
        #endregion
    }
}
