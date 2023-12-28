using Microsoft.Toolkit.Uwp.Notifications;
using MongoDB.Bson;
using MongoDB.Driver;
using mStructural.Function;
using mStructural.MacroFeatures;
using mStructural.WPF;
using System;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace mStructural
{
    /// <summary>
    /// Interaction logic for TaskpaneHostUI.xaml
    /// </summary>

    public partial class TaskpaneHostUI : UserControl
    {
        private SWIntegration swintegrationInstance;
        private string versionStr;

        public TaskpaneHostUI(SWIntegration swintegration)
        {
            InitializeComponent();
            swintegrationInstance = swintegration;
            versionStr = Assembly.GetExecutingAssembly().GetName().Version.Major.ToString()
                + "."
                + Assembly.GetExecutingAssembly().GetName().Version.Minor.ToString()
                + "."
                + Assembly.GetExecutingAssembly().GetName().Version.Build.ToString(); // v1.8.0 added build version
            VersionLabel.Content = "Version: " + versionStr;
            ExpiryDateLabel.Content = "Expiry Date: " + swintegration.ExpiryDateStr.Substring(0, 4) + 
                "-" + swintegration.ExpiryDateStr.Substring(4, 2) + 
                "-" + swintegration.ExpiryDateStr.Substring(6, 2);
            MacAddrLabel.Content = "PC: " + Environment.MachineName;
            SendUpdateToastAsync();
        }

        private void SettingBtn_Click(object sender, RoutedEventArgs e)
        {
            Setting settingForm = new Setting();
            bool? dialogResult = settingForm.ShowDialog();
            switch (dialogResult)
            {
                case true:
                    IsEnabled = true;
                    break;
                case false:
                    IsEnabled = true;
                    break;
                default:
                    IsEnabled = true;
                    break;
            }
        }

        private void CreateBodyViewBtn_Click(object sender, RoutedEventArgs e)
        {
            CreateAllBodyView createAllBodyView = new CreateAllBodyView(swintegrationInstance);
            try
            {
                createAllBodyView.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void SortCutlistBtn_Click(object sender, RoutedEventArgs e)
        {
            SortCutlist sortCutlist = new SortCutlist(swintegrationInstance);
            try
            {
                sortCutlist.Run();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void DimProfileBtn_Click(object sender, RoutedEventArgs e)
        {
            DimProfile dimProfile = new DimProfile(swintegrationInstance);
            try
            {
                dimProfile.Run();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void CreateAllBodiesBtn_Click(object sender, RoutedEventArgs e)
        {
            AutoBodyView autoBodyView = new AutoBodyView(swintegrationInstance.SwApp);
            autoBodyView.Run();
        }

        private void BalloonViewBtn_Click(object sender, RoutedEventArgs e)
        {
            BalloonView balloonView = new BalloonView(swintegrationInstance);
            try
            {
                balloonView.Run();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void PrepDxfBtn_Click(object sender, RoutedEventArgs e)
        {
            PrepDxfEntry prepDxfEntry = new PrepDxfEntry(swintegrationInstance.SwApp);
            try
            {
                prepDxfEntry.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void SaveDxfBtn_Click(object sender, RoutedEventArgs e)
        {
            ExportDxfEntry exportDxfEntry = new ExportDxfEntry(swintegrationInstance.SwApp);
            try
            {
                exportDxfEntry.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ExportBomBtn_Click(object sender, RoutedEventArgs e)
        {
            ExportBom exportBom = new ExportBom(swintegrationInstance);
            try
            {
                exportBom.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void StartModelBtn_Click(object sender, RoutedEventArgs e)
        {
            StartModel startModel = new StartModel(swintegrationInstance);
            try
            {
                startModel.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void UpdateBtn_Click(object sender, RoutedEventArgs e)
        {
            Update update = new Update();
            try
            {
                update.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private async Task SendUpdateToastAsync()
        {

            try
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
                var result = await((await collection.FindAsync(filter)).ToListAsync());
                Version mVersion = new Version(result.First().version);
                if (mVersion > version)
                {
                    // create a toast
                    new ToastContentBuilder()
                        .AddText("mStructural V" + mVersion.Major.ToString() + "." + mVersion.Minor.ToString() + "." + mVersion.Build.ToString()) // added build version
                        .AddText("New version is available, press the upgrade button in your taskpane to upgrade.")
                        .Show();
                }
            }
            catch (Exception)
            {
                return;
            }
        }

        private void QuickNoteBtn_Click(object sender, RoutedEventArgs e)
        {
            QuickNote qn = new QuickNote(swintegrationInstance);
            qn.Show();
        }

        private void PlatformAutomationBtn_Click(object sender, RoutedEventArgs e)
        {
            PlatformAutomationEntry paEntry = new PlatformAutomationEntry(swintegrationInstance.SwApp);
            try
            {
                paEntry.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ExportPdfBtn_Click(object sender, RoutedEventArgs e)
        {
            ExportPdfEntry exportPdfEntry = new ExportPdfEntry(swintegrationInstance.SwApp);
            try
            {
                exportPdfEntry.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void EtchCutBtn_Click(object sender, RoutedEventArgs e)
        {
            EtchFeaturePMP etchFeatPMP = new EtchFeaturePMP(swintegrationInstance.SwApp, 0, null);
            try
            {
                etchFeatPMP.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void PlateWeldmentNoteBtn_Click(object sender, RoutedEventArgs e)
        {
            PlateWeldNote plateWeldNote = new PlateWeldNote(swintegrationInstance.SwApp);
            try
            {
                plateWeldNote.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void AlignLongestEdgeBtn_Click(object sender, RoutedEventArgs e)
        {
            AlignLongestEdge alignLongestEdge = new AlignLongestEdge(swintegrationInstance.SwApp);
            try
            {
                alignLongestEdge.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void DeleteTubeLaserBodyBtn_Click(object sender, RoutedEventArgs e)
        {
            DeleteTubeLaserBody deleteTubeLaserBody = new DeleteTubeLaserBody(swintegrationInstance.SwApp);
            try
            {
                deleteTubeLaserBody.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void NamingProject_Click(object sender, RoutedEventArgs e)
        {
            NamingProject namingProject = new NamingProject(swintegrationInstance.SwApp);
            try
            {
                namingProject.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ExportModelBtn_Click(object sender, RoutedEventArgs e)
        {
            ExportModel exportModel = new ExportModel(swintegrationInstance.SwApp);
            try
            {
                exportModel.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ExperimentalFunctionBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void LinkBomBtn_Click(object sender, RoutedEventArgs e)
        {
            LinkBomEntry linkBomEntry = new LinkBomEntry(swintegrationInstance.SwApp);
            try
            {
                linkBomEntry.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void EtchViewBtn_Click(object sender, RoutedEventArgs e)
        {
            EtchView etchView = new EtchView(swintegrationInstance.SwApp);
            try
            {
                etchView.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ProjectExportBtn_Click(object sender, RoutedEventArgs e)
        {
            ExportProjectWPF exportProjectWPF = new ExportProjectWPF(swintegrationInstance.SwApp);
            exportProjectWPF.Show();
        }

        private void HideRef_Click(object sender, RoutedEventArgs e)
        {
            HideRefWPF hideRefWPF = new HideRefWPF(swintegrationInstance.SwApp);
            hideRefWPF.Show();
        }

        private void ColourViewBtn_Click(object sender, RoutedEventArgs e)
        {
            ColourView colourView = new ColourView(swintegrationInstance.SwApp);
            try
            {
                colourView.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
