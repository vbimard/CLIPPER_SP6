using Actcut.ActcutModelManager;
using Actcut.ResourceManager;
using AF_ImportTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Wpm.Implement.Manager;
//
using System.Reflection;
using System.Resources;
using System.Globalization;
//
namespace AlmaCamTrainingTest_WPF
{
   
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ResourceManager rm = new ResourceManager("MyResource", Assembly.GetExecutingAssembly());
        public MainWindow()
        {
            //InitializeComponent();
            //ResourceManager rm = new ResourceManager("UsingRESX.MyResource",  Assembly.GetExecutingAssembly());
            //String strWebsite = rm.GetString("Website", CultureInfo.CurrentCulture);
            //String strName = rm.GetString("Name");
            //form1.InnerText = "Website: " + strWebsite + "--Name: " + strName;
        }

        private void mnuNew_Exit(object sender, RoutedEventArgs e)
        {

            this.Close();
            
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
             //ReportManager.ge.Resources.Entities_Clip;
          
            //File f = rm.
        }
            /*
            private void MenuItem_Click(object sender, RoutedEventArgs e)
            {
                ///
                //import de la machine clipper
                SimplifiedMethods.NotifyMessage("AF Clipper", "Starting installation process...please wait");


                //IList<CustomizedCommandItem> customizedItemList = GetCustomizedCommandTypeList(context);

                IModelsRepository modelsRepository = new ModelsRepository();

                IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);

                List<Tuple<string, string>> CustoFileCsList = new List<Tuple<string, string>>();

                CustoFileCsList.Add(new Tuple<string, string>("Entities", Properties.Resources.Entities_Clip));
                CustoFileCsList.Add(new Tuple<string, string>("FormulasAndEvents", Properties.Resources.Event_Clip));
                CustoFileCsList.Add(new Tuple<string, string>("Commandes", Properties.Resources.Commandes_Clip));

                foreach (Tuple<string, string> CustoFileCs in CustoFileCsList)
                {
                    string CommandCsPath = System.IO.Path.GetTempPath() + CustoFileCs.Item1 + ".cs";
                    File.WriteAllText(CommandCsPath, CustoFileCs.Item2);

                    string cs = CustoFileCs.Item1;
                    SimplifiedMethods.NotifyMessage("Updating Database", CustoFileCs.Item1);
                    ModelManager modelManager = new ModelManager(modelsRepository);
                    //detection des composants odbc 
                    //activateProgressBar(" Updating " + CustoFileCs.Item1 + " in " + _Context.Connection.DatabaseName);
                    modelManager.CustomizeModel(CommandCsPath, Lst_Model.Text, false);
                    //desactivateProgressBar();

                    File.Delete(CommandCsPath);
                }

                SimplifiedMethods.NotifyMessage("Updating Database", "Done");


                ///update parameters
                //chaine de connexion// sp5 //
                //import machineè
                ImportMachine();

                Update_Clipper_Parameters(_Context);
                Update_Clipper_Stock(_Context);
                //
                //mise a jour des champs
                Update_DefaultValues(_Context);

                MessageBox.Show("Database " + _Context.Connection.DatabaseName + " Prepared for clipper");

            }

            public void ImportMachine()
            {
                try
                {

                    MachineManager machineManager = new MachineManager();
                    IModelsRepository modelsRepository = new ModelsRepository();


                    _Context = modelsRepository.GetModelContext(Lst_Model.Text);
                    string ZipFileName = null;


                    ///activateProgressBar(" Updating Machine in " + _Context.Connection.DatabaseName);
                    //
                    IEntityList Current_MachineList = _Context.EntityManager.GetEntityList("_CUT_MACHINE_TYPE", "_NAME", ConditionOperator.Equal, "Clipper");
                    Current_MachineList.Fill(false);

                    //import cam machine
                    if (Current_MachineList.Count == 0)
                    {
                        ZipFileName = System.IO.Path.GetTempPath() + "Clipper.zip";
                        //ecriture de la machine .zip
                        File.WriteAllBytes(ZipFileName, Properties.Resources.Clipper_Machine);
                        SimplifiedMethods.NotifyMessage("Updating Database", "Clipper machine not detected...Installating Clipper machine...");
                        MachineImporter importMachineEntity = new MachineImporter();
                        importMachineEntity.Init(_Context);//, out string ErrorMessage);
                        importMachineEntity.ReadZipFile(ZipFileName);
                        importMachineEntity.ImportCutMachine(true);


                        File.Delete(ZipFileName);
                    }

                    SimplifiedMethods.NotifyMessage("Updating Database", "Import Clipper machine Done...");

                    machineManager = null;
                    modelsRepository = null;
                   /// desactivateProgressBar();
                }




                catch { }





            }

            */
        }
}
