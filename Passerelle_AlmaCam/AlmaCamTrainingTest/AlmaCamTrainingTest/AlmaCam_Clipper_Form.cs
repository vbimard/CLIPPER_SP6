using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Wpm.Implement.Manager;
using Wpm.Implement.Processor;
using Wpm.Implement.ComponentEditor;  // ouverture de fenetres de selection almacam
using Wpm.Schema.Kernel;



//dll personnalisées
using AF_Clipper_Dll;
using AF_ImportTools;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;
using Wpm.Implement.ModelSetting;
using Alma.BaseUI.DescriptionEditor;
using Actcut.ResourceManager;
using Actcut.ActcutModelManagerUI;
using Alma.BaseUI.ErrorMessage;
using AF_Import_ODBC_Clipper_AlmaCam;
using Actcut.ActcutModelManager;

//suppression de wpm.schema.component.dll et remplacement par wpm.schema.componenteditor.dll 

namespace AlmaCamTrainingTest
{



    public partial class AlmaCam_Clipper_Form : Form, IDisposable
    {

        //initialisation des listes
        IContext _Context = null;

        string DbName = Alma_RegitryInfos.GetLastDataBase();

        public AlmaCam_Clipper_Form()
        {
            try
            {

                InitializeComponent();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

        }
        
        /// <summary>
        /// form load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            AF_ImportTools.SimplifiedMethods.NotifyMessage("ALMACAM CLIPPER", "Starting Almacam Clipper", 20000);
            //creation du model repository
            var databases = new List<string>();
            IModelsRepository modelsRepository = new ModelsRepository();
            foreach (IModel model in modelsRepository.ModelList)
            {
                databases.Add(model.DatabaseName);

            }
            //nom de la base;
            Lst_Model.DataSource = databases;
            Lst_Model.SelectedItem = DbName;
            string infosPasserelle;

            //

            //_Context = modelsRepository.GetModelContext(DbName);
            //int i = _Context.ModelsRepository.ModelList.Count();
            SimplifiedMethods.NotifyMessage("AlmaCam", " Selected " + DbName);
            infosPasserelle = DbName + "-P." + AF_Clipper_Dll.Clipper_Param.GetClipperDllVersion() + "-CAM." + AF_Clipper_Dll.Clipper_Param.GetAlmaCAMCompatibleVersion();
            this.Text = this.Name;
            this.InfosLabel.Text = infosPasserelle;
            this.Text = "Passerelle Clipper V8 validée pour : " + infosPasserelle;
        }
               
        #region Menus
       
        private void quitToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
 
        //purge du stock (a modifier)
        private void purgerStockToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {

                IModelsRepository modelsRepository = new ModelsRepository();
                _Context = modelsRepository.GetModelContext(Lst_Model.Text);

                IEntityList stocks = _Context.EntityManager.GetEntityList("_STOCK");
                stocks.Fill(false);
                DialogResult res = MessageBox.Show("Do you really want to destroy all sheets from the stock?", "Warnig", MessageBoxButtons.OKCancel);
                if (res == DialogResult.OK)
                {
                    foreach (IEntity stock in stocks)
                    {
                        stock.Delete();
                    }


                    IEntityList formats = _Context.EntityManager.GetEntityList("_SHEET");
                    formats.Fill(false);

                    foreach (IEntity format in formats)
                    {
                        format.Delete();
                    }

                }



            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);


            }
            finally { _Context = null; }


        }
       //
        private void suppressionStockClotureToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string stage = "_CLOSED_NESTING";

            IEntitySelector nestingselector = null;

            nestingselector = new EntitySelector();

            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;

            //entity type pointe sur la list d'objet du model
            nestingselector.Init(_Context, _Context.Kernel.GetEntityType(stage));
            nestingselector.MultiSelect = true;


            if (nestingselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (IEntity nesting in nestingselector.SelectedEntity)
                {
                    StockManager.DeleteAlmaCamStock(nesting);

                }
            }


            _Context = null;
        }
        /// <summary>
        /// purge le stock clipper suppression des toel avec le stock
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void purgerToutLeStockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                IModelsRepository modelsRepository = new ModelsRepository();
                _Context = modelsRepository.GetModelContext(Lst_Model.Text);

                //purge stock
                IEntityList stocks = _Context.EntityManager.GetEntityList("_STOCK");
                stocks.Fill(false);
                DialogResult res = MessageBox.Show("Do you really want to destroy all unused sheets from the stock?", "Warning", MessageBoxButtons.OKCancel);
                if (res == DialogResult.OK)
                {
                    foreach (IEntity stock in stocks)
                    {
                        if (stock.GetFieldValueAsEntity("_SHEET").GetFieldValueAsLong("_IN_PRODUCTION_QUANTITY") == 0)
                        {
                            //stock.Delete();

                            if (stock.GetFieldValueAsLong("_USED_QUANTITY") == 0)
                            {
                                stock.Delete();
                            }
                            else if (stock.GetFieldValueAsLong("_BOOKED_QUANTITY") == 0)
                            {
                                stock.Delete();
                            }

                        }
                    }

                    //suppression formats
                    IEntityList formats = _Context.EntityManager.GetEntityList("_SHEET");
                    formats.Fill(false);

                    foreach (IEntity format in formats)
                    {
                        if (format.GetFieldValueAsLong("_IN_PRODUCTION_QUANTITY") == 0)
                        {
                            format.Delete();
                        }

                    }

                }

                MessageBox.Show("Fin...");
            }



            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);

            }
            finally { _Context = null; }
        }
        /// <summary>
        /// import du dispomat.csv fourni par clipper (fichier listant le sotck a importer)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void importerStockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            using (var Stock = new Clipper_8_Import_Stock())
            {
                Stock.Import(_Context);//, csvImportPath);
            }

            _Context = null;

        }
       /// <summary>
       /// import du cahier d'affaire avec dt
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
        private void avecDTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            Clipper_Param.GetlistParam(_Context);
            string csvImportPath = Clipper_Param.GetPath("IMPORT_CDA");
            //recuperation du nom de fichier
            string csvFileName = Path.GetFileNameWithoutExtension(csvImportPath);
            string csvDirectory = Path.GetDirectoryName(csvImportPath);
            string csvImportSandDt = csvDirectory + "\\" + csvFileName + "_SANSDT.csv";


            string dataModelstring = Clipper_Param.GetModelCA();


            using (var CahierAffaire = new Clipper_8_Import_OF())
            {

                CahierAffaire.Import(_Context, csvImportPath, dataModelstring, false);

            }
            _Context = null;
        }
        /// <summary>
        /// import du cahier d'affaire sans dt
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sansDTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            Clipper_Param.GetlistParam(_Context);
            string csvImportPath = Clipper_Param.GetPath("IMPORT_CDA");
            //recuperation du nom de fichier
            string csvFileName = Path.GetFileNameWithoutExtension(csvImportPath);
            string csvDirectory = Path.GetDirectoryName(csvImportPath);
            string csvImportSandDt = csvDirectory + "\\" + csvFileName + "_SANSDT.csv";

            string dataModelstring = Clipper_Param.GetModelCA();


            using (var CahierAffaire = new Clipper_8_Import_OF())
            {

                CahierAffaire.Import(_Context, csvImportSandDt, dataModelstring, true);
                //}

            }
            _Context = null;
        }
      
        private void reinitialiserStockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //IEntity TO_CUT_nesting;

            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            var doonaction = new Clipper_8_Before_Nesting_Restore_Event();

            string stage = "_TO_CUT_NESTING";

            //creation du fichier de sortie
            //recupere les path
            Clipper_Param.GetlistParam(_Context);
            IEntitySelector nestingselector = null;

            nestingselector = new EntitySelector();

            //entity type pointe sur la list d'objet du model
            nestingselector.Init(_Context, _Context.Kernel.GetEntityType(stage));
            nestingselector.MultiSelect = true;


            if (nestingselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (IEntity nesting in nestingselector.SelectedEntity)
                {
                    doonaction.Execute(nesting);

                }
            }
            _Context = null;
        }
        //
        /// <summary>
        /// reconstruit les emf dand le stock. et ecrit le nom de l'emf 
        /// cree a partir de celui de la tole  dans le champs filename du stock
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void recreerLEmfToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            List<IEntity> stocks = SimplifiedMethods.Get_Entity_Selector(_Context, "_STOCK");
            if (stocks.Count > 0)
            {
                MessageBox.Show("Seules les Chutes pourront obtenir de nouveaux emf");


                foreach (IEntity currentstockentity in stocks)
                {
                    // 
                    if (currentstockentity.GetFieldValueAsEntity("_SHEET").GetFieldValueAsInt("_TYPE") == 2)
                    {
                        string fieldvalue = currentstockentity.GetFieldValueAsString("FILENAME");
                        //on ne remplace que les chaine filname vides.
                        if (String.IsNullOrEmpty(fieldvalue) == true)
                        {
                            string emffilepath = StockManager.Create_EMF_Of_Stock_Entity(currentstockentity);
                            currentstockentity.SetFieldValue("FILENAME", emffilepath);
                            currentstockentity.Save();
                        }
                    }
                    //
                }

            }
            _Context = null;




        }
        //
        /// <summary>
        /// relance du fichier de retour destiné a clip 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void relanceenvoiecoupeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //IEntity TO_CUT_nesting;
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            var doonaction = new Clipper_8_DoOnAction_AfterSendToWorkshop();
            string stage = "_TO_CUT_NESTING";

            //creation du fichier de sortie
            //recupere les path
            Clipper_Param.GetlistParam(_Context);
            IEntitySelector nestingselector = null;

            nestingselector = new EntitySelector();

            //entity type pointe sur la list d'objet du model
            nestingselector.Init(_Context, _Context.Kernel.GetEntityType(stage));
            nestingselector.MultiSelect = true;


            if (nestingselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (IEntity nesting in nestingselector.SelectedEntity)
                {
                    doonaction.Execute(nesting);

                }
            }
            _Context = null;
        }
       

        /// <summary>
        ///  lance la preparation de la base pour les clients connecté a clipper
        ///  a integrer dans import_tools.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void preparerLaBasePourClipperToolStripMenuItem_Click(object sender, EventArgs e)
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
                string CommandCsPath = Path.GetTempPath() + CustoFileCs.Item1 + ".cs";
                File.WriteAllText(CommandCsPath, CustoFileCs.Item2);

                string cs = CustoFileCs.Item1;
                SimplifiedMethods.NotifyMessage("Updating Database", CustoFileCs.Item1);
                ModelManager modelManager = new ModelManager(modelsRepository);
                //detection des composants odbc 
                activateProgressBar(" Updating "+ CustoFileCs.Item1+" in "+_Context.Connection.DatabaseName);
                modelManager.CustomizeModel(CommandCsPath, Lst_Model.Text, false);
                desactivateProgressBar();

                File.Delete(CommandCsPath);
            }

            SimplifiedMethods.NotifyMessage("Updating Database", "Done");


            ///update parameters
            //chaine de connexion// sp5 //

            //ouverture du context

            //IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;


            //import machineè
            ImportMachine();



            Update_Clipper_Parameters(_Context);
            Update_Clipper_Stock(_Context);
            //
            //mise a jour des champs
           Update_DefaultValues(_Context);
            //
            

            MessageBox.Show("Database "+_Context.Connection.DatabaseName +" Prepared for clipper");

        }

        public void ImportMachine()
        {
            try
            {

                MachineManager machineManager = new MachineManager();
                IModelsRepository modelsRepository = new ModelsRepository();

                

                

                _Context = modelsRepository.GetModelContext(Lst_Model.Text);
                string ZipFileName = null;

                
                activateProgressBar(" Updating Machine in " + _Context.Connection.DatabaseName);
                //
                IEntityList Current_MachineList = _Context.EntityManager.GetEntityList("_CUT_MACHINE_TYPE", "_NAME", ConditionOperator.Equal, "Clipper");
                Current_MachineList.Fill(false);

                //import cam machine
                if (Current_MachineList.Count == 0)
                {
                    ZipFileName = Path.GetTempPath() + "Clipper.zip";
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
                desactivateProgressBar();
            }




            catch { }





        }

        public void Update_Clipper_Parameters(IContext _Context)
        {

            activateProgressBar(" Updating " + "Parameters dans "+ _Context.Connection.DatabaseName);
          
            
            SimplifiedMethods.NotifyMessage("Updating Database", "Updating Clipper Parameters values...");
            ITransaction transaction = _Context.CreateTransaction();
            IParameterValue parameterValue;
            //
            // create directory 

            //update clipper parameters


            //parameterValue = _Context.ParameterSetManager.GetParameterValue("CLIP_CONFIGURATION", "MODEL_CA");
            if (_Context.ParameterSetManager.TryGetParameterValue("CLIP_CONFIGURATION", "MODEL_CA", out parameterValue))
            {
                parameterValue.SetValue(@"0#_NAME#string;1#AFFAIRE#string;2#THICKNESS#string;3#_MATERIAL#string;4#CENTREFRAIS#string;5#TECHNOLOGIE#string;6#FAMILY#string;7#IDLNROUT#string;8#CENTREFRAISSUIV#string;9#_FIRM#string;10#_QUANTITY#integer;11#QUANTITY#double;12#ECOQTY#string;13#STARTDATE#date;14#ENDDATE#date;15#PLAN#string;16#FORMATCLIP#string;17#IDMAT#string;18#IDLNBOM#string;19#NUMMAG#string;20#FILENAME#string;21#_DESCRIPTION#string;22#_CLIENT_ORDER_NUMBER#string;23#DELAI_INT#date;24#EN_RANG#string;25#EN_PERE_PIECE#string;26#ID_PIECE_CFAO#string");
                parameterValue.Save(transaction);
            }

            parameterValue = null;

            if (_Context.ParameterSetManager.TryGetParameterValue("CLIP_CONFIGURATION", "MODEL_DM", out parameterValue))
            {
                //    parameterValue = _Context.ParameterSetManager.GetParameterValue("CLIP_CONFIGURATION", "MODEL_DM");
                parameterValue.SetValue(@"0#_NAME#string;1#_MATERIAL#string;2#_LENGTH#double;3#_WIDTH#double;4#THICKNESS#double;5#QTY_TOT#integer;6#_REST_QUANTITY#integer;7#GISEMENT#string;8#NUMMAG#string;9#NUMMATLOT#string;10#NUMCERTIF#string;11#NUMLOT#string;12#NUMCOUL#string;13#IDCLIP#string;14#FILENAME#string");
                parameterValue.Save(transaction);
            }

            if (_Context.ParameterSetManager.TryGetParameterValue("CLIP_CONFIGURATION", "ACTIVATE_OMISSION", out parameterValue))
            {
                //parameterValue = _Context.ParameterSetManager.GetParameterValue("CLIP_CONFIGURATION", "ACTIVATE_OMISSION");
                parameterValue.SetValue(true);
                parameterValue.Save(transaction);
            }

            if (_Context.ParameterSetManager.TryGetParameterValue("CLIP_CONFIGURATION", "EXPLODE_MULTIPLICITY", out parameterValue))
            {
                //parameterValue = _Context.ParameterSetManager.GetParameterValue("CLIP_CONFIGURATION", "EXPLODE_MULTIPLICITY");
                parameterValue.SetValue(true);
                parameterValue.Save(transaction);
            }


            if (_Context.ParameterSetManager.TryGetParameterValue("_GLOBAL_CONFIGURATION", "_MANAGE_PART_SET", out parameterValue))
            {
                // parameterValue = _Context.ParameterSetManager.GetParameterValue("_GLOBAL_CONFIGURATION", "_MANAGE_PART_SET");
                parameterValue.SetValue(false);
                parameterValue.Save(transaction);
            }

            parameterValue = _Context.ParameterSetManager.GetParameterValue("_GLOBAL_CONFIGURATION", "_MANAGE_STOCK");
            parameterValue.SetValue(true);
            parameterValue.Save(transaction);

            //update export

            parameterValue = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_ACTCUT_OF_PATH");
            parameterValue.SetValue(@"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Devis");
            parameterValue.Save(transaction);

            //
            

            // 5 jours par defaut
            parameterValue = _Context.ParameterSetManager.GetParameterValue("_CONFIGURATION", "_FABRICATION_PERIOD_TEXT");
            parameterValue.SetValue("5");
            parameterValue.Save(transaction);



            //
            parameterValue = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_ACTCUT_DPR_DIRECTORY");
            parameterValue.SetValue(@"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Devis");
            parameterValue.Save(transaction);

            parameterValue = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_ACTCUT_FORCE_DPR_DIRECTORY");
            parameterValue.SetValue(@"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Devis");
            parameterValue.Save(transaction);

            //


            parameterValue = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_EXPORT_GP_MODE");
            parameterValue.SetValue(2);
            parameterValue.Save(transaction);

            parameterValue = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_EXPORT_GP_TYPE");
            parameterValue.SetValue(99);
            parameterValue.Save(transaction);

            parameterValue = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_EXPORT_CFAO_MODE");
            //1 pour oiece par defaut
            //parameterValue.SetValue(1);
            //0 pour pas d'export
            parameterValue.SetValue(0);
            parameterValue.Save(transaction);

            parameterValue = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_CLIPPER_QUOTE_NUMBER_OFFSET");
            parameterValue.SetValue(10000);
            parameterValue.Save(transaction);

            parameterValue = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_GP_EXPORTER_FILENAME");
            parameterValue.SetValue("AF_Export_Devis_Clipper.dll");
            parameterValue.Save(transaction);

            parameterValue = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_GP_EXPORTER_NAMESPACE");
            parameterValue.SetValue("AF_Export_Devis_Clipper");
            parameterValue.Save(transaction);

            parameterValue = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_GP_EXPORTER_CLASS");
            parameterValue.SetValue("CreateTransFile");
            parameterValue.Save(transaction);


            parameterValue = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_EXPORT_GP_DIRECTORY");
            parameterValue.SetValue(@"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Devis");
            parameterValue.Save(transaction);


            ///paramétrage des prix par matiere
            parameterValue = _Context.ParameterSetManager.GetParameterValue("_QUOTE", "_MAT_COST_BY_MATERIAL");
            parameterValue.SetValue(true);
            parameterValue.Save(transaction);


            ///paramétrage des prix au kilos
            parameterValue = _Context.ParameterSetManager.GetParameterValue("_QUOTE", "_SHEET_SPECIFIC_SALE_COST_BY_WEIGHT");
            parameterValue.SetValue(true);
            parameterValue.Save(transaction);

            
            //update part export value

            //ouverture des imports 
            string Import_Client = Properties.Resources.Import_client_1;
            parameterValue = _Context.ParameterSetManager.GetParameterValue("DATA_IMPORT", "IMPORT_DEFINITION");
            parameterValue.SetValue(Import_Client);
            parameterValue.Save(transaction);

            string Import_Matiere = Properties.Resources.Import_Matiere_2;
            parameterValue = _Context.ParameterSetManager.GetParameterValue("DATA_IMPORT_2", "IMPORT_DEFINITION");
            parameterValue.SetValue(Import_Matiere);
            parameterValue.Save(transaction);

            //update prix devis  au kilos







            //afectaion de la machine clipper par defaut sur laser1 par defaut

            string machinelaser1centrefais = "CLIP";
            _Context.ParameterSetManager.TryGetParameterValue("_MACHINE_LASER", "_CENTRE_FRAIS", out parameterValue);
            if (parameterValue.Value == null) { parameterValue.SetValue(machinelaser1centrefais); }
            parameterValue.Save(transaction);

            IEntityList machinelist = _Context.EntityManager.GetEntityList("_CUT_MACHINE_TYPE", "_NAME", ConditionOperator.Equal, "Clipper");
            machinelist.Fill(false);

                                             
            if (machinelist.Count() > 0)
            {
                _Context.ParameterSetManager.TryGetParameterValue("_MACHINE_LASER", "_ALMACAM_MACHINE", out parameterValue);
                IEntity clipperMachine = machinelist.FirstOrDefault();
                parameterValue.SetValue(clipperMachine.Id32.ToString());
                parameterValue.Save(transaction);
            }

            string machinelaser2centrefais = "LASER";
            _Context.ParameterSetManager.TryGetParameterValue("_MACHINE_LASER_2", "_CENTRE_FRAIS", out parameterValue);
            if (parameterValue.Value == null) { parameterValue.SetValue(machinelaser2centrefais); }
            parameterValue.Save(transaction);



            IEntity firm = SimplifiedMethods.CreateEntity_If_Not_Exists(_Context, "_FIRM", "_NAME", "ALMA");
            IEntity CentreFrais = SimplifiedMethods.CreateEntity_If_Not_Exists(_Context, "_CENTRE_FRAIS", "_CODE", "Clip");

            //
            //ajout des fichiers json

            string CommandJsPath = Directory.GetCurrentDirectory() + Properties.Resources.AF_import_clipper;
            File.WriteAllText(CommandJsPath, "AF_import_clipper.json");

            _Context = null;
            transaction = null;


            //creation des entités par defaut
            SimplifiedMethods.NotifyMessage("Updating Defautl data", "Updating Parameters DONE...");

            desactivateProgressBar();

        }

        public void Update_Clipper_Stock(IContext _Context)
        {

            SimplifiedMethods.NotifyMessage("Almacam_Clipper", " mise à jour des entités de stock de stock pour clipper.");
            activateProgressBar("Updating Stock on " + _Context.Connection.DatabaseName);

            //set value 
            IEntityList stocks = _Context.EntityManager.GetEntityList("_STOCK");
            stocks.Fill(false);


            SimplifiedMethods.NotifyMessage("Almacam_Clipper", "traitement de " + stocks.Count() + " entités de stock.");

            Cursor.Current = Cursors.WaitCursor;

            long position = 0;
            int step = 1;
            foreach (IEntity stock in stocks)
            {
                position++;

                SimplifiedMethods.NotifyStatusMessage("", " Initialisation   en cours ...", stocks.Count(), position, ref step, 0);



                //
                //ajustement des quantités negatives

                long dispoqty = stock.GetFieldValueAsLong("_REST_QUANTITY");

                //if (dispoqty < 0 || dispoqty == null || dispoqty == 0)
                if (dispoqty < 0  || dispoqty == 0)
                {
                    stock.SetFieldValue("_REST_QUANTITY", 0);

                }




                //les données seront settées par clipper


                stock.SetFieldValue("AF_IS_OMMITED", false);
                if (stock.GetFieldValueAsString("IDCLIP") == string.Empty)

                {// non ttraite
                    stock.SetFieldValue("IDCLIP", null);


                }
                //recuperation du champs filename de la tole dans le champs filename du stock
                ///recupération du filename des sheet dans le stock
                string filname = stock.GetFieldValueAsEntity("_SHEET").GetFieldValueAsString("FILENAME");

                if (filname != string.Empty || filname != null)
                {
                    stock.SetFieldValue("FILENAME", filname);
                    //Creating_emf;
                    if (File.Exists(filname) == false)
                    {
                        //create emf //

                    }
                }










                stock.Save();

                desactivateProgressBar();
            }
            SimplifiedMethods.NotifyMessage("Almacam_Clipper", "Done.");


            stocks = null;

        }

        public void Update_DefaultValues(IContext _Context)
        {

            SimplifiedMethods.NotifyMessage("Almacam_Clipper", "Definition des valeurs par defaut.");
            activateProgressBar("Updating default values from "+_Context.Connection.DatabaseName);

            IEntity ClipCFEntity = null;
            IEntity SuperUserEntity = null;

            IEntityList userliste = _Context.EntityManager.GetEntityList("SYS_USER", "USER_NAME", ConditionOperator.Equal, "SUPER");
            userliste.Fill(false);
            if (userliste.Count() != 0)
            {
                SuperUserEntity = userliste.FirstOrDefault();
            }
            else
            {
                SuperUserEntity = _Context.EntityManager.CreateEntity("SYS_USER");
                SuperUserEntity.SetFieldValue("USER_NAME", "SUPER");
                SuperUserEntity.SetFieldValue("ROLES", "_ADMINISTRATOR");
                SuperUserEntity.Save();

            }

            //            

            IEntityList centrefraisliste = _Context.EntityManager.GetEntityList("_CENTRE_FRAIS", "_CODE", ConditionOperator.Equal, "CLIP");
            centrefraisliste.Fill(false);
            if (centrefraisliste.Count() > 0)
            {
                ClipCFEntity = centrefraisliste.FirstOrDefault();
            }
            else
            {
                ClipCFEntity = _Context.EntityManager.CreateEntity("_CENTRE_FRAIS");
                ClipCFEntity.SetFieldValue("_NAME", "CLIP");
                ClipCFEntity.SetFieldValue("_CODE", "CLIP");
                ClipCFEntity.Save();

            }

            //machine clipper
            IEntityList machineList = _Context.EntityManager.GetEntityList("_CUT_MACHINE_TYPE", "_NAME", ConditionOperator.Equal, "Clipper");
            machineList.Fill(false);



            if (machineList.Count() > 0)
            {// non ttraite
                //machineList.FirstOrDefault.try
                //machineList.FirstOrDefault().SetFieldValue()
                IField v;
                if(machineList.FirstOrDefault().EntityType.FieldList.TryGetValue("CENTREFRAIS_MACHINE",out v )==true)
                machineList.FirstOrDefault().SetFieldValue("CENTREFRAIS_MACHINE", ClipCFEntity);
            }

            machineList.FirstOrDefault().Save();


            //société emettrice


            IEntity Societe_Emet_Entity = null;

            IEntityList Societe_Emet_EntityList = _Context.EntityManager.GetEntityList("_MANUFACTURER");
            Societe_Emet_EntityList.Fill(false);

            if (Societe_Emet_EntityList.Count() == 0)
            {
                Societe_Emet_Entity = Societe_Emet_EntityList.FirstOrDefault();

                Societe_Emet_Entity = _Context.EntityManager.CreateEntity("_MANUFACTURER");
                Societe_Emet_Entity.SetFieldValue("_NAME", "AVENIO");
                Societe_Emet_Entity.SetFieldValue("_ADDRESS", "Rue du college de la croix");
                Societe_Emet_Entity.SetFieldValue("_POSTCODE", "84000");
                Societe_Emet_Entity.SetFieldValue("_CITY", "AVIGNON");
                Societe_Emet_Entity.Save();

            }





            SimplifiedMethods.NotifyMessage("Machine Updated", "Done.");


            machineList = null;
            centrefraisliste = null;
            ClipCFEntity = null;
            SuperUserEntity = null;

            ///


            desactivateProgressBar();

            ///


        }
       
        private void initialiserStockToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                //creation du model repository
                IModelsRepository modelsRepository = new ModelsRepository();
                IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;

                Update_Clipper_Stock(_Context);

                _Context = null;
                MessageBox.Show("Mise à jour du stock terminé");


            }

            finally
            {
                _Context = null;
                Cursor.Current = Cursors.Default;
            }

        }

        private void activateProgressBar(string text)
        {
            progresslabel.Visible = true;
            progresslabel.Text = text;
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            progressBar1.MarqueeAnimationSpeed = 30;
        }

        private void desactivateProgressBar()
        {
            progresslabel.Visible = false;
            progresslabel.Text = "";
            progressBar1.Visible = false;
            progressBar1.Style = ProgressBarStyle.Continuous;
            progressBar1.MarqueeAnimationSpeed = 0;
        }

        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            AboutBox1 a = new AboutBox1();
            a.Show();
        }

        /// <summary>
        /// exporte les données techniques de pieces 2d pour clipper
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void remonteeDTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //creation du model repository
            IModelsRepository modelsRepository = new ModelsRepository();
            IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;

            var Export_dt = new AF_Clipper_Dll.Clipper_8_Export_DT_Processor();
            Export_dt.Execute(_Context);

            _Context = null;
        }
        /// <summary>
        /// ajoute les evenements clipper
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ajouterLesEvenementsToolStripMenuItem_Click(object sender, EventArgs e)
        {
        

            IModelsRepository modelsRepository = new ModelsRepository();

            List<Tuple<string, string>> CustoFileCsList = new List<Tuple<string, string>>();


            CustoFileCsList.Add(new Tuple<string, string>("FormulasAndEvents", Properties.Resources.Event_Clip));


            foreach (Tuple<string, string> CustoFileCs in CustoFileCsList)
            {
                string CommandCsPath = Path.GetTempPath() + CustoFileCs.Item1 + ".cs";
                File.WriteAllText(CommandCsPath, CustoFileCs.Item2);

                string cs = CustoFileCs.Item1;

                ModelManager modelManager = new ModelManager(modelsRepository);
                modelManager.CustomizeModel(CommandCsPath, Lst_Model.Text, true);
                File.Delete(CommandCsPath);
            }
        }
        /// <summary>
        /// ajoute les commandes spécifiques clipper
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ajouterLesCommandesToolStripMenuItem_Click(object sender, EventArgs e)
        {


            IModelsRepository modelsRepository = new ModelsRepository();

            List<Tuple<string, string>> CustoFileCsList = new List<Tuple<string, string>>();


            CustoFileCsList.Add(new Tuple<string, string>("Commandes", Properties.Resources.Commandes_Clip));

            foreach (Tuple<string, string> CustoFileCs in CustoFileCsList)
            {
                string CommandCsPath = Path.GetTempPath() + CustoFileCs.Item1 + ".cs";
                File.WriteAllText(CommandCsPath, CustoFileCs.Item2);

                string cs = CustoFileCs.Item1;

                ModelManager modelManager = new ModelManager(modelsRepository);
                modelManager.CustomizeModel(CommandCsPath, Lst_Model.Text, true);
                File.Delete(CommandCsPath);
            }
        }
        /// <summary>
        /// ajoute les champs spécifiques clipper
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ajouerLesChampsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

            IModelsRepository modelsRepository = new ModelsRepository();

            List<Tuple<string, string>> CustoFileCsList = new List<Tuple<string, string>>();

            CustoFileCsList.Add(new Tuple<string, string>("Entities", Properties.Resources.Entities_Clip));


            foreach (Tuple<string, string> CustoFileCs in CustoFileCsList)
            {
                string CommandCsPath = Path.GetTempPath() + CustoFileCs.Item1 + ".cs";
                File.WriteAllText(CommandCsPath, CustoFileCs.Item2);

                string cs = CustoFileCs.Item1;

                ModelManager modelManager = new ModelManager(modelsRepository);
                modelManager.CustomizeModel(CommandCsPath, Lst_Model.Text, true);
                File.Delete(CommandCsPath);
            }
        }
        
        /// <summary>
        /// lance les correction de migraiton de base almacam
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void updateToSpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;
            //SimplifiedMethods.NotifyMessage("Updating Database", "Updating Parameters...");
            Update_Clipper_Parameters(_Context);
            SimplifiedMethods.NotifyMessage("Updating Database", "Done");
        }

        /// <summary>
        /// import la machine clipper
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void importClipperMachineToolStripMenuItem_Click(object sender, EventArgs e)
        {

            ImportMachine();

        }
        /// <summary>
        /// insert les valeurs par defaut de certains champs d'entités dans la base almacam
        /// notament le stock
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void addDefaultValuesForTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);
            Update_DefaultValues(_Context);
        }
        /// <summary>
        /// ouvre le fichier d'importca evc des entetes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void oFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);
            Clipper_Param.GetlistParam(_Context);
            string cda = Clipper_Param.GetModelCA();
            debugTool.OpenWithHeader(Clipper_Param.GetPath("IMPORT_CDA") ,cda);
        }
        /// <summary>
        /// ouvre le fichier de stcok avec les enteetes programmées
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void stockToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);
            Clipper_Param.GetlistParam(_Context);
            string modeldm = Clipper_Param.GetModelDM();

            debugTool.OpenWithHeader(Clipper_Param.GetPath("IMPORT_DM") , modeldm);
        }
        /// <summary>
        /// annulation du retour gp ( supprime le fichier de retour et les chutes)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void anullerRetourGpToolStripMenuItem_Click(object sender, EventArgs e)
        {

            IModelsRepository modelsRepository = new ModelsRepository();
            IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;

            //string stage = "_TO_CUT_SHEET";
            string stage = "_TO_CUT_NESTING";
            //creation du fichier de sortie
            //recupere les path
            Clipper_Param.GetlistParam(_Context);
            IEntitySelector Entityselector = null;

            Entityselector = new EntitySelector();

            //entity type pointe sur la list d'objet du model
            Entityselector.Init(_Context, _Context.Kernel.GetEntityType(stage));
            Entityselector.MultiSelect = false;

            if (Entityselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                /*suppression des fichier de retour gp sur annulation*/
                StockManager.DeleteNestingAssociatedStock(Entityselector.SelectedEntity.FirstOrDefault());
                

       
            }

            _Context = null;
            _Context = null;

            
        }      
      
        // lance l'executable pour l'export des devis 
        private void exportDevisToolStripMenuItem1_Click(object sender, EventArgs e)
        {// en dur pour le moment --> voir pour recupere le path dans clipper_param APPLICATION_1
            Process.Start("C:\\AlmaCam\\Bin\\AF_Export_Devis.exe");
        }
        /// <summary>
        /// retourne un message ok si la connection odbc  paramétré et fonctionnelle
        /// recuper le nom du dns dans le fichier json
        /// ourvre une connection puis la referme ou lance un message d'erreur
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void testerODBCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool rst = AF_Import_ODBC_Clipper_AlmaCam.OdbcTools.Test_Connexion();
        }
        // à virer
        private void exporterFicheToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        #endregion
    }

    /*
    public class IImportMachineEntity : IMachineImporter
    {
        public IList<QualitySetInfo> QualitySetInfoList { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public IList<QualityInfo> QualityInfoList { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public IList<MaterialSetInfo> MaterialSetInfoList { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public IList<MaterialInfo> MaterialInfoList { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public IEntity ImportCutMachine(bool InBackground)
        {
            throw new NotImplementedException();
        }

        public void Init(IContext Context)
        {
            throw new NotImplementedException();
        }

        public void ReadZipFile(string ZipFileName)
        {
            throw new NotImplementedException();
        }

        public IEntity UpdateCutMachine(IEntity CutMachineEntity, bool InBackground)
        {
            throw new NotImplementedException();
        }
    }
    */

    //classe fields
    [Obsolete("Reporté dans les fichirs cs d'import des champs(voir ressources)")]
    public class NewField
    {
        public string Key;
        public string Name;
        public FieldDescriptionEditableType Editable;
        public FieldDescriptionVisibilityType Visible;
        public FieldDescriptionType Type;
        public string Link;
        public NewField(string key, string name, FieldDescriptionEditableType editable, FieldDescriptionVisibilityType visible, FieldDescriptionType type, string link)
        {
            Key = key;
            Name = name;
            Editable = editable;
            Visible = visible;
            Type = type;
            Link = link;
        }
    }


    public class InstallclipperTest : ScriptModelCustomization, IScriptModelCustomization
    {
        IContext _Context;
        
        IContext _HostContext ;

        public InstallclipperTest(IContext context){
          
            IModelsRepository modelsRepository = new ModelsRepository();
            _HostContext = modelsRepository.GetModelContext("ModelsRepository");  //nom de la base;
            _Context = context;

        }
        
            public override bool Execute(IContext _Context, IContext hostContext)
            {


                #region Stock

                {
                    IEntityType entityType = _Context.Kernel.GetEntityType("_STOCK");
                    IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(_Context, 1, entityType, null, "_SUPPLY", null);
                    entityTypeFactory.Key = "_STOCK";
                    entityTypeFactory.Name = "Stock";
                    entityTypeFactory.DefaultDisplayKey = "_NAME";
                    entityTypeFactory.ActAsEnvironment = false;

                    //if(entityType.FieldList)

                    {
                        IFieldDescription fieldDescription = new FieldDescription(_Context.Kernel.UnitSystem, true);
                        fieldDescription.Key = "FILENAME";
                        fieldDescription.Name = "*Renommé ?";
                        fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                        fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                        fieldDescription.Mandatory = false;
                        fieldDescription.FieldDescriptionType = FieldDescriptionType.Boolean;
                        fieldDescription.DefaultValue = false;
                        entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);

                    }

                    if (!entityTypeFactory.UpdateModel())
                    {
                        foreach (ModelSettingError error in entityTypeFactory.ErrorList)
                        {
                            hostContext.TraceLogger.TraceError(error.Message, true);
                        }
                        return false;
                    }

                }

                #endregion
                return true;
            }
        

    }

    public  class debugTool
    {
        public static void OpenWithHeader(string defaultfile,string header)
        {
            string CsvImportPath = "";
            //string DataModelString = "";
            string fileName = "";

            DialogResult dialogResult = MessageBox.Show("Voulez vous selectionner le fichier ?", "debug", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Title = defaultfile;
                    openFileDialog.InitialDirectory =Path.GetDirectoryName( defaultfile);
                    openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 2;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Get the path of specified file
                        CsvImportPath = openFileDialog.FileName;

                    }
                }


            }
            else
            {
                // ajout entete //

                CsvImportPath = defaultfile;
            }


            if (File.Exists(CsvImportPath))
                fileName = Path.GetDirectoryName(CsvImportPath) + @"\" + "DEBUG_" + Path.GetFileName(CsvImportPath);

            if (File.Exists(fileName))
                File.Delete(fileName);

            if (File.Exists(CsvImportPath))
                File.Copy(CsvImportPath, fileName, true);

            if (File.Exists(fileName)) { 
            var txtLines = File.ReadAllLines(fileName).ToList();   //Fill a list with the lines from the txt file.
            txtLines.Insert(0, header);  //Insert the line you want to add last under the tag 'item1'.
            File.WriteAllLines(fileName, txtLines);
            Process.Start(@fileName);
            }





        }
    }
}











