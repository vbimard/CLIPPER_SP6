
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Reflection;

using Alma.BaseUI.Utils;

using Wpm.Schema.Kernel;
using Wpm.Implement.Manager;
using Wpm.Implement.ComponentEditor;
using AF_Export_Devis_Clipper;
using Actcut.ActcutModelManagerUI;
using System.Text.RegularExpressions;
using Alma.BaseUI.ErrorMessage;

namespace AF_Export_Devis_Clipper
{
    class Program
    {
        //private static ClipperApi ClipperApi = null;
        private static IContext Context = null;
        private static string AlmaCamBinFolder = null;
        private static string ParamFolder = null;
        private static string ResultFile = null;
        private static string ParamFile = null;
        private static string WorkFile = null;
        //IContext contextlocal;
        //
        /*exemple argument
         Action=ExportQuote Db="AlmaCAM_Clipper" QuoteNumber="1" OrderNumber="alma" ExportFile="c:\temps"

         Action=GetQuote Db="AlmaCAM_Clipper" QuoteNumber="-1"
             
             */
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            AlmaCamBinFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            ParamFolder = AlmaCamBinFolder + @"\ActcutClipperExeParam"; // ActcutClipperExeParam : nom de sous répertoire imposé
            ParamFile = ParamFolder + @"\Param.txt";                    // Nom du fichier contenant les commandes : nom imposé
            WorkFile = ParamFolder + @"\Param.wrk";                     // Nom du fichier contenant les commandes : nom imposé
            ResultFile = ParamFolder + @"\Result.txt";                  // Nom du fichier contenant les resultats : nom imposé

            if (Directory.Exists(ParamFolder) == false) Directory.CreateDirectory(ParamFolder);
            string[] arguments = Environment.GetCommandLineArgs();

            

            string action; string user; string db;
            long quoteNumber; string orderNumber; string exportFile;
            //on force a l'export
            action = "ExportQuote";
            GetParam(arguments, out action, out user, out db, out quoteNumber, out orderNumber, out exportFile);

            #if DEBUG
            string arg = "";
            foreach (string s in arguments)
            { arg += s + " "; }
            MessageBox.Show(arg);
            MessageBox.Show(action +" user "+ user + " database "+ db + " quote number "+ quoteNumber + " order "+ orderNumber + " export "+ exportFile + " "  );
#endif

            action=arguments[1].Split('=')[1];

            if (exportFile == null && action == "ExportQuote")
            {
                
                //action = "ExportQuote";
                user = "SUPER";
                //db = "AlmaCAM_Clipper_7";
                db = "AlmaCAM_Clipper";
                //id almacam
                //quoteNumber = 10009;
                string input;
                input = "-1";
                //input = "1000-100";
                //input = "10000";
                if (ContainsAlpha(input) == false)
                    quoteNumber =Convert.ToInt64(input.Trim());
                //quoteNumber = -1;
                orderNumber = "alma";
                exportFile = "";// @"c:\Temps\TEST.TXT";

            }
            else
            {

                action = arguments[1].Split('=')[1];
                db = arguments[2].Split('=')[1]; ;
                orderNumber = arguments[3].Split('=')[1]; ;


            }

               

            #region Init

            if (action == "Init" && db != null)
            {
                bool ret = Init(db, user);
                if (ret)
                {
                    File.WriteAllText(ResultFile, "Ok");
                    ManageCommand();
                }
                else
                {
                    File.WriteAllText(ResultFile, "Error");
                    Environment.ExitCode = -1;
                }
            }

            #endregion

            #region Export Quote

            else if (action == "ExportQuote" && db != null)
            {
                bool ret = Init(db, user);
                if (ret)
                {
                    IContext contextlocal;
                    IEntity quote = null;
                    IEntityList quotelist = null;
                    //bool return;

                    string database;
                    database = db;

                    Console.WriteLine("connecting database...");

                    contextlocal = Context;//

                    Console.WriteLine("connecting database...");



                    //  //

                    IQuoteManagerUI quoteManagerUI = new QuoteManagerUI();
                   
                    IEntity quoteEntity = quoteManagerUI.GetQuoteEntity(contextlocal, quoteNumber);
                    if (string.IsNullOrEmpty(exportFile)){
                        IParameterValue iparametervalue;
                        bool rst = contextlocal.ParameterSetManager.TryGetParameterValue("_EXPORT", "_EXPORT_GP_DIRECTORY", out iparametervalue);
                        exportFile = iparametervalue.GetValueAsString()+ "\\Trans_" + quoteEntity.GetFieldValueAsString("_REFERENCE")+".txt";
                    }


                    if (Directory.Exists(Path.GetDirectoryName(exportFile)) == false)
                    {
                        MessageBox.Show("Chemin de sortie du devis non existant, voir les parametres de clipper");
                        Environment.Exit(0);
                    }

                    //IEntity quoteEntity = quoteManagerUI.GetQuoteEntity(contextlocal, devisid);
                    //ret = quoteManagerUI.AccepQuote(contextlocal, quoteEntity, orderNumber, exportFile);
                    ret = quoteManagerUI.AccepQuote(contextlocal, quoteEntity, orderNumber, exportFile, out IErrorMessageList ErrorMessageList);


                    // bool AccepQuote(IContext Context, IEntity QuoteEntity, string OrderNumber, string GpExportFileName);
                    // bool AccepQuote(IContext Context, IEntity QuoteEntity, string OrderNumber, string GpExportFileName, out IErrorMessageList ErrorMessageList);




                    Environment.ExitCode = (ret ? 0 : -1);
                }
                else
                {
                    Environment.ExitCode = -2;
                }
            }
            #endregion

            #region Get Quote

            else if (action == "GetQuote" && db != null)
            {

                bool ret = Init(db, user);
                if (ret)
                {
                    IContext contextlocal;
                    IEntity quote = null;
                    IEntityList quotelist = null;
                    //bool return;

                    string database;
                    database = db;

                    Console.WriteLine("connecting database...");

                    contextlocal = Context;//

                    Console.WriteLine("connecting database...");
                    
                    IQuoteManagerUI quoteManagerUI = new QuoteManagerUI();

                    IEntity quoteEntity = quoteManagerUI.GetQuoteEntity(contextlocal, quoteNumber);

                    

                    //ret = quoteManagerUI.AccepQuote(contextlocal, quoteEntity, orderNumber, exportFile, out IErrorMessageList ErrorMessageList);
                    Environment.ExitCode = (int)quoteEntity.Id;
                    }
                 else
                    {   
                    Environment.ExitCode = -2;
                 }   
                

                    /*
                    bool ret = Init(db, user);
                    if (ret)
                    {
                        long quoteId;
                        //bool status = AF_Export_Devis_Clipper.GetQuote(out quoteId);
                        // Environment.ExitCode = (int)quoteId;
                    }
                    else
                    {
                        Environment.ExitCode = -2;
                    }*/
                }

            #endregion        }
        }

        static private bool Init(string db, string user)
        {
            if (Context == null)
            {
                IModelsRepository modelsRepository = new ModelsRepository();
                Context = modelsRepository.GetModelContext(db);
                if (Context != null)
                {
                    Licence.InitLicence(Context.Kernel, null);
                    SetUser(Context, user);
                }
                else
                {
                    //recuperation de la derniere base ouverte
                    MessageBox.Show("la base demandée par Clipper n'est pas detectée - > nom de la base demandée :" + db +", merci de revoir le nom de la base configurée dans les profiles clipper .  "+"\r\n", db);
                    return false;
                }
            }


            /*

            if (CreateTransFile == null)
            {
                CreateTransFile = new CreateTransFile();
                CreateTransFile.Init(Context);
            }
            */
            return true;
        }

        static private void ManageCommand()
        {
            bool exit = false;

            while (!exit)
            {
                Thread.Sleep(10);

                if (File.Exists(ParamFile))
                {
                    try
                    {
                        if (File.Exists(WorkFile)) File.Delete(WorkFile);

                        File.Move(ParamFile, WorkFile);
                        IList<string> commandLineList = new List<string>(File.ReadLines(WorkFile));
                        string commandLine = commandLineList.FirstOrDefault();
                        if (commandLine != null)
                        {
                            string[] paramList = commandLine.Split(new Char[] { ' ' });

                            string action; string user; string db;
                            long quoteNumber; string orderNumber; string exportFile;
                            GetParam(paramList, out action, out user, out db, out quoteNumber, out orderNumber, out exportFile);

                            if (action == "Exit")
                            {
                                Environment.Exit(0);
                            }

                            if (action == "GetQuote")
                            {
                                ResultFile = Path.Combine(ParamFolder, ResultFile);
                                if (File.Exists(ResultFile)) File.Delete(ResultFile);

                                long quoteId;
                                //bool ret = ClipperApi.GetQuote(out quoteId);
                                //File.WriteAllText(ResultFile, quoteId.ToString());
                            }

                            if (action == "ExportQuote")
                            {
                                ResultFile = Path.Combine(ParamFolder, ResultFile);
                                if (File.Exists(ResultFile)) File.Delete(ResultFile);

                                ///bool ret = ClipperApi.ExportQuote(quoteNumber, orderNumber, exportFile);
                                ///File.WriteAllText(ResultFile, ret ? "0" : "-1");
                            }
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        if (File.Exists(WorkFile)) File.Delete(WorkFile);
                    }
                }
            }
        }

        static private void GetParam(string[] arguments, out string action, out string user, out string db, out long quoteNumber, out string orderNumber, out string exportFile)
        {
            action = null;
            user = null;
            db = null;
            quoteNumber = -1;
            orderNumber = null;
            exportFile = null;
            //log chaine de connexion
            
            string argString = "";
            if (arguments.Count() == 6 )
            {
                foreach (string argument in arguments)
                {
                    if (argument.StartsWith("Action", StringComparison.InvariantCultureIgnoreCase))
                    {
                        string[] values = argument.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                        if (values.Length == 2)
                            action = values[1].Trim();
                    }

                    if (argument.StartsWith("User", StringComparison.InvariantCultureIgnoreCase))
                    {
                        string[] values = argument.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                        if (values.Length == 2)
                            user = values[1].Trim();
                    }

                    if (argument.StartsWith("Db", StringComparison.InvariantCultureIgnoreCase))
                    {
                        string[] values = argument.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                        if (values.Length == 2)
                            db = values[1].Trim();
                    }

                    if (argument.StartsWith("QuoteNumber", StringComparison.InvariantCultureIgnoreCase))
                    {
                        string[] values = argument.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                        string val;
                        if (values.Length == 2)
                            val= values[1].Trim();
                        else
                        {
                            val = "-1";
                        }

                        if (ContainsAlpha(values[1].Trim()) == false) { 
                            quoteNumber = Convert.ToInt64(values[1].Trim());
                        }
                        else
                        {//pas de texte dans cet argument
                            Environment.Exit(0);
                        }

                        argString += "-" + val;
                      }

                    if (argument.StartsWith("OrderNumber", StringComparison.InvariantCultureIgnoreCase))
                    {
                        string[] values = argument.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                        if (values.Length == 2)
                            orderNumber = values[1].Trim();
                    }

                    if (argument.StartsWith("ExportFile", StringComparison.InvariantCultureIgnoreCase))
                    {
                        string[] values = argument.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                        if (values.Length == 2)
                            exportFile = values[1].Trim();
                                //premiere verification
                            if (Directory.Exists(Path.GetDirectoryName(exportFile))==false){
                                MessageBox.Show("Chemin de sortie du devis non existant, voir les parametres de clipper");
                                                Environment.Exit(0);
                                }
                             
                    }



                }
            }
            else
            {
                //MessageBox.Show("Arguments manquants pour la recupération du devis dans almacam. Contacter Clipper");

            }

           
            using (System.IO.StreamWriter file =new System.IO.StreamWriter(Path.GetTempPath() + "\\argClipper.txt"))
            {
                foreach (string argument in arguments)
                {

                    string[] values = argument.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string str in values)
                    {
                        argString +=" - "+ str.Trim();
                    }
                       


                }


                    file.WriteLine(argString);
                
            }





        }
        static private bool SetUser(IContext context, string user)
        {
            IEntityList userEntityList = context.EntityManager.GetEntityList("SYS_USER", "USER_NAME", ConditionOperator.Equal, user);
            userEntityList.Fill(false);

            IEntity userEntity = userEntityList.FirstOrDefault();
            if (userEntity != null)
            {
                context.UserId = userEntity.Id;
                return true;
            }

            return false;
        }



        static private bool ContainsAlpha (string input)
            {
            
             // The regular expression we use to match
                    Regex r1 = new Regex(@"^[a-zA-Z\s,]*$"); //[\t\0x0020] tab and spaces.
                    string test=input.Replace("-", "z");
                    Match match = r1.Match(test);
                    if (match.Success)
                    {
                        
                        MessageBox.Show("Quote Number Reference contains Alpha Caracteres or - : " + input);
                        return true;

                    }
                    else
                    {
                        return false;
                    }

            }





        


    }
}

