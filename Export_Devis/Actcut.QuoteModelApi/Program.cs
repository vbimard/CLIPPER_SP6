
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
using Actcut.QuoteModelManager;
using Actcut.ActcutModelManagerUI;

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
            GetParam(arguments, out action, out user, out db, out quoteNumber, out orderNumber, out exportFile);



            //on force a l'export
            action = "ExportQuote";

            if (exportFile == null)
            {
                action = "ExportQuote";
                user = "SUPER";
                //db = "AlmaCAM_Clipper_7";
                db = "AlmaCAM";
                //id almacam
                quoteNumber = 10009;
                orderNumber = "alma";
                exportFile = @"c:\Temp";



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

                    Console.WriteLine("conecting database...");

                    contextlocal = Context;//

                    Console.WriteLine("conecting database...");
                   
                   

                    //  //

                        IQuoteManagerUI quoteManagerUI = new QuoteManagerUI();
                        IEntity quoteEntity = quoteManagerUI.GetQuoteEntity(contextlocal, quoteNumber);                 
                        ret = quoteManagerUI.AccepQuote(contextlocal, quoteEntity, orderNumber, exportFile);


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
                    long quoteId;
                    //bool status = AF_Export_Devis_Clipper.GetQuote(out quoteId);
                   // Environment.ExitCode = (int)quoteId;
                }
                else
                {
                    Environment.ExitCode = -2;
                }
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

            if (arguments.Count()==6)
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
                        if (values.Length == 2)
                         //   quoteNumber= values[1].Trim(); //= //
                        Convert.ToInt64(values[1].Trim());
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
                }
            }
            }
            else
            {


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

    }
}

