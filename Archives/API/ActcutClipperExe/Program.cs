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


using AF_Actcut.ActcutClipperApi;

namespace Actcut.ActcutClipperExe
{
    class Program
    {
        private static ClipperApi ClipperApi = null;
        private static IContext Context = null;

        private static string AlmaCamBinFolder = null;
        private static string ParamFolder = null;
        private static string ResultFile = null;
        private static string ParamFile = null;
        private static string WorkFile = null;

        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            AlmaCamBinFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            ParamFolder = AlmaCamBinFolder + @"\ActcutClipperExeParam"; // ActcutClipperExeParam : nom de sous répertoire imposé
            ParamFile = ParamFolder + @"\Param.txt"; // Nom du fichier contenant les commandes : nom imposé
            WorkFile = ParamFolder + @"\Param.wrk"; // Nom du fichier contenant les commandes : nom imposé
            ResultFile = ParamFolder + @"\Result.txt"; // Nom du fichier contenant les resultats : nom imposé

            if (Directory.Exists(ParamFolder) == false) Directory.CreateDirectory(ParamFolder);

            string[] arguments = Environment.GetCommandLineArgs();

            string action; string user; string db;
            long quoteNumber; string orderNumber; string exportFile;
            GetParam(arguments, out action, out user, out db, out quoteNumber, out orderNumber, out exportFile);

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
                    bool status = ClipperApi.ExportQuote(quoteNumber, orderNumber, exportFile);
                    Environment.ExitCode = (status ? 0 : -1);
                }
                else
                {
                    Environment.ExitCode = -2;
                }
            }
            #endregion

            #region Get Quote

            else if (action == "SelectQuoteUI" && db != null)
            {
                bool ret = Init(db, user);
                if (ret)
                {
                    long quoteId;
                    bool status = ClipperApi.SelectQuoteUI(out quoteId);
                    Environment.ExitCode = (int)quoteId;
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

            if (ClipperApi == null)
            {
                ClipperApi = new ClipperApi();
                ClipperApi.InitAlmaCam(Context);
            }

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

                            if (action == "SelectQuoteUI")
                            {
                                ResultFile = Path.Combine(ParamFolder, ResultFile);
                                if (File.Exists(ResultFile)) File.Delete(ResultFile);

                                long quoteId;
                                bool ret = ClipperApi.SelectQuoteUI(out quoteId);
                                File.WriteAllText(ResultFile, quoteId.ToString());
                            }

                            if (action == "ExportQuote")
                            {
                                ResultFile = Path.Combine(ParamFolder, ResultFile);
                                if (File.Exists(ResultFile)) File.Delete(ResultFile);

                                bool ret = ClipperApi.ExportQuote(quoteNumber, orderNumber, exportFile);
                                File.WriteAllText(ResultFile, ret ? "0" : "-1");
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
                        quoteNumber = Convert.ToInt64(values[1].Trim());
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

