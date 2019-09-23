
//using Clipper_Dll;
//
using AF_ImportTools;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Drawing;

//almacam
using Wpm.Implement.ComponentEditor;  // ouverture de fenetres de selection almacam
using Wpm.Implement.Manager;
using Wpm.Implement.Processor;
using Wpm.Schema.Kernel;
//actcut
using Actcut.ActcutModelManager;
using Actcut.NestingManager;
using Actcut.ResourceManager;
using Actcut.ResourceModel;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Diagnostics;
using Alma.BaseUI.Utils;
using Actcut.CommonModel;
using System.Globalization;
using Wpm.Implement.ModelSetting;
using System.Collections;
using Alma.BaseUI.DescriptionEditor;
//




namespace AF_Clipper_Dll
//namespace Import_GP
#region commande_processor

{
    /// <summary>
    /// automatisme BO : outils necessaire pour l'envoie d'infos au service windoxs
    /// </summary>
    /// 
    ///automation
    public class Automation_Tools : IDisposable

    {

        public void Dispose()
        {
            //Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// OBLIGATION D EXECUTER EN ADMIN
        /// recupere la liste des placements exportes (confition exported)
        /// recupere la liste des fichier dans les dossiers gpao
        /// cloture si le fichier n'est plus present
        /// ce code utilise les log windwos: pour supprimer les log windows
        /// Get-EventLog -List
        /// C:\> Remove-EventLog -LogName "MyLog"
        /// Remove-EventLog -Source "MyApp"
        /// </summary>
        /// <returns>true/false</returns>
        public bool Automatic_Nestings_Close(IContext Contextlocal, string Stage, AF_ImportTools.WindowsLog log)
        {
            string message = ""; //message de suivi pour les log
            try
            {

                Clipper_Param.GetlistParam(Contextlocal);
                bool rst = false;
                //string stage = "_TO_CUT_NESTING";
                IEntityList nestings_list = null;
                IEntity current_nesting = null;
                System.Console.WriteLine("connecté à" + Contextlocal.Connection.DatabaseName + " Pour cloture");
                System.Console.WriteLine("fermeture des placements " + Stage);

                nestings_list = Contextlocal.EntityManager.GetEntityList(Stage, AF_ImportTools.SimplifiedMethods.Get_Marqued_FieldName(Stage), ConditionOperator.Equal, true);
                nestings_list.Fill(false);


                if (nestings_list != null && nestings_list.Count() > 0)
                {
                    foreach (IEntity nesting in nestings_list)
                    {
                        string nesting_name;
                        string technology;
                        string path_to_file;
                        //string msgstart;
                        //On initialise le message
                        message = "";
                        //IEntityList nestings_to_close;
                        IEntity nesting_to_close;


                        nesting_to_close = nesting;

                        //get the nesting name 
                        nesting_name = current_nesting.GetFieldValueAsString("_NAME");
                        //get the technology--> get the folder
                        System.Console.WriteLine("----------------------------------------");
                        System.Console.WriteLine("Placement: " + nesting_name);
                        message += "Clôture de: " + nesting_name;
                        technology = AF_ImportTools.Machine_Info.GetNestingTechnologyName(ref Contextlocal, ref current_nesting);
                        System.Console.WriteLine(nesting_name + ": Technologie: " + technology);
                        message += "techno detected:" + technology + "\n";
                        //technology = "";
                        //get the filename
                        @path_to_file = Clipper_Param.GetPath("Export_GPAO") + "\\" + technology + "\\" + nesting_name + ".txt";
                        message += "gpao file: " + @path_to_file + "\n";

                        if (!File.Exists(@path_to_file))
                        {//on cloture
                         //on reconstruit une liste des placaments                     
                         //nesting_to_close = current_nesting; //nestings_list.Where(x => x.GetFieldValueAsString("_NAME") == nesting_name).FirstOrDefault();
                            message += "fermeture du placement " + nesting_to_close.GetFieldValueAsString("_NAME") + "\n";
                            SimplifiedMethods.CloseNesting(Contextlocal, nesting_to_close);
                            log.WriteLogSuccess("Synthèse :\n " + message + "\n" + message);
                            System.Console.WriteLine("placement  " + nesting_to_close.GetFieldValueAsString("_NAME") + " fermé");

                        }
                        else
                        {

                            message += "le fichier de placement a été detecté -le placement ne sera pas cloturé " + "\n";
                        }
                        //suppresse the file
                        log.WriteLogEvent("Synthèse :\n " + message + "\n");
                    }
                }

                System.Console.Out.Close();
                return rst;

            }
            catch (Exception ie)
            {

                log.WriteLogWarningEvent("Probleme rencontré log de la cloture des placements :\n " + message);
                log.WriteLogWarningEvent("details :\n " + ie.Message);
                //System.Console.WriteLine("Erreur à la fermeture du placement " +ie.Message);
                //System.Console.ReadLine() ;
                return false;
            }
        }



        /// <summary>
        /// recupere la liste des placements exportes (confition exported)
        /// recupere la liste des fichier dans les dossiers gpao
        /// cloture si le fichier n'est plus present
        ///utilise les log standard alma 
        /// C:\> Remove-EventLog -LogName "MyLog"
        /// Remove-EventLog -Source "MyApp"
        /// </summary>
        /// <returns></returns>
        public bool Automatic_Nestings_Close(IContext Contextlocal, string Stage)
        {
            string message = ""; //message de suivi pour les log
            try
            {

                Clipper_Param.GetlistParam(Contextlocal);


                Alma_Log.Write_Log("Parametre recuperés");

                bool rst = false;
                //string stage = "_TO_CUT_NESTING";
                IEntityList nestings_list = null;
                IEntity current_nesting = null;

                Alma_Log.Write_Log("connecté à" + Contextlocal.Connection.DatabaseName + " Pour cloture");
                System.Console.WriteLine("connecté à" + Contextlocal.Connection.DatabaseName + " Pour cloture");
                Alma_Log.Write_Log("fermeture des placement " + Stage);
                System.Console.WriteLine("fermeture des placement " + Stage);

                nestings_list = Contextlocal.EntityManager.GetEntityList(Stage, Stage + "_GPAO_Exported", ConditionOperator.Equal, true);
                nestings_list.Fill(false);

                if (nestings_list != null && nestings_list.Count() > 0)
                {
                    foreach (IEntity nesting in nestings_list)
                    {
                        string nesting_name;
                        string technology;
                        string path_to_file;
                        //string msgstart;
                        //On initialise le message
                        message = "";
                        //IEntityList nestings_to_close;
                        IEntity nesting_to_close;

                        //nesting_to_close= nestings_list.Where(x=>x.GetFieldValueAsString("_NAME")== nesting.GetFieldValueAsString("_NAME"));
                        current_nesting = nesting;

                        //get the nesting name 
                        nesting_name = current_nesting.GetFieldValueAsString("_NAME");
                        //get the technology--> get the folder
                        message += "closing :" + nesting_name;
                        technology = AF_ImportTools.Machine_Info.GetNestingTechnologyName(ref Contextlocal, ref current_nesting);
                        message += "techno detected:" + technology;
                        //technology = "";
                        //get the filename
                        @path_to_file = Clipper_Param.GetPath("Export_GPAO") + "\\" + technology + "\\" + technology + "\\" + nesting_name + ".txt";
                        message += "gapo file: " + @path_to_file;
                        if (!File.Exists(@path_to_file))
                        {//on cloture
                         //on reconstruit une liste des placaments                     
                            nesting_to_close = nestings_list.Where(x => x.GetFieldValueAsString("_NAME") == nesting_name).FirstOrDefault();
                            Alma_Log.Write_Log("fermeture du placement " + nesting_to_close.GetFieldValueAsString("_NAME"));
                            System.Console.WriteLine("fermeture du placement " + nesting_to_close.GetFieldValueAsString("_NAME"));

                        }
                        //suppresse the file

                    }
                }

                System.Console.Out.Close();
                return rst;

            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log("Probleme rencontré log de la cloture des placements :\n " + message);
                Alma_Log.Write_Log("details :\n " + ie.Message);

                return false;
            }
        }









    }

    
    //le requirement fait etat des toles a commander(a savoir en qté negatives). cette fonction a ete abandonnée car nous suivie par clip
    //bouton : processor
    public class Clipper_Requirement_Processor : CommandProcessor
    {
        public IContext contextlocal = null;


        //appel de la lib d'export des besoins ici
        public override bool Execute()
        { //initialisation des listes
          // IContext _Context = null;
             contextlocal = Context; // modelsRepository.GetModelContext(DbName);
           



            return base.Execute();
        }
    }


    /// <summary>
    /// les commandes processor designent les boutons d'action integerés dans l'interface almaquote 
    /// </summary>


    /// <summary>
    /// cette classe lance l'application d'intergation pour le paramétrage des base, les tests d'import du stock...
    /// pour le moment cette application est l'executable suivant
    /// C:\AlmaCAM\Bin\AlmaCamTrainingTest.exe -->
    /// C:\AlmaCAM\Bin\AF_Clipper.exe
    /// </summary>
    /// en V7
    public class ClipperIE : CommandProcessor
    {
        public IContext contextlocal = null;
        public override bool Execute()
        {
            string Arguments = "";
            string FileName = @"C:\AlmaCAM\Bin\AF_Clipper.exe";
            var start_test = new ProcessStartInfo(FileName, Arguments);


            Process.Start(start_test);

            return base.Execute();
        }
    }


    // en V8 la commande est passée en globale
    public class ClipperIE_Global : SimpleCommandProcessor
    {
        //public IContext contextlocal = null;
        public override bool Execute()
        {
            string Arguments = "";
            string FileName = @"C:\AlmaCAM\Bin\AF_Clipper.exe";
            var start_test = new ProcessStartInfo(FileName, Arguments);


            Process.Start(start_test);

            return base.Execute();
        }
    }

    #endregion
    

    //PARAMETRES
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////IMPORT///////////////OF////////////////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    #region parameters
    /// <summary>
    /// la classe clipper_param recupere les paramètres d almacam dans les options spécifiques de la passerelle 
    /// comme la lecture des paramètres est indispensable, cette classe verifie aussi la presence des dossier clipper , le nom de la base ainsi
    /// que la compatibilité  d almacam avec la passerelle via les informations de ressources.
    /// </summary>
    [Obsolete("bientot remplacer par  Clipper_Param_21sp6 pour la sp6")]
    public static class Clipper_Param
    {

        public static Dictionary<string, object> Parameters_Dictionnary; // liste des path et des types pour le format du fichier de stock et des of





        /*
        // public static IContext context depuis sp5
        public void Clipper_Param(IContext context)
        {
            Parameters_Dictionnary = new Dictionary<string, object>();
            GetlistParam( context);

        }*/

        // public static IContext context;
        static Clipper_Param()
        {
            Parameters_Dictionnary = new Dictionary<string, object>();
          

        }
        /// <summary>
        /// recuperation des path clipper+fichier.csv (echange csv) ou autre
        /// creer le dictionnaire des paramètres necessaire pour la passerelle depuis sp5 
        /// ATTENTION IL S AGIT D UN DICTIONNAIRE STRING OBJECT
        /// UTILISER GetParam pour recuperer un parametre convertie correctement
        /// </summary>
        /// //exemple "H:\tutu\toto\cahieraffaire.csv"
        /// <param name="context"> contexte </param>
        public static Boolean GetlistParam(IContext context)
        {

            try
            {
                string parametre_name;

                Parameters_Dictionnary.Clear();


                
                // CHANGEMENT DU dictionnaire de parametres NOM DES PARAMETRES DEPUIS SP4
                string parametersetkey = "CLIP_CONFIGURATION";
                parametre_name = "IMPORT_CDA";
                context.ParameterSetManager.TryGetParameterValue(parametersetkey, parametre_name, out IParameterValue sp5);
                if (sp5 == null) { MessageBox.Show("Les paramètres de configurations generaux de la passerelle ne sont pas definis, merci de mettre à jour la base avec les outils de migration");
                    System.Environment.Exit(0);
                }

                Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary);

                parametre_name = "IMPORT_DM";
                Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary);

                parametre_name = "Export_GPAO";
                //Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                 Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "EXPORT_Rp", ref Parameters_Dictionnary);

                parametre_name = "EXPORT_DT";
                Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "EXPORT_Dt", ref Parameters_Dictionnary);

                //description import
                parametre_name = "IMPORT_AUTO";
                Get_bool_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary, false);

                //description import
                parametre_name = "ACTIVATE_OMISSION";
                Get_bool_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary, true);

                //description import
                parametre_name = "AF_ACTIVATE_SHEET_ON_SENDTOWSHOP";
                Get_bool_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary, true);


                if (Convert.ToBoolean(Parameters_Dictionnary["ACTIVATE_OMISSION"]) == false)
                { }//Alma_Log.Warning("ATTENTION : Le traitement par omission du stock est desactivé.", "PARAMATER_DICTIONARY");}


                parametre_name = "EMF_DIRECTORY";
                 Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary);
                
                parametre_name = "MODEL_CA";
                Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary);

                parametre_name = "MODEL_DM";
                Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary);

                parametre_name = "MODEL_PATH";
                Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary);

                parametre_name = "APPLICATION1";
                Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary);

                parametre_name = "SHEET_REQUIREMENT";
                Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary);

                //valeur de quantité par defaut sur les chutes af_almacam
                //repertoire de exports dpr
                parametre_name = "AF_ACTIVATE_QTY_ON_SENDTOWSHOP";
                Get_long_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary,0);

                //log
                parametre_name = "VERBOSE_LOG";
                //Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue(parametersetkey, "VERBOSE_LOG").GetValueAsBoolean());
                //nom mahine clipper
                parametre_name = "CLIPPER_MACHINE_CF";
                //Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                //Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue(parametersetkey, "CLIPPER_MACHINE_CF").GetValueAsString());
                Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary);

                /*parametres de sorties*/
                parametre_name = "STRING_FORMAT_DOUBLE";
                //Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary);

                parametre_name = "ALMACAM_EDITOR_NAME";
                Get_string_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary);

                //parametre export : chemin de sortie des devis
                parametre_name = "_EXPORT_GP_DIRECTORY";
                Get_string_Parameter_Dictionary_Value(context, "_EXPORT", parametre_name, "", ref Parameters_Dictionnary);

                //repertoire de exports dpr
                parametre_name = "_ACTCUT_DPR_DIRECTORY";
                Get_string_Parameter_Dictionary_Value(context, "_EXPORT", parametre_name, "", ref Parameters_Dictionnary);
                //devis
                //recuperatin de la case a cocher prix au kilo --> elle doit pour le moment roujour etre cochée 
                parametre_name = "_SHEET_SPECIFIC_SALE_COST_BY_WEIGHT";
                Get_bool_Parameter_Dictionary_Value(context, "_QUOTE", parametre_name, "", ref Parameters_Dictionnary, true);

                parametre_name = "AF_MULTIDIM_MODE";
                Get_bool_Parameter_Dictionary_Value(context, "CLIP_CONFIGURATION", parametre_name, "", ref Parameters_Dictionnary, true);


                //Champs Spécifique a reporter à partir des information des pieces à produire et de lors de l'import gpao dans les pieces2d (reference almacam)
                //entrez l'information du nom de champs 2d puis le nom du champs du line_dictionnary
                // NOM_CHAMPS|NOM_CHAMPS_DU_LINE_DICTIONNARY
                //CUSTOMER_REFERENCE_INFOS;
                parametre_name = "CUSTOMER_REFERENCE_INFOS";
                //Alma_Log.Info("recuperation du parametre " + parametre_name, "GetlistParam");
                ///on entre les infos 
                ///Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue("_EXPORT", "CUSTOMER_REFERENCE_INFOS").GetValueAsString());
                Parameters_Dictionnary.Add(parametre_name, "CUSTOMER|_FIRM");
                ///auteur
                parametre_name = "_AUTHOR";
                Parameters_Dictionnary.Add(parametre_name, context.UserId);

                //option du workshop GlobalClosedSeparated or GlobalCloseOneClic
                WorkShopOptionType WORKSHOP_OPTION = ActcutModelOptions.GetWorkShopOption(context);
                parametre_name = "_WORKSHOP_OPTION ";
                Parameters_Dictionnary.Add(parametre_name, WORKSHOP_OPTION.ToString());

                //verification des path
                //Alma_Log.Info("verification de l'existance des path ", "GetlistParam");
                if (CheckClipperFolderExists() == false) { throw new System.ApplicationException("Certains chemin d'echanges de la Passerelle AlmaCam-clipper ne sont pas accessibles"); };
                if (CkeckCompatibilityVersion() == false) { throw new System.ApplicationException("Version de la Dll clipper n'est pas validée pour cette version d'AlmaCam"); }


                //verification des chemins
                parametre_name = "EXPLODE_MULTIPLICITY";
                Get_bool_Parameter_Dictionary_Value(context, parametersetkey, parametre_name, "", ref Parameters_Dictionnary, false);


                return true;
            }



            catch (KeyNotFoundException)
            {
                //Alma_Log.Error(ex, "CETTE BASE NE SEMBLE PAS ETRE CONFIGUREE POUR CLIPPER !!! ");
                //Alma_Log.Error(ex, "Veuiller verifier la configuration des paramètres de l'import clipper (nom et id des champs....)");
                MessageBox.Show(Alma_RegitryInfos.GetLastDataBase() + " :CETTE BASE NE SEMBLE PAS ETRE CONFIGUREE POUR CLIPPER !!! \r\n " +
                "Veuillez verifier la base selectionnées pour l'ouverture d'AlmaCam");
                //on sort
                //ClipperExit.Close();
                return false;
            }
            catch (System.ApplicationException exVersion)
            {
                MessageBox.Show(exVersion.Message);
                //Alma_Log.Error(exVersion, "Version incompatible ou mauvaise configuration de la base almacam");
                //ClipperExit.Close();
                return false;
            }

            catch (System.IO.DirectoryNotFoundException exFolder)
            {
                MessageBox.Show(exFolder.Message);
                //Alma_Log.Error(exFolder, "L'un des dossier de d'echange n'existe pas");
                return false;
           
            }

            finally
            {
                //Alma_Log.Error("done", "Clipper Param Read");
            }
          
        }

        /// <summary>
        /// retourn la valeur de la clé recherché dans les paramètres
        /// </summary>
        /// <typeparam name="T">type générique </typeparam>
        /// <param name="context">contexte </param>
        /// <param name="PathVariable">clé a rechercher</param>
        /// <returns></returns>
        public static T GetParam<T>(this IDictionary<string, object> dic, string key)
        {
            if (Parameters_Dictionnary.ContainsKey(key))
            {
                return (T)dic[key];
            }
            else { return default(T); }
        }

        /// <summary>
        /// ecrit la valeur de la clé recherché dans les paramètres
        /// </summary>
        /// <typeparam name="T">type générique </typeparam>
        /// <param name="context">contexte </param>
        /// <param name="PathVariable">clé a rechercher</param>
        /// <returns></returns>
        public static void SetParam<T>(this IDictionary<string, object> dic, string key, T value)
        {
            try
            {
                if (Parameters_Dictionnary.ContainsKey(key))
                {
                    dic[key] = value;
                }


            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message, "Clipper Param  : SetParam ERROR");
            }

        }
        /// <summary>
        /// recuperation des parametres de type string
        /// </summary>
        /// <param name="contextlocal">contexte a etudier</param>
        /// <param name="parametersetkey">cle du jeu d'option general nom de la dll </param>
        /// <param name="parameter_name">nom du parametre dans le dictionnaire</param>
        /// <param name="parameterkeyname">nom de la clé almacam stockant le parametre</param>
        /// <param name="parameters_dictionnary">nom du dictionnaire</param>
        /// <returns></returns>
        private static bool Get_string_Parameter_Dictionary_Value(IContext contextlocal, string parametersetkey, string parameter_name, string parameterkeyname, ref Dictionary<string, object> parameters_dictionnary)
        {
            try
            {
                bool rst = false;
                //IParameterValue value;
                if (string.IsNullOrEmpty(parameterkeyname)) { parameterkeyname = parameter_name; }
                Alma_Log.Info("recuperation du parametre " + parameter_name, "GetlistParam");
                rst = contextlocal.ParameterSetManager.TryGetParameterValue(parametersetkey, parameterkeyname, out IParameterValue value);
                if (rst == false)
                {
                    rst = false;
                    parameters_dictionnary.Add(parameter_name, "");
                    throw new MissingParameterException(parameter_name);

                }
                else
                {
                    // Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue(parametersetkey, "IMPORT_CDA").GetValueAsString());
                    parameters_dictionnary.Add(parameter_name, value.GetValueAsString());
                }


                return rst;
            }
            catch (MissingParameterException)
            {

                return false;
            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);
                return false;
            }



        }
        /// <summary>
        /// recuperation des parametres de type string
        /// </summary>
        /// <param name="contextlocal">contexte a etudier</param>
        /// <param name="parametersetkey">cle du jeu d'option general nom de la dll </param>
        /// <param name="parameter_name">nom du parametre dans le dictionnaire</param>
        /// <param name="parameterkeyname">nom de la clé almacam stockant le parametre</param>
        /// <param name="parameters_dictionnary">nom du dictionnaire</param>
        ///  <param name="parameters_dictionnary">valeur par defaut du parametre</param>
        /// <returns></returns>
        private static bool Get_bool_Parameter_Dictionary_Value(IContext contextlocal, string parametersetkey, string parameter_name, string parameterkeyname, ref Dictionary<string, object> parameters_dictionnary, bool defaultvalue)
        {
            try
            {
                bool rst = false;

                if (string.IsNullOrEmpty(parameterkeyname)) { parameterkeyname = parameter_name; }
                Alma_Log.Info("recuperation du parametre " + parameter_name, "GetlistParam");
                rst = contextlocal.ParameterSetManager.TryGetParameterValue(parametersetkey, parameterkeyname, out IParameterValue value);
                if (rst == false)
                {
                    rst = false;
                    parameters_dictionnary.Add(parameter_name, defaultvalue);
                    throw new MissingParameterException(parameter_name);

                }
                else
                {
                    // Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue(parametersetkey, "IMPORT_CDA").GetValueAsString());
                    parameters_dictionnary.Add(parameter_name, value.GetValueAsBoolean());
                  
                    
                }


                return rst;
            }
            catch (MissingParameterException)
            {

                return false;
            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);
                return false;
            }



        }
        /// <summary>
        /// recuperation des parametres de type long
        /// </summary>
        /// <param name="contextlocal">contexte a etudier</param>
        /// <param name="parametersetkey">cle du jeu d'option general nom de la dll </param>
        /// <param name="parameter_name">nom du parametre dans le dictionnaire</param>
        /// <param name="parameterkeyname">nom de la clé almacam stockant le parametre</param>
        /// <param name="parameters_dictionnary">nom du dictionnaire</param>
        ///  <param name="parameters_dictionnary">valeur par defaut du parametre</param>
        /// <returns></returns>
        private static bool Get_long_Parameter_Dictionary_Value(IContext contextlocal, string parametersetkey, string parameter_name, string parameterkeyname, ref Dictionary<string, object> parameters_dictionnary, long defaultvalue)
        {
            try
            {
               bool rst = false;

                if (string.IsNullOrEmpty(parameterkeyname)) { parameterkeyname = parameter_name; }
                Alma_Log.Info("recuperation du parametre " + parameter_name, "GetlistParam");
                rst = contextlocal.ParameterSetManager.TryGetParameterValue(parametersetkey, parameterkeyname, out IParameterValue value);
                if (rst == false)
                {
                    rst = false;
                    parameters_dictionnary.Add(parameter_name, defaultvalue);
                    throw new MissingParameterException(parameter_name);

                }
                else
                {
                    // Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue(parametersetkey, "IMPORT_CDA").GetValueAsString());
                    parameters_dictionnary.Add(parameter_name, value.GetValueAsLong());


                }


                return rst;
            }
            catch (MissingParameterException)
            {

                return false;
            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);
                return false;
            }



        }
        /// <summary>
        /// retourn la valeur de la clé recherché dans les paramètres
        /// </summary>
        /// <typeparam name="T">type générique </typeparam>
        /// <param name="context">contexte </param>
        /// <param name="PathVariable">clé a rechercher</param>
        /// <returns></returns>
        public static T TryGetParam<T>(this IDictionary<string, object> dic, string key)
        {
            //if (Parameters_Dictionnary.ContainsKey(key))
            {
                try
                {



                    dic.TryGetValue(key, out object obj);

                    return (T)obj;

                }

                catch (KeyNotFoundException)
                {
                    return default(T);


                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    return default(T);
                }

            }
            ///else { return default(T); }
        }

        /// <summary>
        /// retourne un chemin windows type string
        /// </summary>
        /// <param name="key">nom de la clé dans le dictionnaire</param>
        /// <returns>chemin windows : type c:\actcut...</returns>
        public static string GetPath(string key)
        {

            //GetlistParam(context);//    Alma_Log.Info("verification de la clé " + key, "GetPath");
            try
            {
                // if (Parameters_Dictionnary.ContainsKey(key))
                {
                    return Parameters_Dictionnary[key].ToString();
                }
            }
            catch (Exception ie)
            {
                Alma_Log.Info("impossible de trouver la clé  " + key, "GetPath");
                Alma_Log.Info("impossible de trouver la clé  " + ie.Message, "GetPath");
                return "Undef";
            }


        }
        /// <summary>
        /// recupere la case a cocher d'automation
        /// </summary>
        /// <returns>boolean true/false</returns>
        public static bool IsAutomatiqueImport()
        {
            string key = "IMPORT_AUTO";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (bool)Parameters_Dictionnary[key]; } else { return false; }

        }

        /// <summary>
        /// recupere
        /// </summary>
        /// <returns>un model de fichier csv sous la forme d'une liste de champs 
        /// champs : numero du champ dans #nom du champs  dans almacam#Type  -> plus tard si besoin numero du champ dans #nom du champs  dans almacam#Type  # taille max
        /// 0#NAME#string;1#AFFAIRE#string;2#THICKNESS#string;3#MATERIAL_CLIPPER#string;4#CENTREFRAIS#string;5#TECHNOLOGIE#string;6#FAMILY#string;7#IDLNROUT#string;8#CENTREFRAISSUIV#string;9#CUSTOMER#string;10#PART_INITIAL_QUANTITY#double;11#QUANTITY#double;12#ECOQTY#double;13#STARTDATE#date;14#ENDDATE#date;15#PLAN#string;....
        /// 
        /// </returns>
        public static string GetModelCA()
        {
            string key = "MODEL_CA";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Undef  model CA"; }

        }


        /// <summary>
        /// recupere la validation tole
        /// </summary>
        /// <returns>tue if valide false if not
        /// </returns>
        public static bool GetSheetAutoValidationMode()
        {
            string key = "AF_ACTIVATE_SHEET_ON_SENDTOWSHOP";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (bool)Parameters_Dictionnary[key]; } else { return false; }

        }

        public static string GetModelDM()
        {
            string key = "MODEL_DM";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Undef  model  DM"; }

        }


        public static string GetModelPATH()
        {
            string key = "MODEL_PATH";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Undef model PATH"; }

        }


        public static string Get_string_format_double()
        {

            string key = "STRING_FORMAT_DOUBLE";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Undef STRING_FORMAT_DOUBLE"; }

        }
        /// <summary>
        /// explode multiplicity : mutliplicité, 
        /// explose la multiplicité 
        /// False : un fichier pour une mutliplicité n
        /// True : tole ou n fichiers pour une mutliplicité n
        /// </summary>
        /// <returns></returns>
        public static bool Get_Multiplicity_Mode()
        {

            string key = "EXPLODE_MULTIPLICITY";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (bool)Parameters_Dictionnary[key]; } else { return false; }

        }

        public static string Get_application1()
        {
            string key = "APPLICATION1";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)@Parameters_Dictionnary[key]; } else { return "Undef model PATH"; }
        }

        public static bool GetVerbose_Log()
        {
            string key = "VERBOSE_LOG"; //log verbeux
            if (Parameters_Dictionnary.ContainsKey(key)) { return (bool)Parameters_Dictionnary[key]; } else { return true; }
        }
        
        public static string Get_Clipper_Machine_Cf()
        {
            string key = "CLIPPER_MACHINE_CF"; //log verbeux
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Undef clipper machine"; }

        }

        //retourne le get_mutlidim//
        //retourne la valeur de la case a cocher is_mutlidim
        public static bool Get_MULTIDIM_MODE(IContext contextlocal)
        {
            
            //IContext contextlocal, string parametersetkey, string parameter_name, string parameterkeyname, ref Dictionary<string, object> parameters_dictionnary, bool defaultvalue)
            string parameterkeyname = "AF_MULTIDIM_MODE";
            string parameter_name = "";
            string parametersetkey = "CLIP_CONFIGURATION";
            bool IsMultidimMode = true;
            try
            {
                bool rst = false;

                
                Alma_Log.Info("recuperation du parametre " + parameterkeyname, "GetlistParam");
                rst = contextlocal.ParameterSetManager.TryGetParameterValue(parametersetkey, parameterkeyname, out IParameterValue value);
                if (rst == false)
                {
                    rst = false;
                    //parameters_dictionnary.Add(parameter_name, defaultvalue);
                    throw new MissingParameterException(parameter_name);

                }
                else
                {
                    // Parameters_Dictionnary.Add(parametre_name, context.ParameterSetManager.GetParameterValue(parametersetkey, "IMPORT_CDA").GetValueAsString());
                    //parameters_dictionnary.Add(parameter_name, value.GetValueAsBoolean());
                    IsMultidimMode = value.GetValueAsBoolean();

                }


                return IsMultidimMode;
            }
            catch (MissingParameterException)
            {

                return false;
            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);
                return false;
            }




        }



        //retourne le get_mutlidim//
        //retourne la valeur de la case a cocher is_mutlidim
        public static bool Get_Price_Mode()
        {
            string key = "_SHEET_SPECIFIC_SALE_COST_BY_WEIGHT"; //multidim
            if (Parameters_Dictionnary.ContainsKey(key)) { return (bool)Parameters_Dictionnary[key]; } else { return false; }
        }

        /// <summary>
        ///  Ce paramètre active ou descative m'omission 
        /// </summary>
        /// <returns></returns>
        public static bool Get_Omission_Mode()
        {

            string key = "ACTIVATE_OMISSION";
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (bool)Parameters_Dictionnary[key]; } else { return false; }

        }

        public static string Get_AlmaCamEditorName()
        {
            string key = "ALMACAM_EDITOR_NAME"; //log verbeux
            //GetlistParam(context);//
            if (Parameters_Dictionnary.ContainsKey(key)) { return (string)Parameters_Dictionnary[key]; } else { return "Undef almacam editor"; }

        }
        /// <summary>
        /// verification de l'existance des dossiers d'echange
        /// </summary>
        /// <returns>true si tous les dossier existent, false si ils n'existent pas</returns>
        public static Boolean CheckClipperFolderExists()
        {
            try
            {

                Alma_Log.Info("checking IMPORT_CDA", "checkClipperFolderExists");
                Directory.GetDirectories(Path.GetDirectoryName(Clipper_Param.GetPath("IMPORT_CDA")));
                Alma_Log.Info("checking IMPORT_DM", "checkClipperFolderExists");
                Directory.GetDirectories(Path.GetDirectoryName(Clipper_Param.GetPath("IMPORT_DM")));
                Alma_Log.Info("checking Export_GPAO", "checkClipperFolderExists");
                Directory.GetDirectories(Path.GetDirectoryName(Clipper_Param.GetPath("Export_GPAO")));
                Alma_Log.Info("checking EXPORT_DT", "checkClipperFolderExists");
                Directory.GetDirectories(Path.GetDirectoryName(Clipper_Param.GetPath("EXPORT_DT")));


                return true;
            }
            catch (System.IO.IOException ie)   //.DirectoryNotFoundException ie)
            {
                Alma_Log.Error(ie, " les dossiers d'echange sont mal definis verifier : IMPORT_CDA, IMPORT_DM ,Export_GPAO, EXPORT_Dt ,_EXPORT_GP_DIRECTORY");
                return false;
                throw;
            }


        }
        /// <summary>
        /// verification sur le nom de la base de données
        /// car clipper a écrit en dur le nom de la base de données
        /// </summary>
        /// <returns>true si le nom de la base est correcte</returns>
        public static Boolean CheckDatabasename(string actualdatabasename)
        {
            try
            {
                bool res = true;

                string lastopenDbname = Alma_RegitryInfos.GetLastDataBase();
                if (Alma_RegitryInfos.GetLastDataBase().Count() == 0)
                {

                    new Exception("Impossible de lire le nom de la dernière base ouvert dans le registre");
                    { throw new Exception(Properties.Resources.Clipper_Almacam_Database_Name.ToString() + ", la base :" + Properties.Resources.Clipper_Almacam_Database_Name + " est introuvable"); }

                }

                else
                {

                    string working_db = Properties.Resources.Clipper_Almacam_Database_Name.ToString();
                    Alma_Log.Write_Log("derniere base ouverte: " + working_db);
                    if (Clipper_Param.CheckDatabasename(working_db) == false) { throw new Exception(Alma_RegitryInfos.GetLastDataBase() + ", le Nom de la base est incorrecte,  elle doit se nommer :" + Properties.Resources.Clipper_Almacam_Database_Name); }


                }
                return res;
            }
            //catch (Exception e)
            catch (Exception ex)
            {
                Alma_Log.Write_Log(ex.Message);
                MessageBox.Show(ex.Message);
                return false;
                throw;
            }




        }
        /// <summary>
        /// retourne la verison de la dll clipperDll.dll
        /// </summary>
        /// <returns></returns>
        public static string GetClipperDllVersion()
        {
            return Application.ProductVersion.ToString().Substring(0, 3);
        }

        /// <summary>
        /// retourn la version compatible almacam indiquee dans les ressources almacam
        /// </summary>
        /// <returns></returns>
        public static string GetAlmaCAMCompatibleVersion()
        {
            return Properties.Resources.Almcam_Version.ToString();

        }
        /// <summary>
        /// recupere le numero de version compatible. et le compare a la version de l'executable almacam
        /// la version est bloquée par une infos des ressources de la dll clipper_dll
        /// </summary>
        /// <returns>true si le test est accepté</returns>
        public static Boolean CkeckCompatibilityVersion()
        {
            bool res = false;

            try
            {

                bool compatible = false;
                string versionalmacam;
                string almacameditorfullpath = Directory.GetCurrentDirectory().ToString() + "\\" + Get_AlmaCamEditorName();
                string almacamCompatibleversion = Properties.Resources.Almcam_Version.ToString();
                //get version//
                versionalmacam = FileVersionInfo.GetVersionInfo(almacameditorfullpath).ProductVersion.ToString();


                foreach (string v in almacamCompatibleversion.Split(';'))
                {
                    if (versionalmacam.StartsWith(v))
                    {
                        compatible = true;
                        break;
                    };
                }


                ///

                if (compatible == true)
                {
                    Alma_Log.Write_Log("verion wpm.exe :" + versionalmacam + " version autoriséee pour cette dll " + almacamCompatibleversion);
                    Alma_Log.Write_Log("test ok");
                    res = true;
                }
                else
                {
                    Alma_Log.Write_Log("la librairie clipperdll.dll version " + almacamCompatibleversion + " est  incompatible Almacam  " + versionalmacam);
                }

                return res;

            }
            catch
            {

                Alma_Log.Write_Log(": time tag:  ");
                return res;

            }

        }


    }
    

    //pour sp6 on passe en classe non static derivé d'import param
    public  class Clipper_Param_21sp6 : Import_Param
    {

        // liste des path et des types pour le format du fichier de stock et des of

        public Clipper_Param_21sp6(IContext contextlocal):base(contextlocal)
        { }




    }

    


    #endregion


















    #region clipper_8

    #region CommandePorcessor
    public class Clipper_8_Import_OF_Processor : CommandProcessor
    {
        //public IContext contextlocal = null;
        public override bool Execute()
        {
            //recuperation du context
            //string DbName = Alma_RegitryInfos.GetLastDataBase();
            // IModelsRepository modelsRepository = new ModelsRepository();
            //contextlocal = modelsRepository.GetModelContext(DbName);
            //contextlocal = Context;
            Clipper_Param.GetlistParam(Context);
            string csvImportPath = Clipper_Param.GetPath("IMPORT_CDA");
            //recuperation du nom de fichier
            string csvFileName = Path.GetFileNameWithoutExtension(csvImportPath);
            string dataModelstring = Clipper_Param.GetModelCA();


            // if (contextlocal != null)
            if (Context != null)
            {
                using (Clipper_8_Import_OF cahierAffaire = new Clipper_8_Import_OF())
                {
                    cahierAffaire.Import(Context, csvImportPath, dataModelstring, false);//), csvImportPath);

                }
            }
            return base.Execute();
        }
    }
    ////boutnon d'import du stock
    ///import du stock --> stock et pas sheet
    public class Clipper_8_Import_Stock_Processor : CommandProcessor
    {
        //public IContext contextlocal = null;
        public override bool Execute()
        {


            //declaration des listes pour post traitement

            using (var Stock = new Clipper_8_Import_Stock())
            {

                Stock.Import(Context);


            }

            // MessageBox.Show(" Import terminé");

            return base.Execute();
        }
    }
    //:export des données technique --> piece 2d
    public class Clipper_8_Export_DT_Processor : CommandProcessor
    {

        public override bool Execute()
        {

            try
            {
                Execute(Context);

                return base.Execute();
            }

            catch (Exception ie)
            {

                MessageBox.Show(ie.Message);
                return base.Execute();


            }




        
        }

        public bool Execute(IContext contextlocal)
        {


            if (contextlocal != null)
            {
                AF_Clipper_Dll.Clipper_8_RemonteeDt Remontee_Dt = new AF_Clipper_Dll.Clipper_8_RemonteeDt();

                DialogResult result1 = MessageBox.Show("Voulez vous exporter des pieces à produire ?", "WARNING !!!", MessageBoxButtons.YesNo);

                if (result1 == DialogResult.No)
                {
                    //export sur la base des peices 2d (pas de numero de gamme associé)
                    Remontee_Dt.Export_Piece_To_File(contextlocal, false);
                }
                else
                {   //export des données sur la base des pieces à produire
                    Remontee_Dt.Export_Part_To_Produce_To_File(contextlocal, true);
                }






            }
            return true;

        }

        

    }
    //export des données technique --> piece 2d
    public class Clipper_8_Export_DT_Export_Piece_To_File_Processor : CommandProcessor
    {

        public override bool Execute()
        {

            try
            {
                Execute(Context);

                return base.Execute();
            }

            catch (Exception ie)
            {

                MessageBox.Show(ie.Message);
                return base.Execute();


            }





        }

        public bool Execute(IContext contextlocal)
        {


            if (contextlocal != null)
            {
                AF_Clipper_Dll.Clipper_8_RemonteeDt Remontee_Dt = new AF_Clipper_Dll.Clipper_8_RemonteeDt();
                  
                    Remontee_Dt.Export_Piece_To_File(contextlocal, false);
            }
            return true;

        }



    }
    //gestion de production de sortie clipper ou retour GP
    public class Clipper_8_Export_GP_Processor : CommandProcessor
    {

        public override bool Execute()
        {

            try
            {
                Execute(Context);

                return base.Execute();
            }

            catch (Exception ie)
            {

                MessageBox.Show(ie.Message);
                return base.Execute();


            }





        }

        public bool Execute(IContext contextlocal)
        {


            if (contextlocal != null)
            {

                var doonaction = new Clipper_8_DoOnAction_AfterSendToWorkshop();
                //string stage = "_TO_CUT_NESTING";
                //creation du fichier de sortie
                //recupere les path
                Clipper_Param.GetlistParam(contextlocal);


                //IEntitySelector nestingselector = null;
                //nestingselector = new EntitySelector();
                //entity type pointe sur la list d'objet du model
                //nestingselector.Init(contextlocal, contextlocal.Kernel.GetEntityType(stage));
                //nestingselector.MultiSelect = true;
                

                //if (nestingselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (IEntity nesting in Command.WorkOnEntityList)
                    {
                        doonaction.Execute(nesting);

                    }
                }
               

                




            }
            return true;

        }



    }
    //export des données technique --> piece 2d avec les informztion des pieces a produire
    public class Clipper_8_Export_Part_To_Produce_To_File_Processor : CommandProcessor
    {

        public override bool Execute()
        {

            try
            {
                Execute(Context);

                return base.Execute();
            }

            catch (Exception ie)
            {

                MessageBox.Show(ie.Message);
                return base.Execute();


            }





        }

        public bool Execute(IContext contextlocal)
        {


            if (contextlocal != null)
            {
                AF_Clipper_Dll.Clipper_8_RemonteeDt Remontee_Dt = new AF_Clipper_Dll.Clipper_8_RemonteeDt();

                
                    Remontee_Dt.Export_Part_To_Produce_To_File(contextlocal, true);

            }
            return true;

        }



    }
    #endregion


    #region cahieraffaire

    /// <summary>
    /// recupere les pieces exportées de clipper dans un fchier csv
    /// ce processus se deroule en 2 phases
    /// une phase de verification des données contenu dans le csv
    /// une phase d'ecriture dans la base almacam
    /// </summary>
    public class Clipper_8_Import_OF : IDisposable, IImport
    {
        string CsvImportPath = null;


        public void Dispose()
        {
           GC.SuppressFinalize(this);
        }



        /// <summary>
        /// creer une nouvelle reference a produire sous reserve de sans_donnees_technique=true et de centre de frais et is non presente
        /// PREND EN COMPTE LES PIECES SANS DONNEES TECHNIQUES (voir clipper pour explicaiton)
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="line_dictionnary"></param>
        /// <param name="CentreFrais_Dictionnary">dictionnaire des centres de frais</param>
        /// <param name="reference_to_produce">retourn de la nouvelle reference a produire si besoin</param>
        /// <param name="reference">reference a pointer</param>
        /// <param name="timetag">time tag de l'import : pour groupement </param>
        /// <param name="sans_donnees_technique">si true alors oon ne creer jamais de reference</param>
        /// <returns></returns>
        public bool CreateNewPartToProduce(IContext contextlocal, Dictionary<string, object> line_dictionnary, Dictionary<string, string> CentreFrais_Dictionnary, ref IEntity reference_to_produce, ref IEntity reference, string timetag, bool sans_donnees_technique)
        {
            bool result = false;

            try
            {
                //la piece ne contient pas de gamme
                //cas des pieces oranges : Pas de cf, pas de id_piece_cfao, on considere que c'est une piece orange--> on ne creer que la reference. 
                if (Data_Model.ExistsInDictionnary("CENTREFRAIS", ref line_dictionnary) == false || sans_donnees_technique == true)
                {
                    return false;
                }
                //string referenceName = null;
                Boolean need_prep = true;
                //int index_extension = 0;  //> 0 si ;emf;dpr detectée
                PartInfo machinable_part_infos = null; //infos de machinabe part

                //creation de la nouvelle reference
                reference_to_produce = contextlocal.EntityManager.CreateEntity("_TO_PRODUCE_REFERENCE");
                //recuperation et assignaton de la machine si elle existe
                string machine_name = "";
                Data_Model.ExistsInDictionnary(line_dictionnary["CENTREFRAIS"].ToString(), ref CentreFrais_Dictionnary);
                //lecture des part infos (optionnel) car le get reference fait le travail                 
                machinable_part_infos = new PartInfo();
                bool fastmode = true;
                //bool result = false;
                machinable_part_infos.GetPartinfos(ref contextlocal, reference);
                //on controle que la matiere de la reference correspondante -soit bonne sinon on ignore la ligne courante et on passe à la ligne suivante//
                if (fastmode)
                {

                    if (Data_Model.ExistsInDictionnary(line_dictionnary["CENTREFRAIS"].ToString(), ref CentreFrais_Dictionnary))
                    {
                        try { need_prep = need_prep && machinable_part_infos.IsPartDefault_Preparation(reference, CentreFrais_Dictionnary[line_dictionnary["CENTREFRAIS"].ToString()]); }

                        catch (KeyNotFoundException)
                        {
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + "le centre de Frais  ne pointe pas vers une machine connue");
                            MessageBox.Show(string.Format("le centre de Frais {0} ne pointe pas vers une machine connue", CentreFrais_Dictionnary[line_dictionnary["CENTREFRAIS"].ToString()]));
                        }
                    }

                }
                else
                //mode plus lent qui consiste a utiliser la premiere machine qui existe pour l'affecter a la piece a produire en cas de machine non resnseignée
                //slowmode
                //methode par comparaison d'id, jamais utilisé
                {
                    if (machine_name != "")
                    {
                        IEntityList machines = null;
                        IEntity machine = null;

                        machine = AF_ImportTools.SimplifiedMethods.GetFirtOfList(machines);
                        string mm = machinable_part_infos.DefaultMachineName;
                        need_prep = need_prep && true;
                    }
                    else
                    {
                        need_prep = need_prep && true;

                    }

                }


                ///ecriture de la piece a produire
                //les times tag permettent de savoir si la piece la piece a réelement ete importée et non cree a la main
                //ecriture du time tag
                reference_to_produce.SetFieldValue("OF_IMPORT_NUMBER", timetag.Replace("_", ""));
                //ecriture de la reference ( piece 2d)
                reference_to_produce.SetFieldValue("_REFERENCE", reference.Id32);
                reference_to_produce.SetFieldValue("NEED_PREP", need_prep);

                //mise a jour des autres champs 
                Update_Part_Item(contextlocal, ref reference_to_produce, ref CentreFrais_Dictionnary, ref line_dictionnary);
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + "update infos succeed");
                reference_to_produce.Save();
                line_dictionnary.Clear();

                return result;
            }
            catch { return result; }
            finally
            {

            }
        }


        /// <summary>
        /// creer une nouvelle reference a produire sous reserve de sans_donnees_technique=true et de centre de frais et is non presente
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="line_dictionnary"></param>
        /// <param name="CentreFrais_Dictionnary"></param>
        /// <param name="reference_to_produce"></param>
        /// <param name="reference"></param>
        /// <param name="timetag"></param>
        /// <returns></returns>
        public bool CreateNewPartToProduce(IContext contextlocal, Dictionary<string, object> line_dictionnary, Dictionary<string, string> CentreFrais_Dictionnary, ref IEntity reference_to_produce, ref IEntity reference, string timetag)
        {
            bool result = false;

            try
            {
                //la piece ne contient pas de gamme
                //cas des pieces oranges : Pas de cf, pas de id_piece_cfao, on considere que c'est une piece orange--> on ne creer que la reference. 
                if ((Data_Model.ExistsInDictionnary("CENTREFRAIS") == false) && (Data_Model.ExistsInDictionnary("CENTREFRAIS") == false))
                {
                    return false;
                }
                //string referenceName = null;
                Boolean need_prep = false;
                //int index_extension = 0;  //> 0 si ;emf;dpr detectée
                PartInfo machinable_part_infos = null; //infos de machinabe part

                //creation de la nouvelle reference
                reference_to_produce = contextlocal.EntityManager.CreateEntity("_TO_PRODUCE_REFERENCE");
                //recuperation et assignaton de la machine si elle existe
                string machine_name = "";
                //Data_Model.ExistsInDictionnary(line_dictionnary["CENTREFRAIS"].ToString(), ref CentreFrais_Dictionnary);
                //lecture des part infos (optionnel) car le get reference fait le travail                 
                machinable_part_infos = new PartInfo();
                bool fastmode = true;
                //bool result = false;
                machinable_part_infos.GetPartinfos(ref contextlocal, reference);
                //
                //on control que la matiere de la reference correspond -soit bonne sinon on continue a la ligne suivante//
                if (fastmode)
                {

                    //try { reference_to_produce.SetFieldValue("NEED_PREP", !part_infos.IsPartDefault_Preparation(contextlocal, reference, CentreFrais_Dictionnary[line_Dictionnary["CENTREFRAIS"].ToString()])); }
                    if (Data_Model.ExistsInDictionnary(line_dictionnary["CENTREFRAIS"].ToString(), ref CentreFrais_Dictionnary))
                    {
                        try { need_prep = need_prep && machinable_part_infos.IsPartDefault_Preparation(reference, CentreFrais_Dictionnary[line_dictionnary["CENTREFRAIS"].ToString()]); }
                        //reference_to_produce.SetFieldValue("NEED_PREP", !machinable_part_infos.IsPartDefault_Preparation(reference, CentreFrais_Dictionnary[line_Dictionnary["CENTREFRAIS"].ToString()])); }
                        catch (KeyNotFoundException)
                        {
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + "le centre de Frais  ne pointe pas vers une machine connue");
                            MessageBox.Show(string.Format("le centre de Frais {0} ne pointe pas vers une machine connue", CentreFrais_Dictionnary[line_dictionnary["CENTREFRAIS"].ToString()]));
                        }
                    }

                }
                else
                //slowmode
                //methode par comparaison d'id
                {
                    if (machine_name != "")
                    {
                        IEntityList machines = null;
                        IEntity machine = null;

                        machine = AF_ImportTools.SimplifiedMethods.GetFirtOfList(machines);
                        string mm = machinable_part_infos.DefaultMachineName;
                        need_prep = need_prep && true;
                    }
                    else
                    {
                        need_prep = need_prep && true;

                    }

                }


                ///ecriture de la piece a produire
                //les times tag permettent de savoir si la piece la piece a réelement ete importée et non crée a la main
                //ecriture du time tag
                reference_to_produce.SetFieldValue("OF_IMPORT_NUMBER", timetag.Replace("_", ""));

                reference_to_produce.SetFieldValue("_REFERENCE", reference.Id32);
                reference_to_produce.SetFieldValue("NEED_PREP", need_prep);

                //mise a jour des autres champs
                Update_Part_Item(contextlocal, ref reference_to_produce, ref CentreFrais_Dictionnary, ref line_dictionnary);
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + "update infos succeed");
                reference_to_produce.Save();
                line_dictionnary.Clear();

                return result;
            }
            catch { return result; }
        }



        /// <summary>
        /// Recupere la liste de toutes les machines sous la forme litterale "nom machine" "centre de frais"
        /// </summary>
        /// <param name="contextlocal">context local</param>
        /// <param name="Clipper_Machine">entité machine clipper</param>
        /// <param name="Clipper_Centre_Frais">entité centre de frais clipper</param>
        /// <returns></returns>
        public Boolean Get_Clipper_Machine(IContext contextlocal, out IEntity Clipper_Machine, out IEntity Clipper_Centre_Frais, out Dictionary<string, string> CentreFrais_Dictionnary)
        {



            CentreFrais_Dictionnary = new Dictionary<string, string>();
            IEntityList machine_liste = null;
            //recuperation de la machine clipper et initialisation des listes
            //CentreFrais_Dictionnary = null;
            Clipper_Machine = null;
            Clipper_Centre_Frais = null;
            //CentreFrais_Dictionnary.Clear();
            //verification que toutes les machineS sont conformes pour une intégration clipper
            ///remplissage des machines et verification de la presence du centre de frais demandé par clipper
            machine_liste = contextlocal.EntityManager.GetEntityList("_CUT_MACHINE_TYPE");
            machine_liste.Fill(false);


            foreach (IEntity machine in machine_liste)

            {
                IEntity cf;
                cf = machine.GetFieldValueAsEntity("CENTREFRAIS_MACHINE");

                if (!object.Equals(machine.GetFieldValueAsEntity("CENTREFRAIS_MACHINE"), null))
                {
                    cf = machine.GetFieldValueAsEntity("CENTREFRAIS_MACHINE");
                }
                else
                {
                    cf = null;
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": Missing  cost center on : " + machine.DefaultValue);
                    Alma_Log.Error("centre de frais non defini sur la machine  !!!" + machine.DefaultValue, MethodBase.GetCurrentMethod().Name);

                }

                ///creation du dictionnaire des machines installées   
                if (cf.DefaultValue != "" && machine.DefaultValue != "" && Clipper_Param.Get_Clipper_Machine_Cf() != null
                    )
                {
                    if (CentreFrais_Dictionnary.ContainsKey(cf.DefaultValue) == false) { CentreFrais_Dictionnary.Add(cf.DefaultValue, machine.DefaultValue); }

                    if (cf.DefaultValue == Clipper_Param.Get_Clipper_Machine_Cf())
                    {
                        if (Clipper_Param.Get_Clipper_Machine_Cf() != "Undef clipper machine")
                        {
                            Clipper_Centre_Frais = cf;
                            Clipper_Machine = machine;
                        }
                        else
                        {
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": Missing  clipper machine !!! ");
                            Alma_Log.Error("IL MANQUE LA MACHINE CLIPPER !!!", MethodBase.GetCurrentMethod().Name);
                            return false;//throw new Exception(machine.DefaultValue + " : Missing  cost center definition"); 
                        }

                    }

                }

                else
                { /*on log on arrete tout */
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": Missing  cost center definition on a machine !!! ");
                    Alma_Log.Error("IL MANQUE LE CENTRE DE FRAIS SUR L UNE DES MACHINES INSTALLEE !!!", MethodBase.GetCurrentMethod().Name);
                    return false;//throw new Exception(machine.DefaultValue + " : Missing  cost center definition"); 
                }
            }
            return true;


        }



        /// <summary>
        /// creation d'une references vide pour les nouvelles pieces qui n'ont pas encore de geometrie, il s'agit d'un contenaire
        /// </summary>
        /// <param name="contextlocal">context</param>
        /// <param name="line_dictionnary">dictionnaire de ligne</param>
        public IEntity CreateNewReference(IContext contextlocal, Dictionary<string, object> line_dictionnary, ref string NewReferenceName, IEntity clipper_machine, bool Sans_Donnees_Technique)
        {
            try
            {



                IEntity newreference = null;
                IEntity material = null;
                IEntity machine = null;

                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": creation d'une nouvelle piece !! ");
                // string referenceName = null;
                int index_extension = 0;
                //si la machine clipper n'est pas nulle
                //on initialise la machine a la machine clipper
                if (clipper_machine.Id32 != 0) { machine = clipper_machine; }

                if (line_dictionnary.ContainsKey("_MATERIAL") && line_dictionnary.ContainsKey("THICKNESS") && line_dictionnary.ContainsKey("_NAME"))
                {
                    //recuperation de la matiere 
                    material = GetMaterialEntity(contextlocal, ref line_dictionnary);
                    //recupe du nom de la geométrie                 
                    //string referenceName = "undef";    just in case mais inutiel
                    if (NewReferenceName == null)
                    {

                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": Unfortunate error: NewreferenceName does not existes and new reference has been created !! ");
                        if (Data_Model.ExistsInDictionnary("FILENAME", ref line_dictionnary))
                        {

                            NewReferenceName = line_dictionnary["FILENAME"].ToString();
                            if (NewReferenceName.ToUpper().IndexOf(".DPR.EMF") > 0) { index_extension = 7; }
                            if (NewReferenceName.ToUpper().IndexOf(".DPR") > 0) { index_extension = 4; }
                        }
                        else
                        {
                            NewReferenceName = line_dictionnary["_NAME"].ToString();
                        }

                        NewReferenceName = Path.GetFileNameWithoutExtension(@NewReferenceName.Substring(0, (NewReferenceName.Length) - index_extension));

                    }


                    ///////////////////////////recuperation de la machine envoyé par le cahier d'affaire
                    //verification de la machine sur la ligne courante
                    //affectation de la machiune clipper par defaut
                    if (line_dictionnary.ContainsKey("CENTREFRAIS") == true)
                    {
                        ///remplissage des machines et verification de la presence du centre de frais demandé par clipper
                        IEntityList machine_liste = contextlocal.EntityManager.GetEntityList("_CUT_MACHINE_TYPE");
                        machine_liste.Fill(false);

                        foreach (IEntity m in machine_liste)
                        {
                            IEntity cf = m.GetFieldValueAsEntity("CENTREFRAIS_MACHINE");
                            string cfbase = cf.GetFieldValueAsString("_CODE").ToUpper();
                            string cffile = "";
                            /*** SI CHAMP VIDE***/
                            if (line_dictionnary.ContainsKey("CENTREFRAIS") != true)
                            {
                                //recup de la machine clipper par defaut
                                machine = clipper_machine;
                            }
                            else
                            {
                                cffile = line_dictionnary["CENTREFRAIS"].ToString().ToUpper();
                                if (string.Compare(cfbase, cffile) == 0) { machine = m; break; }
                                else { machine = clipper_machine; }
                                /*on log on arrete tout */ //Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": centre de frais inconnu !! "); throw new Exception(machine.DefaultValue + " : Missing  cost center definition");


                            }


                        }
                    }
                    ///si vide alors on recupere ma machine clipper
                    else { if (clipper_machine.Id32 != 0) { machine = clipper_machine; } }


                    //

                    //creation des infos complementaires de reference notamment les données sans dt
                    //creation de l'entité
                    newreference = contextlocal.EntityManager.CreateEntity("_REFERENCE");
                    //remplacement par la machine clipper dont le cf est clip7
                    //avant tou on let la machine clipper par defaut
                    //champs standards

                    newreference.SetFieldValue("_DEFAULT_CUT_MACHINE_TYPE", machine.Id32);
                    newreference.SetFieldValue("_NAME", NewReferenceName);
                    newreference.SetFieldValue("_MATERIAL", material.Id32);

                    if (contextlocal.UserId != -1)
                    {
                        newreference.SetFieldValue("_AUTHOR", contextlocal.UserId);
                    }
                    //newreference.SetFieldValue("_AUTHOR", contextlocal.UserId); 

                    /*
                    //infos liées a l'import cfao
                    //CUSTOMER_REFERENCE_INFOS;
                    Clipper_Param.GetlistParam(contextlocal);
                    string Field_value = Clipper_Param.GetPath("CUSTOMER_REFERENCE_INFOS");
                    newreference.SetFieldValue("CUSTOMER", ""); 
                    */
                    //champs specifiques 

                    //nous retournons un minimum d'infos pour la remontée de données technique
                    if (Sans_Donnees_Technique == true || (line_dictionnary.ContainsKey("ID_PIECE_CFAO") == false))
                    {

                        if (line_dictionnary.ContainsKey("ID_PIECE_CFAO") == false) { Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": piece manquante creer a la volée, un retour clip sera necessaire !! "); }
                        else { Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": ecriture des données sans dt !! "); }
                        newreference.SetFieldValue("AFFAIRE", AF_ImportTools.SimplifiedMethods.ConvertNullStringToEmptystring("AFFAIRE", ref line_dictionnary));
                        newreference.SetFieldValue("REMONTER_DT", true);
                        newreference.SetFieldValue("_REFERENCE", AF_ImportTools.SimplifiedMethods.ConvertNullStringToEmptystring("_DESCRIPTION", ref line_dictionnary));
                        newreference.SetFieldValue("EN_RANG", AF_ImportTools.SimplifiedMethods.ConvertNullStringToEmptystring("EN_RANG", ref line_dictionnary));
                        newreference.SetFieldValue("EN_PERE_PIECE", AF_ImportTools.SimplifiedMethods.ConvertNullStringToEmptystring("EN_PERE_PIECE", ref line_dictionnary));




                    }

                    //concatenation affaire-nom-id
                    //construction d'une description de piece contenant nom matiere affaire.
                    //cette description peut etre exploitée en id d'import.

                    if (Data_Model.ExistsInDictionnary("AFFAIRE", ref line_dictionnary) && Sans_Donnees_Technique == false)
                    {
                        //concatenation dans le champs description
                        string affaire = line_dictionnary["AFFAIRE"].ToString().ToUpper();
                        string material_name = AF_ImportTools.Material.getMaterial_Name(contextlocal, material.Id32);
                        newreference.SetFieldValue("_REFERENCE", NewReferenceName + "-" + material_name + "-" + affaire);

                    }

                    newreference.Save();
                    //creation de la prepâration associée
                    AF_ImportTools.SimplifiedMethods.CreateMachinablePartFromReference(contextlocal, newreference, machine);

                }



                return newreference;

            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + " : Fails");
                System.Windows.Forms.MessageBox.Show(ie.Message);
                return null;
            }






        }



        /// <summary>
        ///N ' EST PLUS UTILIS2 utilisée
        /// controle de l'integerite des données du fichier texte
        /// on controle les champs obligatoires pour l'import, et l'existantce du centre de frais avant de continue l'import
        /// </summary>
        /// <param name="line_dictionnary">dictionnaire de ligne interprété par le datamodel</param>
        /// <returns>false ou tuue si integre</returns>
        [Obsolete]
        public Boolean CheckDataIntegerity(IContext contextlocal, Dictionary<string, object> line_dictionnary) { return true; }

        /// <summary>
        /// controle de l'integerite des données du fichier texte si cette methode retourne false : alors la ligne est ignorée
        /// on controle les champs obligatoires pour l'import, et l'existantce du centre de frais avant de continue l'import
        /// liste des verifications
        /// MATIERES INCONNUES
        /// EPAISSEURS NULLES
        /// AFFAIRES NULLES
        /// QTE NULLES
        /// CENTRE DE FRAIS INCONNU
        /// PAS DE NOMENCLATURE
        /// PAS DE NUMERO DE PIECE CLIP
        /// </summary>
        /// <param name="line_dictionnary">dictionnaire de ligne interprété par le datamodel</param>
        /// <returns>false ou tuue si integre</returns>
        public Boolean CheckDataIntegerity(IContext contextlocal, Dictionary<string, object> line_dictionnary, Dictionary<string, string> CentreFrais_Dictionnary, bool Sans_Donnees_Technique)
        {
            //
            try
            {
                ///////////////////////////////////////////////////////////////////////////
                ///condition cumulées sur result?                
                Boolean result = true;
                ///////////////////////////////////////////////////////////////////////////
                string currenfieldsname;
                ///matiere
                ///
                currenfieldsname = "_MATERIAL";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {
                    result = result & true;
                }
                else
                {
                    //MessageBox.Show(currenfieldsname + " : missing ");
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + currenfieldsname + " : missing ");
                    Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_MATERIAL"] + ": champs obligatoire : matiere non detectée sur la ligne a importée, line ignored"); result = result & false;
                    result = result & false;
                }


                ///epaisseur
                ///
                currenfieldsname = "THICKNESS";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                { result = result & true; }
                else
                {
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + currenfieldsname + " : missing ");
                    Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_THICKNESS"] + ": champs obligatoire epaisseur non detectée sur la ligne a importée, line ignored"); result = result & false;
                    result = result & false;
                }


                //L' Affaire existe t elle ?
                currenfieldsname = "AFFAIRE";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {
                    result = result & true;
                }
                else
                {
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + currenfieldsname + " : missing ");
                    Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ": champs obligatoire Affaire non detectée sur la ligne a importée, line ignored"); result = result & false;
                    result = result & false;
                }


                ///////////////////////////////////////////////////////////////////////////
                //les quantités negatives sont interdites
                currenfieldsname = "_QUANTITY";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {
                    if (int.Parse(line_dictionnary["_QUANTITY"].ToString().Trim()) < 0 || int.Parse(line_dictionnary["_QUANTITY"].ToString().Trim()) == 0)
                    {
                        Alma_Log.Error(line_dictionnary["_NAME"] + ":_QUANTITY negative ou null detecté sur la ligne a importée, line ignored", MethodBase.GetCurrentMethod().Name);
                        Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ": champs obligatoire :_QUANTITY non detecté sur la ligne a importée, line ignored"); result = false;
                    }
                }
                else
                {
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + currenfieldsname + " : missing ");
                    result = result & false;
                }

                ///////////////////////////////////////////////////////////////////////////
                //le nom de la piece à produire doit exister
                currenfieldsname = "_NAME";
                if (line_dictionnary.ContainsKey(currenfieldsname) != true)
                {
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + currenfieldsname + " : missing ");
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": champs obligatoire:  pas de nom de reference trouvée"); result = result & false;
                    Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ": champs obligatoire: pas de non de piece detecté sur la ligne a importée, line ignored"); result = result & false;
                }
                else
                {
                    result = result & true;
                }

                //////////////////////////////////////////////////////////////////////////
                //la machine (centre de frais... )
                //si la ligne ne possede pas de cf  c'est que c'est une piece sans gamme, cette piece prendre la machine clipper par defaut
                currenfieldsname = "CENTREFRAIS";
                if (line_dictionnary.ContainsKey(currenfieldsname) != true)
                {
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": pas de centre de frais --> aucune gamme detectées: piece Orange identifiée");
                    Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ": piece Orange identifiée : centre de frais non detecté sur la ligne a importée"); result = result & true;
                }
                else
                {
                    // si la machien envoyée n'existe pas on ne fait rien

                    if (Data_Model.ExistsInDictionnary(line_dictionnary[currenfieldsname].ToString(), ref CentreFrais_Dictionnary) == false)
                    {
                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": le centre de frais spécifié n'existe pas --> la ligne sera ignorée");
                        Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ":le centre de frais spécifié n'existe pas --> centre de frais non detecté sur la ligne à importée, la ligne sera ignorée");
                        result = result & false;
                    }

                    result = result & true;
                }

                ///////////////////////////////////////////////////////////////////////////
                //les matieres sont désormais obligatoires
                //string nuance_name = line_dictionnary["_MATERIAL"].ToString().Replace('§', '*');
                string nuance = null;
                string material_name = null;
                double thickness = 0;

                nuance = line_dictionnary["_MATERIAL"].ToString().Replace('§', '*');
                //thickness = line_dictionnary["THICKNESS"];
                thickness = AF_ImportTools.SimplifiedMethods.GetDoubleInvariantCulture(line_dictionnary["THICKNESS"].ToString());
                material_name = AF_ImportTools.Material.getMaterial_Name(contextlocal, nuance, thickness);

                if (material_name == string.Empty)
                { /*on log matiere non existante*/
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": matiere non existante :" + nuance + " ep " + thickness);
                    result = result & false;
                }

                ///////////////////////////////////////////////////////////////////////////
                //les matieres sont désormais obligatoires
                //pour uniquement pourles lignes jaunes ( pas pour les ligne  sans_dt)

                if (line_dictionnary.ContainsKey("IDLNROUT") != true && Sans_Donnees_Technique == false)
                {
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": champs obligatoire:  pas de numero de gamme unique indiqué"); result = result & false;
                    Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["IDLNROUT"] + ":champs obligatoire:  pas de numero de gamme unique indiqué sur la ligne a importée, line ignored"); result = result & false;
                }
                else { result = result & true; }

                if (line_dictionnary.ContainsKey("IDLNBOM") != true && Sans_Donnees_Technique == false)
                {
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": champs obligatoire:  pas d'identification unique de piece trouvée"); result = result & false;
                    Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["IDLNBOM"] + ": champs obligatoire:  pas d'identification unique de piece detecté sur la ligne a importée, line ignored"); result = result & false;
                }
                else { result = result & true; }






                return result;
            }


            catch (Exception ie)
            {
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": erreur : " + ie.Message);
                // MessageBox.Show(ie.Message);
                return false;
            }

            finally
            {
               // Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": done : " );
            }

        }
        /// <summary>
        /// renvoie l'entite matiere a partie de la nuance et de l'epaisseur contenu dans le line dictionnary
        /// </summary>
        /// <param name="contextlocal">ientity context</param>
        /// <param name="material_name">ientity  material</param>
        /// <param name="line_dictionnary">dictionnary <string,object> line_dictionnary</param>
        /// <returns></returns>
        public IEntity GetMaterialEntity(IContext contextlocal, ref Dictionary<string, object> line_dictionnary)
        {
            IEntity material = null;

            try
            {

                //IEntityList materials = null;
                //verification simple par nom nuance*etat epaisseur en rgardnat une structure comme ceci
                //"SPC*BRUT 1.00" //attention pas de control de l'obsolecence pour le moment
                if (line_dictionnary.ContainsKey("_MATERIAL") && line_dictionnary.ContainsKey("THICKNESS"))
                {
                    // material = Material.getMaterial_Entity(contextlocal, line_dictionnary["_MATERIAL"].ToString(), Convert.ToDouble(line_dictionnary["THICKNESS"]));
                    material = Material.getMaterial_Entity(contextlocal, line_dictionnary["_MATERIAL"].ToString(), SimplifiedMethods.GetDoubleInvariantCulture(line_dictionnary["THICKNESS"].ToString()));
                }





                return material;
            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": erreur :");
                MessageBox.Show(ie.Message);
                return material;
            }

        }


        /// <summary>
        /// recupere une reference en fonction d'un 
        /// numero d'indice //(remplacé par un indice d'identification piece en position 27)
        /// ce numero d'indice est egale a l'id de la piece dans la table reference sauf si l'indice est negatif.
        /// Si l'indice est negatif alors l'indice vient d'une piece cotée.
        /// 
        /// </summary>
        /// <param name="contextlocal">contexte local</param>
        /// <param name="reference">entite reference</param>
        /// <param name="line_dictionnary">dictionnaire de ligne</param>
        /// <returns>true si la reference est detectee en fonction du numero de plan</returns>
        public bool GetReference(IContext contextlocal, ref IEntity reference, ref Dictionary<string, object> line_dictionnary, ref string NewReferenceName)
        {
            reference = null;
            //IEntityList references = null;
            Int32 new_reference_id = 0;
            IEntityList quote_part_list = null;
            IEntity quote_part = null;
            //IEntity material = null;
            bool result = false;


            try
            {
                //int index_extension = 7;
                if (Data_Model.ExistsInDictionnary("ID_PIECE_CFAO", ref line_dictionnary))
                { //IEntity reference sur la base d'un id de la quotepart

                    IEntityList reference_partlist;
                    IEntity reference_part;
                    Int32 id_piece_cfao = 0;

                    if (line_dictionnary["ID_PIECE_CFAO"].GetType() == typeof(string))
                    {
                        id_piece_cfao = Convert.ToInt32(line_dictionnary["ID_PIECE_CFAO"]);
                    }
                    else
                    {
                        id_piece_cfao = (int)line_dictionnary["ID_PIECE_CFAO"];
                    }

                    //on recherche une reference cree par le devis ou alors la reference directement dans la table reference                    
                    //on regarde ensuite si le champs est negatif (--> a ce moment la c'est un quote part)post
                    if (id_piece_cfao < 0)
                    {
                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":Pièce" + id_piece_cfao + " venant d'almaquote identifiée. ");
                        //depuis le sp3 on recherche dans quote part
                        id_piece_cfao = id_piece_cfao * (-1);
                        quote_part_list = contextlocal.EntityManager.GetEntityList("_QUOTE_PART", "ID", ConditionOperator.Equal, id_piece_cfao);
                        quote_part = AF_ImportTools.SimplifiedMethods.GetFirtOfList(quote_part_list);
                        if (quote_part != null)
                        {
                            // on calcul l'id de reference sur la base du quote part
                            new_reference_id = quote_part.GetFieldValueAsInt("_ALMACAM_REFERENCE");
                            IEntityList Existreferences = contextlocal.EntityManager.GetEntityList("_REFERENCE", "ID", ConditionOperator.Equal, new_reference_id);
                            Existreferences.Fill(false);

                            //on test si la pieces existe vraiment dans les references
                            if (Existreferences.Count == 0)
                            {
                                Alma_Log.Error("REFERENCE NON TROUVEE dans LES PIECES AlmaCam : ref " + new_reference_id.ToString(), MethodBase.GetCurrentMethod().Name + "Reference non trouvée : import impossible");
                                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + id_piece_cfao + ": not found :");
                                result = false;
                            }

                        }
                        else
                        {
                            // sinon erreur et on creer une nouvelle piece
                            Alma_Log.Error("Quote Part NON TROUVEE dans les pieces de devis : ref " + id_piece_cfao.ToString(), MethodBase.GetCurrentMethod().Name + "Quotepart non trouvée : import impossible");
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + id_piece_cfao + ": not found :");
                            result = false;
                        }

                    }
                    else
                    {
                        //on recupere directment l'id de reference
                        new_reference_id = id_piece_cfao;
                    }

                    //on recupere la reference piece et on regarde si elle existe bien
                    reference_partlist = contextlocal.EntityManager.GetEntityList("_REFERENCE", "ID", ConditionOperator.Equal, new_reference_id);
                    reference_partlist.Fill(false);
                    reference_part = AF_ImportTools.SimplifiedMethods.GetFirtOfList(reference_partlist);



                    if (reference_part != null)
                    {
                        NewReferenceName = reference_part.GetFieldValueAsString("_NAME");
                        reference = reference_part;
                        result = true;
                    }
                    else
                    {
                        Alma_Log.Error("REFERENCE NON TROUVEE : ref " + id_piece_cfao.ToString(), MethodBase.GetCurrentMethod().Name + "Reference non trouvée : import impossible");
                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + id_piece_cfao + ": not found :");
                        NewReferenceName = null;
                        result = false;
                    }



                }

                else
                {

                    //reference non indiqué on cree  une nouvelle piece 
                    Alma_Log.Error("AUCUNE REFERENCE TRANSMISE : ", MethodBase.GetCurrentMethod().Name + "Reference non trouvée : crreation d'une nouvelle piece");
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":Part id_cfao   : not found :");
                    NewReferenceName = line_dictionnary["_NAME"].ToString(); ;
                    result = false;
                    //result = true;


                }

                return result;


            }

            catch (Exception ie)
            {
                //on log
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": erreur :");
                MessageBox.Show(ie.Message);
                return result;
            }



        }

        /// met a jour les valeurs   dans les pieces a produires d almacam avec les données envoyée par clipper
        /// </summary>
        /// <param name="contextlocal">contexte context</param>
        /// <param name="sheet">ientity sheet  </param>
        /// <param name="stock">inentity stock</param>
        /// <param name="line_dictionnary">dictionnary linedisctionary</param>
        /// <param name="type_tole">type tole  ou chute</param>
        /// 
        public void Update_Part_Item(IContext contextlocal, ref IEntity reference_to_produce, ref Dictionary<string, string> CentreFrais_Dictionnary, ref Dictionary<string, object> line_dictionnary)
        {
            try
            {
                foreach (var field in line_dictionnary)
                {
                    //on recupere la reference a usiner
                    //rien pour le moment
                    //on verifie que le champs existe bien avant de l'ecrire
                    if (contextlocal.Kernel.GetEntityType("_TO_PRODUCE_REFERENCE").FieldList.ContainsKey(field.Key))
                    {
                        //traitement specifique
                        switch (field.Key)
                        {

                            case "_MATERIAL":
                                //rien pour le moment mais on pourrait verifier si une nouvelle matiere est a declarer ou non
                                //recherche de l'epaisseur de la chaine 
                                //on importe jamais une matiere qui n'existe pas

                                break;

                            case "CENTREFRAIS":


                                IEntityList centre_frais = contextlocal.EntityManager.GetEntityList("_CENTRE_FRAIS", "_CODE", ConditionOperator.Equal, field.Value);
                                centre_frais.Fill(false);
                                if (centre_frais.Count() > 0)
                                {
                                    //premier de la liste ou rien
                                    reference_to_produce.SetFieldValue(field.Key, centre_frais.FirstOrDefault());
                                    if (Data_Model.ExistsInDictionnary(centre_frais.FirstOrDefault().GetFieldValueAsString("_CODE"), ref CentreFrais_Dictionnary))
                                    { reference_to_produce.SetFieldValue("MACHINE_FROM_CF", CentreFrais_Dictionnary[centre_frais.FirstOrDefault().GetFieldValueAsString("_CODE")]); }
                                    else
                                    { centre_frais.FirstOrDefault().GetFieldValueAsString("_CODE"); }

                                }
                                else
                                {
                                    reference_to_produce.SetFieldValue("MACHINE_FROM_CF", string.Format(" !!{0} pas de correspondance machine sur ce centre de frais", field.Value.ToString()));
                                    reference_to_produce.SetFieldValue("NEED_PREP", true);
                                }
                                break;



                            case "_FIRM":

                                IEntityList firmlist = contextlocal.EntityManager.GetEntityList("_FIRM", "_NAME", ConditionOperator.Equal, field.Value);
                                firmlist.Fill(false);
                                if (firmlist.Count() > 0)
                                {
                                    //premier de la liste ou rien
                                    reference_to_produce.SetFieldValue(field.Key, firmlist.FirstOrDefault().Id);


                                }

                                break;

                            case "IDLNROUT":
                                //on verifie si la reference n'exist pas deja

                                IEntityList idlnrout = contextlocal.EntityManager.GetEntityList("_TO_PRODUCE_REFERENCE", "IDLNROUT", ConditionOperator.Equal, field.Value);
                                idlnrout.Fill(false);
                                if (idlnrout.Count() == 0)
                                {
                                    reference_to_produce.SetFieldValue(field.Key, field.Value);
                                }
                                else
                                { //pas de mise à jour des quantités a produire                                         
                                    //write **_
                                    MessageBox.Show(string.Format("La gamme n° {0} a été trouvé en double, elle sera prefixé du caractère **_ pour control", field.Value));
                                    reference_to_produce.SetFieldValue(field.Key, "**_" + field.Value);
                                    //eventuellement on lance une exception
                                    //throw new InvalidDataException("doublon sur le numéro de gamme  'idlnrout' voir numero prefixé par **_"); 

                                }

                                break;

                            case "ECOQTY":

                                //formatage de La date;
                                //en cas d'erreur sur les types /// les ecoqty sont toujours en string mais dans certains base  on peut avoir l'erreur
                                
                                if (reference_to_produce.EntityType.FieldList["ECOQTY"].GetType().Name == "LongField")
                                {
                                    reference_to_produce.SetFieldValue(field.Key, int.Parse(field.Value.ToString()));
                                }
                                else
                                {
                                    reference_to_produce.SetFieldValue(field.Key, field.Value);
                                }
                                
                                //reference_to_produce.SetFieldValue(field.Key, field.Value);

                                break;

                            case "STARTDATE":
                                //formatage de La date;
                                reference_to_produce.SetFieldValue("_DATE", field.Value);
                                reference_to_produce.SetFieldValue(field.Key, field.Value);

                                break;

                            case "AF_CDE":
                                //formatage de La date;
                                reference_to_produce.SetFieldValue("_CLIENT_ORDER_NUMBER", field.Value);
                                //reference_to_produce.SetFieldValue(field.Key, field.Value);

                                break;

                            default:
                                reference_to_produce.SetFieldValue(field.Key, field.Value);
                                break;
                        }
                    }


                }

            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": erreur :");
                MessageBox.Show(ie.Message);
            }
        }


        /// <summary>
        /// METHODE PRINCIPALE
        /// les csv des of sont généralement de petite tailles un traitement à la volé dans un streamreader est possible
        /// cette fonction lit le csv 
        /// traduit chaque ligne dans un dictionnaire de ligne (interpreteur)
        /// enregistre la ligne dans almacam
        ///     rempli les champs clipper
        ///     pointe les pap vers les pieces 2d existantes
        ///     creer les nouvelles pieces (piece 2d vides) et repointe les pap 
        ///     
        /// 
        /// en standard
        /// import un of 
        /// </summary>
        /// <param name="contextlocal">contexte alma cam</param>
        /// <param name="pathToFile">chemin vers le fichier csv separateur ";"</param>
        /// <param name="sans_donnees_technique">true si import sans données techniques</param>
        /// <param name="DataModelString">string de description d'une ligne csv sous la forme 
        /// numeroIndex#NomChampAlmaCam#Type  exemple : 0#AFFAIRE#STRING</param>
        /// <summary>
        public void Import(IContext contextlocal, string pathToFile, string DataModelString, Boolean Sans_Donnees_Technique)
        {

            //recuperation des path
            CsvImportPath = pathToFile;


            try

            {
                //verification standards
                //creation du timetag d'import
                string timetag = string.Format("{0:d_M_yyyy_HH_mm_ss}", DateTime.Now);
                Alma_Log.Create_Log(Clipper_Param.GetVerbose_Log());
                //bool import_sans_donnee_technique = false;
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": importe tag :" + timetag);
                //ouverture du fichier csv lancement du curseur
                // Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor;



                if (File.Exists(CsvImportPath) == false)
                {
                    Alma_Log.Error("Fichier Non Trouvé:" + CsvImportPath, MethodBase.GetCurrentMethod().Name);
                    throw new Exception("csv File Note Found:\r\n" + CsvImportPath);
                }
                //avec ou sans dt
                if (Sans_Donnees_Technique) { Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": import  sans dt !! " + CsvImportPath); } else { Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": import standard !! " + CsvImportPath); }


                using (StreamReader csvfile = new StreamReader(CsvImportPath, Encoding.Default))
                {
                    //recuperation des elements de la base almacam
                    //declaration des dictionaires
                    Dictionary<string, object> line_Dictionnary = new Dictionary<string, object>();
                    //on lit les centres de frais 
                    ; //= null;
                    Data_Model.setFieldDictionnary(DataModelString);
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": reading data model :success !! ");
                    ///remplissage des machines et verification de la presence du centre de frais demandé par clipper
                   
                    //recuperation de la machine clipper et construction de la liste machine
                    Get_Clipper_Machine(contextlocal, out IEntity Clipper_Machine, out IEntity Clipper_Centre_Frais, out Dictionary<string, string> CentreFrais_Dictionnary);

                    //verification que toutes les machines sont conformes pour une intégration clipper


                    int ligneNumber = 0;
                    //lecture à la ligne
                    string line;
                    line = null;

                    while (((line = csvfile.ReadLine()) != null))
                    {

                        //on ne traite pas les lignes vides
                        ligneNumber++;
                        if ((line.Trim()) == "")
                        {
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + " : " + ligneNumber + ": empty line detected !! ");
                            continue;
                        }

                        //lecture des donnees
                        line_Dictionnary = Data_Model.ReadCsvLine_With_Dictionnary(line);
                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + " : " + ligneNumber + ": line disctionnary interpreter succeeded !! ");

                        //control des données    //verification des donnée importées
                        if (CheckDataIntegerity(contextlocal, line_Dictionnary, CentreFrais_Dictionnary, Sans_Donnees_Technique) != true)
                        {
                            /*on log et on continue(on passe a la ligne suivante*/
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + " : " + ligneNumber + ": data integrity fails, line ignored !!! ");
                            continue;
                        }





                        IEntity reference_to_produce = null;
                        IEntity reference = null;
                        string NewReferenceName = null;
                        //
                        //on recherche la refeence avec la bonne matiere /epaisseur si elles n'existe pas on la creer 
                        if (GetReference(contextlocal, ref reference, ref line_Dictionnary, ref NewReferenceName) == false | Sans_Donnees_Technique == true)
                        {
                            /*aucune reference n'existe dans cette matiere  alors on  la creer*/
                            /*sauf si NewReferenceName est null */
                            if (NewReferenceName != null)
                            {
                                reference = CreateNewReference(contextlocal, line_Dictionnary, ref NewReferenceName, Clipper_Machine, Sans_Donnees_Technique);
                                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + " : " + ligneNumber + " no reference found new part creation : success. ");
                                //on active le need prep
                                //need_prep = true;
                            }
                            else
                            {

                                continue;
                            }

                        }

                        //////on met a jour les données sur les piece 2d  : CUSTOMER_REFERENCE_INFOS
                        if (reference != null)
                        {///champs spécifique piece 2d
                            //infos liées a l'import cfao
                            //CUSTOMER_REFERENCE_INFOS;
                            Clipper_Param.GetlistParam(contextlocal);
                            string Field_value = Clipper_Param.GetPath("CUSTOMER_REFERENCE_INFOS");
                            Field_value.Split('|');//"CUSTOMER"
                            reference.SetFieldValue(Field_value.Split('|')[0], line_Dictionnary[Field_value.Split('|')[1]]);
                            reference.Save();
                        }




                        //creation de la nouvelle piece a produire associée
                        if (Sans_Donnees_Technique == false)
                        {
                            CreateNewPartToProduce(contextlocal, line_Dictionnary, CentreFrais_Dictionnary, ref reference_to_produce, ref reference, timetag, Sans_Donnees_Technique);
                        }
                    }



                }

                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
                File_Tools.Rename_Csv(CsvImportPath, timetag);
                Alma_Log.Final_Open_Log();
                //File_tools
            }

            catch (Exception e)
            {
                Alma_Log.Write_Log(e.Message);
                Alma_Log.Final_Open_Log();

            }
            finally
            {
                //purge
                

            }
        }

    }

    #endregion
    #region stock
    /// decrit l'import du stock
    /// <summary>
    /// recupere le stock de clipper a partir d'un fichier csv source nommé dispo mat.
    /// cet import se fait en trois phases
    /// les fichier sont beaucoup tres volumineux
    /// une premiere etape  fait la verification de l'integrité des données du csv
    /// pas de filename --> emf
    /// 
    /// Traitement des données
    /// plusieurs cas sont ensuite traité 
    ///     les toles neuves (pas de fichier emf d'apercu)
    ///     les chute possédant un idclip : les donnee sont simplement mise a jour
    ///     les nouvelles chutes : si l'emf existe et qu'il est retrouvé dans almacam, alors toutes les données de la chute sont mise a jour avec les données clip
    ///     exception sur les qté : un bolleen Isalterable = true permet de dire si les qté sont modifiables ou non
    /// Omission
    ///     les toles qui ne sont plus envoyées dans le fichier dispomat sont forcée à une quantité de 0 et marquées comme omisses
    ///     
    /// </summary>
    public class Clipper_8_Import_Stock : IDisposable
    {


        ~Clipper_8_Import_Stock()
        {
            Dispose();
        }

        //disposable
        public void Dispose()
        {
            
            GC.SuppressFinalize(this);
        }


        /// <summary>
        /// verification des données du fichier texte
        /// quantité nulle ou decimale
        ///     matiere inconnue
        ///     dim tole nulles
        ///     id_clip (identifiant unique de l'article) vide
        /// </summary>
        /// <param name="line_dictionnary">dictionnaire de ligne interprété par le datamodel</param>
        /// <returns>false ou true si integre</returns>
        public Boolean CheckDataIntegerity(IContext contextlocal, Dictionary<string, object> line_dictionnary)
        {
            //
            try
            {


                ///////////////////////////////////////////////////////////////////////////
                //les quantités negatives sont interdites
                //comptatibilité sp4
                //IEntityList materials;
                Boolean result = true;
                //string nuance_name = null;
                string currenfieldsname;

                currenfieldsname = "_QUANTITY";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {
                    if ((int)line_dictionnary[currenfieldsname] < 0)
                    {
                        Alma_Log.Error(line_dictionnary["_NAME"] + ":_QUANTITY non detectée sur la ligne a importer, line ignored", MethodBase.GetCurrentMethod().Name);
                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ":_QUANTITY non detecté sur la ligne a importée, line ignored");
                        result = false;
                    }
                }
               
                else
                {

                    currenfieldsname = "_REST_QUANTITY";
                    if (line_dictionnary.ContainsKey(currenfieldsname))
                    {

                        if ((int)line_dictionnary[currenfieldsname] < 0)
                        {
                            Alma_Log.Error("_QUANTITY non detectée sur la ligne a importer, line ignored", MethodBase.GetCurrentMethod().Name);
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + "_QUANTITY non detecté sur la ligne a importée, line ignored");
                            result = false;
                        }

                        
                    }
                   


                 }
                


                ///////////////////////////////////////////////////////////////////////////
                ///matiere exits?
               

                currenfieldsname = "_MATERIAL";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {

                    ///////////////////////////////////////////////////////////////////////////
                    //les matiere
                    //
                    string nuance = null;
                    string material_name = null;
                    double thickness = 0;

                    nuance = line_dictionnary[currenfieldsname].ToString().Replace('§', '*');
                    //thickness = Convert.ToDouble(line_dictionnary["THICKNESS"]);
                    //thickness = SimplifiedMethods.GetDoubleInvariantCulture(line_dictionnary["THICKNESS"].ToString());
                    thickness = (double)line_dictionnary["THICKNESS"];
                    material_name = AF_ImportTools.Material.getMaterial_Name(contextlocal, nuance, thickness);
                    if (material_name == string.Empty)
                    {
                        /*on log matiere non existante*/
                        Alma_Log.Error("TOLE " + line_dictionnary["IDCLIP"] + nuance + " - " + thickness + " mm : " + line_dictionnary["_NAME"] + " matiere non existante --> line ignored", MethodBase.GetCurrentMethod().Name);
                        Alma_Log.Write_Log("TOLE " + line_dictionnary["IDCLIP"] + " : " + nuance + " " + thickness + "mm :" + line_dictionnary["_NAME"] + " matiere non existante, line ignored"); result = false;
                    }
                }

                else

                { result = false; }

                ///////////////////////////////////////////////////////////////////////////
                //si pas d'idclip la ligne n'a pas d'interet
                currenfieldsname = "IDCLIP";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {
                    if (line_dictionnary[currenfieldsname].ToString() == "")
                    {
                        Alma_Log.Error(line_dictionnary[currenfieldsname] + ":IDCLIP non detecté sur la ligne a importer, line ignored", MethodBase.GetCurrentMethod().Name);
                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ":_ID non detecté sur la ligne a importer, line ignored"); result = false;
                    }
                }
                else { result = false; }
                
                ///////////////////////////////////////////////////////////////////////////
                //les longeur negative ou egales a 0  sont interdites
                currenfieldsname = "_WIDTH";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {
                    if ((double)line_dictionnary[currenfieldsname] <= 0)
                    {
                        Alma_Log.Error(line_dictionnary["_NAME"] + ": NULL _WIDTH , line ignored", MethodBase.GetCurrentMethod().Name);
                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ": NULL _WIDTH , line ignored"); result = false;
                    }
                }
                else { result = false; }
                ///////////////////////////////////////////////////////////////////////////
                //les quantités negatives sont interdites
                currenfieldsname = "_LENGTH";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {
                    if ((double)line_dictionnary[currenfieldsname] <= 0)
                    {
                        Alma_Log.Error(line_dictionnary["_NAME"] + ": NULL _LENGTH , line ignored", MethodBase.GetCurrentMethod().Name);
                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ": NULL _LENGTH , line ignored"); result = false;
                    }
                }
                else { result = false; }

                ///////////////////////////////////////////////////////////////////////////
                //les epaisseurs negatives sont interdites
                currenfieldsname = "THICKNESS";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {
                    if ((double)line_dictionnary[currenfieldsname] <= 0)
                    {

                        Alma_Log.Error(line_dictionnary["_NAME"] + ": NULL THICKNESS, line ignored", MethodBase.GetCurrentMethod().Name);
                        Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ": NULL OR NEGATIVE THICKNESS this ligne , line ignored"); result = false;
                    }
                }
                else { result = false; }

                ///////////////////////////////////////////////////////////////////////////
                //nom de l'emf
                currenfieldsname = "FILENAME";
                if (line_dictionnary.ContainsKey(currenfieldsname))
                {
                    string path_To_Emf = line_dictionnary[currenfieldsname].ToString();
                    if (File.Exists(@path_To_Emf))
                    {

                        //Alma_Log.Error(line_dictionnary["_NAME"] + ": EMF  FOUND", MethodBase.GetCurrentMethod().Name); result = true;

                    }
                    else
                    {
                        Alma_Log.Error(line_dictionnary["_NAME"] + ": EMF NOT FOUND, line ignored", MethodBase.GetCurrentMethod().Name); result = false;

                    }

                }

                return result;
            }

            catch (Exception ie)
            {
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":" + line_dictionnary["_NAME"] + ":" + ie.Message);
                return false;


            }



        }
        /// <summary>
        /// lancement de la progress barre
        /// </summary>
        /// <param name="context"></param>
        /// 

        /// <summary>
        /// main import methode
        /// </summary>
        /// <param name="param">parametre</param>
        /// <param name="pwMessage">message </param>
        public void Import(IContext contextlocal)
        {
            try
            {
                //gc collector optimized
                //GC.Collect(2, GCCollectionMode.Optimized);
                Clipper_Param.GetlistParam(contextlocal);
                //creation du timetag d'import//
                string methodename = MethodBase.GetCurrentMethod().Name;
                string timetag = string.Format("{0:d_M_yyyy_HH_mm_ss}", DateTime.Now);
                //creation du log
                bool testlog = Alma_Log.Create_Log(Clipper_Param.GetVerbose_Log());
                long ligneNumber = 0;
                Alma_Log.Write_Log(methodename + " time tag:  " + timetag);
                string CsvImportPath = Clipper_Param.GetPath("IMPORT_DM");
                Alma_Log.Write_Log("[Import du stock ]:" + CsvImportPath);
                string DataModelString = Clipper_Param.GetModelDM();
                Alma_Log.Write_Log("lecture du DataModel du stock:Success !!!");
                Alma_Log.Write_Log(" DataModel du stock valide.");
                // fin entete //

                ////chargement du datamodel

                //construction du dictionnaire de champs
                Dictionary<string, object> line_Dictionnary = new Dictionary<string, object>();
                Data_Model.setFieldDictionnary(DataModelString);
                Alma_Log.Write_Log(methodename + ": Set DataModel String success !!   ");

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ///chargement du stock avec idclip en memoire pour mise a jour - statuer les toles omises - mettre un nouveau champs //
                //
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ///risque pour la memoire // chargement des tole non ommises //
                ///notification


                using (StreamReader csvfile = new StreamReader(CsvImportPath, Encoding.Default, true))
                {
                    //IEntity Stock;
                    int etape = 0;
                    int totaletapes = 5;

                    List<string> csvfile_list_lines = new List<string>();
                    Dictionary<string, Dictionary<string, object>> csvfile_list_dictionnary = new Dictionary<string, Dictionary<string, object>>();
                    //Tuple<long, Dictionary<string, object> tuple = new Tuple<long, Dictionary<string, object>>();
                    //Tuple<int, string, bool> tuple = new Tuple<int, string, bool>(1, "cat", true);
                    //declaration de la liste des toles clipper (gpao)
                    List<string> sheetId_list_from_txt_file = new List<string>();
                    ///declaration de  la liste des toles almacam
                    List<string> sheetId_list_from_database = new List<string>();
                    //declaration de la liste des toles neuves
                    List<string> Liste_new_Toles = new List<string>();
                    //declaration de la liste des toles neuves
                    List<string> Liste_new_chutes = new List<string>();
                    // les liste ci dessous sont crées pour verification
                    ///declaration de  la liste des id de chutes envoyée par clipper
                    //List<string> verif_newchutes = new List<string>();
                    //declaration de la liste des id de toles neuves envoées par clip^per
                    //List<string> verif_newToles = new List<string>();
                    /////////////////////////////////////////////////////////////////////
                    ///declaration des query
                    ///declaration des conditions composites
                    ///("_STOCK", LogicOperator.And, 
                    ///"IDCLIP", ConditionOperator.NotEqual, string.Empty, 
                    ///"AF_IS_OMMITED", ConditionOperator.Equal, false);
                    // creation des types stock
                    //"_STOCK",  "AF_STOCK_CFAO", ConditionOperator.Equal, true, 
                    //"FILENAME", ConditionOperator.NotEqual, string.Empty
                    #region QUERY
                    IEntityType stocktype = contextlocal.Kernel.GetEntityType("_STOCK");

                    IConditionType IDCLIP_STRING_EMPTY = null;
                    IDCLIP_STRING_EMPTY = contextlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                         stocktype.ExtendedEntityType.GetExtendedField("_STOCK\\IDCLIP"),
                         ConditionOperator.NotEqual,
                    contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("IDCLIP", string.Empty));

                    //condition "IDCLIP" not empty
                    IConditionType IDCLIP_NOT_EMPTY = null;
                    IDCLIP_NOT_EMPTY = contextlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                         stocktype.ExtendedEntityType.GetExtendedField("_STOCK\\IDCLIP"),
                         ConditionOperator.NotEqual,
                         contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("IDCLIP", string.Empty));
                    //condition "IDCLIP"  empty
                    IConditionType IDCLIP_EMPTY = null;
                    IDCLIP_EMPTY = contextlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                         stocktype.ExtendedEntityType.GetExtendedField("_STOCK\\IDCLIP"),
                         ConditionOperator.Equal,
                         //contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("IDCLIP", string.Empty));
                         contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("IDCLIP", string.Empty));
                    //condition "IDCLIP" not empty
                    IConditionType IDCLIP_NOT_NULL = null;
                    IDCLIP_NOT_NULL = contextlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                         stocktype.ExtendedEntityType.GetExtendedField("_STOCK\\IDCLIP"),
                         ConditionOperator.NotEqual,
                         contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("IDCLIP", null));
                    //condition "IDCLIP"  empty
                    IConditionType IDCLIP_NULL = null;
                    IDCLIP_NULL = contextlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                         stocktype.ExtendedEntityType.GetExtendedField("_STOCK\\IDCLIP"),
                         ConditionOperator.Equal,
                         //contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("IDCLIP", string.Empty));
                         contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("IDCLIP", null));

                    //condition AF_IS_OMMITED FALSE 
                    IConditionType AF_IS_OMMITED_FALSE = null;
                    AF_IS_OMMITED_FALSE = contextlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                         stocktype.ExtendedEntityType.GetExtendedField("_STOCK\\AF_IS_OMMITED"),
                         ConditionOperator.Equal,
                         contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("AF_IS_OMMITED", false));

                    //condition AF_IS_OMMITED NULL : non setter ( a corriger dans l'actin de creation des chutes.
                    IConditionType AF_IS_OMMITED_NULL = null;
                    AF_IS_OMMITED_NULL = contextlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                         stocktype.ExtendedEntityType.GetExtendedField("_STOCK\\AF_IS_OMMITED"),
                         ConditionOperator.Equal,
                         contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("AF_IS_OMMITED", null));

                    //condition AF_STOCK_CFAO TRUE  
                    IConditionType AF_STOCK_CFAO_TRUE = null;
                    AF_STOCK_CFAO_TRUE = contextlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                         stocktype.ExtendedEntityType.GetExtendedField("_STOCK\\AF_STOCK_CFAO"),
                         ConditionOperator.Equal,
                         contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("AF_STOCK_CFAO", true));

                    //condition AF_STOCK_CFAO FALSE 
                    IConditionType AF_STOCK_CFAO_FALSE = null;
                    AF_STOCK_CFAO_FALSE = contextlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                         stocktype.ExtendedEntityType.GetExtendedField("_STOCK\\AF_STOCK_CFAO"),
                         ConditionOperator.Equal,
                         contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("AF_STOCK_CFAO", false));

                    //condition FILENAME_NOT_EMPTY
                    IConditionType FILENAME_NOT_EMPTY = null;
                    FILENAME_NOT_EMPTY = contextlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                         stocktype.ExtendedEntityType.GetExtendedField("_STOCK\\FILENAME"),
                         ConditionOperator.NotEqual,
                         contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("FILENAME", string.Empty));

                    //condition FILENAME_EMPTY
                    IConditionType FILENAME_EMPTY = null;
                    FILENAME_EMPTY = contextlocal.Kernel.ConditionTypeManager.CreateSimpleConditionType(
                         stocktype.ExtendedEntityType.GetExtendedField("_STOCK\\FILENAME"),
                         ConditionOperator.Equal,
                         contextlocal.Kernel.ConditionTypeManager.CreateConditionTypeConstantParameter("FILENAME", string.Empty));
                    ///creation des query
                    //recuperation du stock identifié
                    //condition composite 
                    IConditionType stocks_indentifie_condition_type = null;
                    stocks_indentifie_condition_type = contextlocal.Kernel.ConditionTypeManager.CreateCompositeConditionType(
                        LogicOperator.And,
                      IDCLIP_NOT_EMPTY,
                      AF_IS_OMMITED_FALSE);

                    IQuery QUERY_STOCK_IDENTIFIE = contextlocal.QueryManager.CreateQuery("_STOCK", stocks_indentifie_condition_type);

                    //recuperation du stock cfao non identifié autrement dit les nouvelles chutes;

                    IConditionType chutes_cfao_non_identifiees_en_cours_condition_type = null;
                    chutes_cfao_non_identifiees_en_cours_condition_type = contextlocal.Kernel.ConditionTypeManager.CreateCompositeConditionType(
                        LogicOperator.And,
                        //IDCLIP_EMPTY,
                        IDCLIP_NULL,
                        AF_IS_OMMITED_FALSE
                        //AF_STOCK_CFAO_TRUE//enlevé au cas ou les chutes soient d'anciennes chutes
                        );

                    IQuery QUERY_CHUTES_CFAO_NON_IDENTIFIEES = contextlocal.QueryManager.CreateQuery("_STOCK", chutes_cfao_non_identifiees_en_cours_condition_type);

                    /// recuperation des nouvelles toles pleines
                    IConditionType new_full_sheet_stock_condition_type = null;
                    new_full_sheet_stock_condition_type = contextlocal.Kernel.ConditionTypeManager.CreateCompositeConditionType(
                        LogicOperator.And,
                        IDCLIP_NOT_EMPTY,
                        AF_IS_OMMITED_FALSE
                        );

                    IQuery QUERY_NEW_FULL_SHEET_STOCK = contextlocal.QueryManager.CreateQuery("_STOCK", new_full_sheet_stock_condition_type);

                    //full_UnOmitted_stock
                    /// recuperation de tout le stock
                    /// full_sheet_stocks = contextlocal.EntityManager.GetEntityList("_STOCK", LogicOperator.And, "AF_IS_OMMITED", ConditionOperator.Equal, false, "AF_STOCK_CFAO", ConditionOperator.Equal, false, "FILENAME", ConditionOperator.Equal, string.Empty);
                    //contextlocal.EntityManager.GetEntityList("_STOCK", LogicOperator.And, "AF_IS_OMMITED", ConditionOperator.Equal, false, "AF_STOCK_CFAO", ConditionOperator.Equal, false, "FILENAME", ConditionOperator.Equal, string.Empty);
                    IConditionType new_full_stock_condition_type = null;
                    new_full_stock_condition_type = contextlocal.Kernel.ConditionTypeManager.CreateCompositeConditionType(
                        LogicOperator.And,
                        IDCLIP_NOT_EMPTY
                        //AF_IS_OMMITED_FALSE  // car il faut detecté toutes les toles pour eviter les doublons
                        //AF_STOCK_CFAO_FALSE,
                        //FILENAME_EMPTY
                        );
                    IQuery QUERY_NEW_FULL_STOCK = contextlocal.QueryManager.CreateQuery("_STOCK", new_full_stock_condition_type);

                    #endregion



                    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////// VERIFICATION DU FICHIER DE STOCK ET CHARGEMENT DE LA LISTE EN MEMOIRE //
                    /// on travaille sur les donnees possedant ds idclip
                    /// on travail sur les stock identifié
                    /// 
                    #region VERIFICATION_STOCK
                    string PHASE = "CHECKING STOCK INTEGRITY";
                    double seuil = 0.3; //afficher un message a 30% du traitement
                    etape++;
                    string etapemessage = etape.ToString() + " \\ " + totaletapes.ToString();
                    SimplifiedMethods.NotifyMessage(methodename,etapemessage+ " verification du fichier de stock...");
                    /// on liste les toles cfao dont le champs idclip est non et qui ne sont pas omises
                    long totallinenumber = 0;
                    ligneNumber = 0;
                    string newline = null;
                    Alma_Log.Write_Log(methodename + " Lecture du fichier de stock et construction du dictionnaire de ligne");
                    string[] lines = System.IO.File.ReadAllLines(CsvImportPath);
                    totallinenumber = lines.Count();
                    int step = 1;

                    //using (StreamReader csvfile = new StreamReader(CsvImportPath, Encoding.Default, true))
                    //{

                        while (!csvfile.EndOfStream)
                        {
                            ligneNumber++;
                            //affichage a traiter dans un thread
                            SimplifiedMethods.NotifyStatusMessage(methodename, etapemessage, totallinenumber, ligneNumber, ref step, seuil);

                            ///

                            newline = csvfile.ReadLine();

                            //check de la ligne si elle est vie
                            if (string.IsNullOrEmpty(newline.Trim()))
                            {
                                Alma_Log.Write_Log(methodename + PHASE + ": empty line detected  :  " + ligneNumber);
                                continue;
                            }

                            line_Dictionnary = Data_Model.ReadCsvLine_With_Dictionnary(newline);
                            object idclipobject = null;
                            line_Dictionnary.TryGetValue("IDCLIP", out idclipobject);
                            if (idclipobject != null)
                            {

                                if (CheckDataIntegerity(contextlocal, line_Dictionnary) == false)
                                {


                                    Alma_Log.Write_Log(methodename + ":-----> line " + ligneNumber + ":" + line_Dictionnary["_NAME"] + ":integrity tests fails, line ignored");
                                    continue;


                                }


                                //chargement des indexes
                                sheetId_list_from_txt_file.Add(idclipobject.ToString());
                                //chargement de la liste des lignes csv non ignorée
                                //csvfile_list_lines.Add(newline);
                                //construction du dictionnaire de lignes index
                                csvfile_list_dictionnary.Add(idclipobject.ToString(), line_Dictionnary);


                            }



                        }

                   // }
                    Alma_Log.Write_Log(methodename + PHASE + ": nombre de toles / chutes modifiables apres verification du fichier de stcok  (found)  :  " + csvfile_list_lines.Count());


                    #endregion



                    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////// STOCK IDENTIFIE //
                    /// on travaille sur les donnees possedant ds idclip
                    /// on travail sur les stock identifié
                    /// on recupere en meme temps la liste des nouvelles toles et des nouvelles chutes
                    /// dans les liste new_list_chutes et new_list_toles
                    #region TOLE_IDENTIFIEE
                    PHASE = " TOLES_IDENTIFIEES";
                    etape++;
                    etapemessage = etape.ToString() + " \\ " + totaletapes.ToString();
                    SimplifiedMethods.NotifyMessage(methodename + etapemessage, " Mise à jour du stock existant...");
                    /// on liste les toles cfao dont le champs idclip est non et qui ne sont pas omises
                    IExtendedEntityList stocks_indentifie_en_cours = null;
                    stocks_indentifie_en_cours = contextlocal.EntityManager.GetExtendedEntityList(QUERY_STOCK_IDENTIFIE);
                    stocks_indentifie_en_cours.Fill(false);


                    if(stocks_indentifie_en_cours.Count()>0)
                    {
                        Alma_Log.Write_Log(methodename + PHASE + ": stock almacam : nombre de toles identifiees et non omises  :  " + stocks_indentifie_en_cours.Count());
                    }
                    else
                    {
                        Alma_Log.Write_Log(methodename + PHASE + ": stock almacam : nombre de toles identifiees et non omises  :  0" );

                    }

                    {
                        
                        var stock_identifed_dictionnary = new Dictionary<string, IEntity>();


                        foreach (IExtendedEntity xstock in stocks_indentifie_en_cours)
                        { string idclip = null;
                            idclip = xstock.Entity.GetFieldValueAsString("IDCLIP");
                            if (stock_identifed_dictionnary.ContainsKey(idclip) == false  && string.IsNullOrEmpty(idclip) == false) {
                                stock_identifed_dictionnary.Add(idclip, xstock.Entity); }
                        }

                        Alma_Log.Write_Log_Important(methodename + ":*********************** traitement du stock existant (possedant des idclip)********************************* ");
                      




                        stocks_indentifie_en_cours = null;
                        Dispose();
                        //premiere lecture du fichier de stock clipper 

                        ligneNumber = 0;
                        step = 1;

                        foreach (string idclip in sheetId_list_from_txt_file)
                        {
                            ligneNumber++;

                            //affichage a traiter dans un thread
                            //string notification_message = etapemessage ;
                            SimplifiedMethods.NotifyStatusMessage(methodename, etapemessage, totallinenumber, ligneNumber, ref step, seuil);

                            Alma_Log.Write_Log(" +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++   LIGNE  " + ligneNumber + " +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ ");
                            //interpretation de la ligne dans le dictionnaire de ligne

                            line_Dictionnary = csvfile_list_dictionnary[idclip];

                            ///fonction classique de recherche

                            IEntity stockentity = null;

                            ///on recupere l'entité
                            //bool hasValue = stock_identifed_dictionnary.TryGetValue(idclip, out stockentity);

                            if (stock_identifed_dictionnary.TryGetValue(idclip, out stockentity) == true)
                            {
                                //recuperation du stock
                                //ici on travail les modifs.
                                ///mise a jours

                              
                                //pour sp5 //

                                Update_Stock_Item_Sp5(stockentity, line_Dictionnary);
                                Alma_Log.Write_Log(methodename + PHASE + " " + line_Dictionnary["IDCLIP"] + ":-----> line " + ligneNumber + " STOCK UPDATED.");
                                //stock = null;
                            }

                            ///traitement des toles neuves elles seront ajoutées a la fin pour 
                            ///ne pas refaire de requetes ni meme ajouter des données a la liste des toles
                            //construction de liste des toles neuves
                            //si les toles envoyées par clip  ne sont pas dans le stock 
                            else if (stockentity == null)
                            {
                                ///lecture du stock alma
                                //remplacé par la soustraction mais peut servir
                                //idclip n'existe pas dans le stock ou est omise
                                object filename = null;
                                line_Dictionnary.TryGetValue("FILENAME", out filename);

                                //
                                Alma_Log.Write_Log("La Tole/chute envoyée par clipper  " + idclip + " : " + filename + ": n'existe pas encore dans le stock");

                                if (filename == null)
                                {
                                    /////filename  vide et idclip non vide c'est une tole neuve
                                    //declaration de la liste des toles neuves
                                    Liste_new_Toles.Add(idclip);
                                    //verif_newToles.Add(idclip);
                                    Alma_Log.Write_Log("La Tole " + idclip + " sera ajoutée en phase 3  ");

                                }
                                else if (idclip != null && filename != null)
                                {
                                    ///filename non vide et idclip  vide c'est une chute neuve
                                    Liste_new_chutes.Add(idclip);
                                    Alma_Log.Write_Log("La chute " + idclip + " sera completée en phase 3  " + " : " + filename);
                                }
                                else
                                {
                                    //aucun id aucune filename --> cas non traité
                                    Alma_Log.Write_Log(ligneNumber.ToString() + " Impossible de traiter cette ligne ");
                                }
                            }



                        }
                        //purge = null;

                        Alma_Log.Write_Log(methodename + PHASE + ": nombre de toles neuves detectées dans le fichier csv  :  " + Liste_new_Toles.Count());
                        Alma_Log.Write_Log(methodename + PHASE + ": nombre de chutesnouvelles  cfao  detectées dans le fichier csv :  " + Liste_new_chutes.Count());




                        //aucun stock n'est identifié


                        //stocks_indentifie_en_cours = null;
                        stock_identifed_dictionnary.Clear();
                        stock_identifed_dictionnary = null;
                    }
                    Dispose();
                    #endregion
                    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////// CHUTE CFAO ///////////////////
                    /// mise a jour des chutes cfao
                    /// autrement dit du stock n ayant pas d'id clip mais un bitmap + stockcfao=true
                    #region CHUTE_CFAO
                    PHASE = "TRAITEMENT DES NOUVELLES CHUTES";
                    etape++;
                    etapemessage = etape.ToString() + " \\ " + totaletapes.ToString();
                    SimplifiedMethods.NotifyMessage(methodename, etapemessage + " Mise à jour des nouvelles chutes...");
                    /// on liste les toles cfao (crees par la passerelle) dont le champs idclip est vide le filname et non vide et non omises

                    IExtendedEntityList chutes_cfao_non_identifiees_en_cours = null;
                    chutes_cfao_non_identifiees_en_cours = contextlocal.EntityManager.GetExtendedEntityList(QUERY_CHUTES_CFAO_NON_IDENTIFIEES);
                    chutes_cfao_non_identifiees_en_cours.Fill(false);

                    if (chutes_cfao_non_identifiees_en_cours.Count > 0)
                    {
                        //chargement de la liste des entity dans un dictionnaire
                        //var chutes_cfao_non_identifiees_dictionnary = chutes_cfao_non_identifiees_en_cours.ToDictionary(keySelector => keySelector.Entity.GetFieldValueAsString("FILENAME"));
                        //on est obligé de gerer les doublons
                        var chutes_cfao_non_identifiees_dictionnary = new Dictionary<string, IEntity>();

                        foreach (IExtendedEntity xstock in chutes_cfao_non_identifiees_en_cours)
                        {
                            string filename = null;
                            filename = xstock.Entity.GetFieldValueAsString("FILENAME");
                            if (string.IsNullOrEmpty(filename) == false) { 
                                if (chutes_cfao_non_identifiees_dictionnary.ContainsKey(filename) == false )
                                {
                                    chutes_cfao_non_identifiees_dictionnary.Add(filename, xstock.Entity);
                                }
                            }
                        }

                        Alma_Log.Write_Log(methodename + PHASE + ": stock almacam : nombre de chutes avec idclip null (attention pas de chaines vides) , non omises, crees par l'interface cfao  " + chutes_cfao_non_identifiees_en_cours.Count());

                        chutes_cfao_non_identifiees_en_cours = null;


                        Dispose();



                        /// pas de chutes // on ignore cette etape //
                        /// on se base sur la liste des nouvelles chutes
                        if (chutes_cfao_non_identifiees_dictionnary.Count() != 0)
                        {
                            //on rembobine
                            //csvfile.BaseStream.Position = 0;
                            ligneNumber = 0;
                            totallinenumber = Liste_new_chutes.Count();
                            step = 1;
                            foreach (string idclip in Liste_new_chutes)
                            {

                                ///fonction classique de recherche
                                //IExtendedEntity stockXentity=null;
                                IEntity stockentity = null;
                                string filename = null;
                                //line = csvfile.ReadLine();
                                ligneNumber++;

                                //affichage a traiter dans un thread
                                SimplifiedMethods.NotifyStatusMessage(methodename, etapemessage, totallinenumber, ligneNumber, ref step, seuil);

                                line_Dictionnary = csvfile_list_dictionnary[idclip];

                                if (SimplifiedMethods.GetDictionnaryValue(line_Dictionnary, "FILENAME") == null)
                                {
                                    Alma_Log.Write_Log(methodename + ": no file name detected, moving next line :   for" + SimplifiedMethods.GetDictionnaryValue(line_Dictionnary, "IDCLIP").ToString());
                                    continue;
                                }
                                else
                                {
                                    filename = SimplifiedMethods.GetDictionnaryValue(line_Dictionnary, "FILENAME").ToString();
                                }
                                ///
                                ///
                                ///pas obligatoire
                                ///

                                // bool hasValue = chutes_cfao_non_identifiees_dictionnary.TryGetValue(filename, out stockentity);
                                if (chutes_cfao_non_identifiees_dictionnary.TryGetValue(filename, out stockentity) == true)
                                {   ///mise a jours
                                    Update_Stock_Item_Sp5(stockentity, line_Dictionnary);

                                }

                                else
                                { //rien
                                }



                            }

                        }
                        chutes_cfao_non_identifiees_en_cours = null;
                        chutes_cfao_non_identifiees_dictionnary.Clear();
                        chutes_cfao_non_identifiees_dictionnary = null;

                    }
                     Dispose();

                    #endregion


                    //////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////// CREATION DES NOUVELLES TOLE PLEINES NEUVES   
                    /// on travaille sur les donnees ne possedant pas de bitmap et non cfao
                    /// autrement dit les toles pleines
                    //List<string> sheetId_list_from_txt_file = new List<string>();
                    ///declaration de  la liste des toles almacam
                    //List<string> sheetId_list_from_database = new List<string>();
                    //calcul des nouvelles toles

                    #region NOUVELLE_TOLES_PLEINES
                    PHASE = "  NOUVELLE_TOLES_PLEINES";
                    etape++;
                    etapemessage = etape.ToString() + " \\ " + totaletapes.ToString();
                    SimplifiedMethods.NotifyMessage(methodename, etapemessage + " Ajout des nouvelles Toles pleines...");
                    Alma_Log.Write_Log_Important(methodename + "******************************************** TRAITEMENT NOUVELLE TOLES ********************************** ");
                    /// on liste les toles pleines du nouveau stock

                    IExtendedEntityList new_full_stocks = null;
                    bool stockneuf=false;
                    var new_full_stocks_dictionnary = new Dictionary<string, IEntity>();
                    //new_full_sheet_stocks = contextlocal.EntityManager.GetExtendedEntityList(QUERY_NEW_FULL_SHEET_STOCK);
                    new_full_stocks = contextlocal.EntityManager.GetExtendedEntityList(QUERY_NEW_FULL_STOCK);
                    new_full_stocks.Fill(false);
                    //detection d'un stockneuf 
                    if (new_full_stocks.Count()==0) { stockneuf = true;
                        Alma_Log.Write_Log(methodename + PHASE + ": stock almacam : nombre de toles nom omises, avec idclip : 0, stock neuf ");

                    }
                    else
                    {

                        Alma_Log.Write_Log(methodename + PHASE + ": stock almacam : nombre de toles nom omises, avec idclip  " + new_full_stocks.Count());

                        foreach (IExtendedEntity xstock in new_full_stocks)
                        {
                            string idclip = null;
                            idclip = xstock.Entity.GetFieldValueAsString("IDCLIP");
                            if (new_full_stocks_dictionnary.ContainsKey(idclip) == false)
                            {
                                new_full_stocks_dictionnary.Add(idclip, xstock.Entity);
                                sheetId_list_from_database.Add(idclip);
                            }
                            else
                            {
                                Alma_Log.Write_Log(methodename + PHASE + "< : doublons detectés sur idclip :   " + idclip + " : >");

                            }
                        }

                    }
                    //on est obligé de gerer les doublons
                    //un dictionnaire n'est pas vraiment necessaire ici
                    //mais cela peut servir au cas ou une information sur le doublon doivent etre ajouté au log
                    // on purge        
                    new_full_stocks = null;
                    Dispose();


               
                                      
                    //csvfile.BaseStream.Position = 0;
                    ligneNumber = 0;
                    totallinenumber = Liste_new_Toles.Count();
                    step = 1;

                     foreach ( string idclip in Liste_new_Toles)
                     //foreach (var line_Dictionnary_value in csvfile_list_dictionnary)
                    {

                        IEntity stockentity = null;
                        line_Dictionnary = null;
                        

                        ligneNumber++;
                        SimplifiedMethods.NotifyStatusMessage(methodename, etapemessage, totallinenumber, ligneNumber, ref step, seuil);
                        //recuperation des valeurs
                        line_Dictionnary = csvfile_list_dictionnary[idclip]; //line_Dictionnary_value.Value; //csvfile_list_dictionnary.Values; // csvfile_list_dictionnary[idclip];                        
                        //idclip = line_Dictionnary_value.Key;


                        if (new_full_stocks_dictionnary.TryGetValue(idclip, out stockentity)==false || stockneuf == true)
                            {
                                //Alma_Log.Write_Log(methodename + " CONFIRMATION:  AUCUNE TOLE trouver avec l'id" + idclip);
                                //Alma_Log.Write_Log(methodename + " CREATION TOLE DE LA TOLE " + idclip);
                                Create_Stock_Item_Sp5(contextlocal, line_Dictionnary);
                                Alma_Log.Write_Log(methodename + " TOLE CREEE " + idclip);
                            }
                            else {
                            Alma_Log.Write_Log(methodename + " CETTE TOLE A ETE DESACTIVEE PAR CLIPPER  " + idclip);
                        }


                        Dispose();
                    }
                    new_full_stocks = null;
                    new_full_stocks_dictionnary.Clear();
                    new_full_stocks_dictionnary = null;
                    #endregion
                    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////// TRAITEMENT DES OMISSIONS SUR TOUT LE STOCK  ////
                    /// on traite les id non envoyé par clip et on mets leurs qté à 
                    /// 
                    //List<string> sheetId_list_from_txt_file = new List<string>();
                    ///declaration de  la liste des toles almacam
                    //List<string> sheetId_list_from_database = new List<string>();
                    //calcul des omissions
                    #region OMISSION
                    etape++;
                    etapemessage = etape.ToString() + " \\ " + totaletapes.ToString();
                    SimplifiedMethods.NotifyMessage(methodename, etapemessage + " Traiement des omissions...");

                    IExtendedEntityList new_full_UnOmitted_stocks = null;

                    new_full_UnOmitted_stocks = contextlocal.EntityManager.GetExtendedEntityList(QUERY_NEW_FULL_STOCK);
                    new_full_UnOmitted_stocks.Fill(false);

                    Alma_Log.Write_Log(methodename + PHASE + ": stock almacam :   " + new_full_UnOmitted_stocks.Count());
                    Alma_Log.Write_Log_Important(methodename + "-------------------------------   OMISSION  -------------------------------------- ");
                    Alma_Log.Write_Log(methodename + ": found " + new_full_UnOmitted_stocks.Count().ToString());
                   //calcul des toles a ommettre
                    Alma_Log.Write_Log(methodename + PHASE + ": stock almacam : QTE CSV  " + sheetId_list_from_txt_file.Count());
                    Alma_Log.Write_Log(methodename + PHASE + ": stock almacam : QTE ALMACAM  " + sheetId_list_from_database.Count());
                    Alma_Log.Write_Log(methodename + PHASE + ": stock almacam : dif =  " + (sheetId_list_from_txt_file.Count() - sheetId_list_from_database.Count()));



                    if (new_full_UnOmitted_stocks.Count() > 0)
                    {
                      
                        List<string> liste_tole_a_omettre = GetOmmittedSheet(sheetId_list_from_txt_file, sheetId_list_from_database);
                        Alma_Log.Write_Log(methodename + PHASE + ": stock almacam : a omettre  " + liste_tole_a_omettre.Count());


                        if (liste_tole_a_omettre.Count > 0)
                        {

                            ligneNumber = 0;
                            step = 1;
                            totallinenumber = liste_tole_a_omettre.Count();

                            foreach (string idclip in liste_tole_a_omettre)
                            {
                                IEntity stockentity = SimplifiedMethods.GetEntityFrom_ClipId(new_full_UnOmitted_stocks, idclip.Trim());
                                //IEnumerable<IExtendedEntity> stockentityList = new_full_UnOmitted_stocks.Where(e => { return e.GetFieldValue("IDCLIP").ToString().Trim() == idclip.Trim(); });
                                ligneNumber++;

                                //affichage a traiter dans un thread
                                SimplifiedMethods.NotifyStatusMessage(methodename, etapemessage, totallinenumber, ligneNumber, ref step, seuil);




                                if (stockentity != null)
                                {
                                    Alma_Log.Write_Log(methodename + ": found and not ommitted " + idclip);

                                    //recuperation du stock

                                    //mise a jour des omission
                                    SetOmitted_Sp5(stockentity);
                                    Alma_Log.Write_Log(methodename + ": omission set " + idclip);

                                }
                                else
                                {
                                    Alma_Log.Write_Log(methodename + ":  not found " + idclip);
                                    Alma_Log.Write_Log(methodename + ": no omission set " + idclip);
                                }



                            }
                        }


                        new_full_UnOmitted_stocks = null;
                        liste_tole_a_omettre.Clear();
                        #endregion
                        }


                    //purge des listes
                    sheetId_list_from_database.Clear();
                     //chargement des indexes
                    sheetId_list_from_txt_file.Clear();
                    //chargement de la liste des lignes csv non ignorée
                    csvfile_list_lines.Clear();
                    //construction du dictionnaire de lignes indexe
                    csvfile_list_dictionnary.Clear();

                }

                ///////renommage du fichier
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
                //rename the File once imported
                //ImportTools.File_Tools.Rename_Csv(CsvImportPath);
                File_Tools.Rename_Csv(CsvImportPath, timetag);
                Alma_Log.Write_Log(methodename + " fichier   " + CsvImportPath + " renommé");
                Alma_Log.Write_Log_Important(" Stock imported successfully.");
                Alma_Log.Final_Open_Log(ligneNumber);
                ////remlpacer dans finally
                //Dispose();


                //csvfile.close

            }




            catch (Exception ie)
            {
                System.Windows.Forms.MessageBox.Show(System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ie.Message);
                
            }

            finally
            {
               // purge
                Dispose();


            }
        }





        /// <summary>
        /// met a jour les valeurs stock et sheet  dans le stock almacam
        /// attention on ne met à jour que les chutes tole qui n'ont pas de qtés reservées 
        /// </summary>
        /// <param name="contextlocal">contexte context</param>
        /// <param name="sheet">ientity sheet  </param>
        /// <param name="stock">inentity stock</param>
        /// <param name="line_dictionnary">dictionnary linedisctionary</param>
        /// <param name="type_tole">type tole  ou chute</param>
        /// n'est plus utilisé  ->02/2019
        [Obsolete]
        public void Update_Stock_Item(IEntity stock, Dictionary<string, object> line_dictionnary)
        {
            try
            {
                IContext contextlocal = stock.Context;
                foreach (var field in line_dictionnary)
                {

                    //on verifie que le champs existe bien avant de l'ecrire
                    if (contextlocal.Kernel.GetEntityType("_STOCK").FieldList.ContainsKey(field.Key))
                    {

                        //traitement specifique

                        switch (field.Key)
                        {


                            case "_WIDTH":
                                //données de sheet pas de stock
                                break;
                            case "_LENGTH":
                                break;
                            case "_MATERIAL":
                                break;
                            //case "NUMMATLOT":

                                //stock.SetFieldValue(field.Key, field.Value);
                                //stock.SetFieldValue("_HEAT_NUMBER", field.Value);
                                //break;
                              


                            case "_QUANTITY":

                                //long clipperQty = 0;
                                long Qty = 0;
                                Qty = CaclulateSheetQuantity(stock, line_dictionnary);
                                stock.SetFieldValue(field.Key, Qty);
                                //cas barou ou le client n'utilise pas les numéro de lot
                                //permet de debloquer les clotures.
                                stock.SetFieldValue("_USED_QUANTITY", 0);

                                break;


                            default:


                                stock.SetFieldValue(field.Key, field.Value);


                                break;
                        }
                    }
                }


                //on sauvegarde
                //sheet.Save();
                stock.Save();
                //creation du nom auto d'almacam
                CommonModelBuilder.ComputeSheetReference(stock.Context, stock.GetFieldValueAsEntity("_SHEET"));
                stock.GetFieldValueAsEntity("_SHEET").Save();
                stock = null;
                //Dispose();
                
            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);
            }
            finally
            {
               // Dispose();
            }
        }
              


        /// <summary>
        /// met a jour les valeurs stock et sheet  dans le stock almacam
        /// attention on ne met à jour que les chute tole qui n'ont pas de qtés reservées 
        /// </summary>
        /// <param name="contextlocal">contexte context</param>
        /// <param name="sheet">ientity sheet  </param>
        /// <param name="stock">inentity stock</param>
        /// <param name="line_dictionnary">dictionnary linedisctionary</param>
        /// <param name="type_tole">type tole  ou chute</param>
        public void Update_Stock_Item_Sp5(IEntity stock, Dictionary<string, object> line_dictionnary)
        {
            try
            {
                IContext contextlocal = stock.Context;
                foreach (var field in line_dictionnary)
                {

                    //on verifie que le champs existe bien avant de l'ecrire
                    if (contextlocal.Kernel.GetEntityType("_STOCK").FieldList.ContainsKey(field.Key))
                    {

                        //traitement specifique

                        switch (field.Key)
                        {


                            case "_WIDTH":
                                //données de sheet pas de stock
                                break;
                            case "_LENGTH":
                                break;
                            case "_MATERIAL":
                                break;
                            //case "NUMMATLOT":

                                //stock.SetFieldValue(field.Key, field.Value);
                                //stock.SetFieldValue("_HEAT_NUMBER", field.Value);
                                //break;


                            case "_QUANTITY":

                                //long clipperQty = 0;
                                long Qty = 0;

                                //Qty = CaclulateSheetQuantity(stock, line_dictionnary);
                                Qty = CalculateSheetQuantity_Sp5(stock, line_dictionnary);

                                stock.SetFieldValue(field.Key, Qty);
                                //cas barou ou le client n'utilise pas les numéro de lot
                                //permet de debloquer les clotures.
                                //stock.SetFieldValue("_USED_QUANTITY", 0);
                                break;

                            case "_REST_QUANTITY":

                                //long clipperQty = 0;
                                long Rest_Qty = 0;
                                Rest_Qty = CalculateSheetQuantity_Sp5(stock, line_dictionnary);
                                stock.SetFieldValue(field.Key, Rest_Qty);
                                //cas barou ou le client n'utilise pas les numéro de lot
                                //permet de debloquer les clotures.
                                //stock.SetFieldValue("_USED_QUANTITY", 0);

                                break;


                            default:


                                stock.SetFieldValue(field.Key, field.Value);


                                break;
                        }
                    }
                }


                //on sauvegarde

                //sheet.Save();
                stock.Save();
                //creation du nom auto d'almacam
                CommonModelBuilder.ComputeSheetReference(stock.Context, stock.GetFieldValueAsEntity("_SHEET"));
                stock.GetFieldValueAsEntity("_SHEET").Save();
                stock = null;
                Dispose();

            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);
            }
        }

        /// <summary>
        /// met a jour les valeurs stock et sheet  dans le stock almacam
        /// attention on ne met à jour que les chute tole qui n'ont pas de qtés reservées 
        /// </summary>
        /// <param name="contextlocal">contexte context</param>
        /// <param name="sheet">ientity sheet  </param>
        /// <param name="stock">inentity stock</param>
        /// <param name="line_dictionnary">dictionnary linedisctionary</param>
        /// <param name="type_tole">type tole  ou chute</param>
        /// n'est plus utilisé  ->02/2019
        [Obsolete]
        public void Create_Stock_Item(IContext contextlocal, Dictionary<string, object> line_Dictionnary)
        {
            string methodename = System.Reflection.MethodBase.GetCurrentMethod().Name;

            try
            {

                string idclip = AF_ImportTools.SimplifiedMethods.GetDictionnaryValue(line_Dictionnary, "IDCLIP").ToString();
                Alma_Log.Write_Log_Important(methodename + ": creating  stock for " + idclip);
                //construction de la liste des sheet

                IEntityList sheets_list = contextlocal.EntityManager.GetEntityList("_SHEET");
                sheets_list.Fill(false);
                //construction de la reference clip
                string sheet_to_update_reference = string.Format("{0}*{1}*{2}*{3}",

                       line_Dictionnary["_MATERIAL"].ToString().Replace('§', '*'),
                       AF_ImportTools.SimplifiedMethods.NumberToString((double)line_Dictionnary["_LENGTH"]),
                       AF_ImportTools.SimplifiedMethods.NumberToString((double)line_Dictionnary["_WIDTH"]),
                       AF_ImportTools.SimplifiedMethods.NumberToString((double)line_Dictionnary["THICKNESS"]));
                //tole pleine ou chute
                TypeTole type_tole = TypeTole.Tole;
                if (line_Dictionnary.ContainsKey("FILENAME")) { type_tole = TypeTole.Chute; }

                string nuance_name = null;
                IEntity material;
                nuance_name = line_Dictionnary["_MATERIAL"].ToString().Replace('§', '*');
                string material_name = string.Format("{0} {1:0.00} mm", nuance_name, line_Dictionnary["THICKNESS"]);
                //normaleement il n'ya pas besoin de condition car l'integrite et deja verifier dans checkdataintegrity
                material = GetMaterialEntity(contextlocal, line_Dictionnary);
                Alma_Log.Write_Log(methodename + ": material success !!    ");

                ///le sheet existe t il?
                ///
                IEntity newsheet = null;
                newsheet = AF_ImportTools.SimplifiedMethods.GetEntityFromFieldNameAsString(sheets_list, "_REFERENCE", sheet_to_update_reference.Trim());

                //sinon creation d'un nouveau sheet


                if (newsheet == null)
                {
                    ///creation des  ///
                    newsheet = contextlocal.EntityManager.CreateEntity("_SHEET");
                    newsheet.SetFieldValue("_REFERENCE", sheet_to_update_reference);
                    newsheet.SetFieldValue("_MATERIAL", material);
                    newsheet.SetFieldValue("_TYPE", (int)type_tole);

                    //newsheet.SetFieldValue("_NAME", sheet_to_update_reference);
                    double w = ((double)line_Dictionnary["_WIDTH"]);// SimplifiedMethods.GetDoubleInvariantCulture(line_Dictionnary["_WIDTH"].ToString());
                    double l = ((double)line_Dictionnary["_LENGTH"]);//((double)line_Dictionnary["_WIDTH"]).ToString()
                    newsheet.SetFieldValue("_WIDTH", w);
                    newsheet.SetFieldValue("_LENGTH", l);
                    newsheet.Complete = true;

                    //creation du nom auto d'almacam
                    CommonModelBuilder.ComputeSheetReference(newsheet.Context, newsheet);
                    newsheet.Save();
                }


                ///creation de la chute
                long sheetid = newsheet.Id;

                IEntity newstock = contextlocal.EntityManager.CreateEntity("_STOCK");
         
                newstock.SetFieldValue("_SHEET", newsheet);
                newstock.Save();
                Update_Stock_Item_Sp5(newstock, line_Dictionnary);




                newstock = null;
                newsheet = null;
                material = null;
                sheets_list = null;

                Dispose();

            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);
            }
        }

                /// <summary>
        /// met a jour les valeurs stock et sheet  dans le stock almacam
        /// attention on ne met à jour que les chute tole qui n'ont pas de qtés reservées 
        /// </summary>
        /// <param name="contextlocal">contexte context</param>
        /// <param name="sheet">ientity sheet  </param>
        /// <param name="stock">inentity stock</param>
        /// <param name="line_dictionnary">dictionnary linedisctionary</param>
        /// <param name="type_tole">type tole  ou chute</param>
        public void Create_Stock_Item_Sp5(IContext contextlocal, Dictionary<string, object> line_Dictionnary)
        {
            string methodename = System.Reflection.MethodBase.GetCurrentMethod().Name;

            try
            {

                string idclip = AF_ImportTools.SimplifiedMethods.GetDictionnaryValue(line_Dictionnary, "IDCLIP").ToString();
                Alma_Log.Write_Log_Important(methodename + ": creating  stock for " + idclip);
                //construction de la liste des sheet

                IEntityList sheets_list = contextlocal.EntityManager.GetEntityList("_SHEET");
                sheets_list.Fill(false);
                //construction de la reference clip
                
                string sheet_to_update_reference = string.Format("{0}*{1}*{2}*{3}",

                       line_Dictionnary["_MATERIAL"].ToString().Replace('§', '*'),
                       AF_ImportTools.SimplifiedMethods.NumberToString((double)line_Dictionnary["_LENGTH"]),
                       AF_ImportTools.SimplifiedMethods.NumberToString((double)line_Dictionnary["_WIDTH"]),
                       AF_ImportTools.SimplifiedMethods.NumberToString((double)line_Dictionnary["THICKNESS"])); 
                  

                //tole pleine ou chute
                TypeTole type_tole = TypeTole.Tole;
                if (line_Dictionnary.ContainsKey("FILENAME")) { type_tole = TypeTole.Chute; }

                string nuance_name = null;
                IEntity material;
                nuance_name = line_Dictionnary["_MATERIAL"].ToString().Replace('§', '*');
                string material_name = string.Format("{0} {1:0.00} mm", nuance_name, line_Dictionnary["THICKNESS"]);
                //normaleement il n'ya pas besoin de condition car l'integrite et deja verifier dans checkdataintegrity
                material = GetMaterialEntity(contextlocal, line_Dictionnary);
                Alma_Log.Write_Log(methodename + ": material success !!    ");

                ///le sheet existe t il?
                ///
                IEntity newsheet = null;
                newsheet = AF_ImportTools.SimplifiedMethods.GetEntityFromFieldNameAsString(sheets_list, "_REFERENCE", sheet_to_update_reference.Trim());

                //sinon creation d'un nouveau sheet


                if (newsheet == null)
                {
                    ///creation des  ///
                    newsheet = contextlocal.EntityManager.CreateEntity("_SHEET");
                    newsheet.SetFieldValue("_REFERENCE", sheet_to_update_reference);
                    newsheet.SetFieldValue("_MATERIAL", material);
                    newsheet.SetFieldValue("_TYPE", (int)type_tole);

                    //newsheet.SetFieldValue("_NAME", sheet_to_update_reference);
                    double w = ((double)line_Dictionnary["_WIDTH"]);// SimplifiedMethods.GetDoubleInvariantCulture(line_Dictionnary["_WIDTH"].ToString());
                    double l = ((double)line_Dictionnary["_LENGTH"]);//((double)line_Dictionnary["_WIDTH"]).ToString()
                    newsheet.SetFieldValue("_WIDTH", w);
                    newsheet.SetFieldValue("_LENGTH", l);
                    newsheet.Complete = true;
                    newsheet.Save();
                    //creation du nom auto d'almacam
                    CommonModelBuilder.ComputeSheetReference(newsheet.Context, newsheet);
                    newsheet.Save();
                }


                ///creation de la chute
                long sheetid = newsheet.Id;

                IEntity newstock = contextlocal.EntityManager.CreateEntity("_STOCK");

                newstock.SetFieldValue("_SHEET", newsheet);
                newstock.Save();
                Update_Stock_Item_Sp5(newstock, line_Dictionnary);




                newstock = null;
                newsheet = null;
                material = null;
                sheets_list = null;

                Dispose();

            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);
            }
        }

        /// <summary>
        /// les chutes / toles omisses sont les elements de stock qui ne sont plus renvoyé par clipper
        /// par exemple une tole lotie pleinement consommée dans clipper n'est plus retournée dans le fichier dispomat
        /// --
        /// cette méthode retourne la liste des chutes ommises dans le fichier en comparant les toles de qté >0
        /// avec les toles envoyées par clipper
        /// si clipper n'envoie pas le ficher la liste de ces toles sera mise a 0
        /// sheetId_list_from_txt_file, sheetId_list_from_database
        /// en l'occurence  : 
        /// toles clip = liste des toles du dispo mat
        /// toles cam = liste des toles du stock cam
        /// [tole a omettre] = [toles_clip]- [toles_cam]
        /// /// 
        /// </summary>
        /// <param name="sheetId_list_from_txt_file,">liste des toles venant du fichier impor dm </param>
        /// <param name="sheetId_list_from_database">liste des toles venant de la base clipper</param>
        /// <returns></returns>
        public List<string> GetOmmittedSheet(List<string> sheetId_list_from_txt_file, List<string> sheetId_list_from_database)
        {
            try
            {
                List<string> getOmmittedSheet = new List<string>();
                getOmmittedSheet = sheetId_list_from_database.Except(sheetId_list_from_txt_file).ToList();
                //a voir les condition qui peuvent faire qu il  ait moins de toles dans cam que dans clip
                //getOmmittedSheet = ToleImportDM.Except(ToleImportAlmaDaraBase).ToList();
                return getOmmittedSheet;


            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log(ie.Message);
                return null;
            }

        }

        /// <summary>
        /// retourne la liste des nouvelles toles / chutes
        /// 
        /// en l'occurence  : 
        /// tole neuves :  toles_cam- toles_clip
        /// 
        /// </summary>
        /// <param name="sheetId_list_from_txt_file,">liste des toles venant du fichier impor dm </param>
        /// <param name="sheetId_list_from_database">liste des toles venant de la base clipper</param>
        /// <returns></returns>
        /// n'est plus utilisé  ->02/2019
        [Obsolete]
        public List<string> GetNewStock(List<string> sheetId_list_from_txt_file, List<string> sheetId_list_from_database)
        {
            try
            {
                List<string> getnewstock = new List<string>();
                if (sheetId_list_from_database != null || sheetId_list_from_database.Count() != 0)
                {
                    getnewstock = sheetId_list_from_txt_file.Except(sheetId_list_from_database).ToList();
                }
                else
                {
                    getnewstock = sheetId_list_from_txt_file;
                }

                return getnewstock;


            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log(ie.Message);
                return null;
            }

        }


        /// <summary>
        /// calcul les quantités de toles/chutes pour assurer les clotures
        /// car les clotures ne peuvent pas se faire si les qtés dispos sont insuffisantes
        ///  on enleve les qtés qui sont utilisées dans un cycle de production
        ///     finalqty = clipper_quantity- booked;
        /// </summary>
        /// <param name="clipperqty">qtés envoyées par clipper</param>
        /// <param name="StockEntity">qtés envoyées </param>
        /// <returns></returns>
        /// n'est plus utilisé  ->02/2019
        [Obsolete]
        public long CaclulateSheetQuantity(IEntity StockEntity, Dictionary<string, object> Line_Dictionnary)
        {
            long initial = 0;
            long booked = 0;
            long used = 0;
            long format_In_Production = 0;


            //
            try
            {
                long finalqty = 0;
                long clipper_quantity = 0;
                initial = StockEntity.GetFieldValueAsLong("_QUANTITY");
                booked = StockEntity.GetFieldValueAsLong("_BOOKED_QUANTITY");
                used = StockEntity.GetFieldValueAsLong("_USED_QUANTITY");

                Alma_Log.Write_Log(" qté initiale=" + initial + " qté reservé=" + booked + " qté utilidee=" + used + " qté initiale" + "qté en prod" + format_In_Production);

                //if (clipperqty - (initial - booked ) <= 0)
                object o = AF_ImportTools.SimplifiedMethods.GetDictionnaryValue(Line_Dictionnary, "_QUANTITY");
                if (o != null)
                {
                    clipper_quantity = Convert.ToInt64(o);

                    if (Is_Alterable_Qty(StockEntity, clipper_quantity) == false)
                    //if (clipperqty - format_In_Production < 0)
                    {   ////
                        Alma_Log.Write_Log("qté calculée negative ou nulle : pas de modification du stock ");
                        finalqty = initial;

                    }
                    else
                    {
                        //on enleve les qtés qui sont utilisées dans un cycle de production
                        finalqty = clipper_quantity- booked;
                    }
                }
                Alma_Log.Write_Log("qté calculée= " + finalqty);
                return finalqty;


            }
            //
            catch (Exception ie)
            {

                Alma_Log.Write_Log("Erreur sur le calcul des quatité msg : " + ie.Message);

                return 0;
            }
            //
        }


        /// <summary>
        /// calcul les quantités de toles/chutes pour assurer les clotures
        /// car les clotures ne peuvent pas se faire si les qtés dispos sont insuffisantes
        ///  on enleve les qtés qui sont utilisées dans un cycle de production
        /// </summary>
        /// <param name="clipperqty">qtés envoyées par clipper</param>
        /// <param name="StockEntity">qtés envoyées </param>
        /// <returns></returns>
        public long CalculateSheetQuantity_Sp5(IEntity StockEntity, Dictionary<string, object> Line_Dictionnary)
        {
            long initial = 0;
            long booked = 0;
            long used = 0;
            long rest = 0;
            long format_In_Production = 0;


            //
            try
            {
                long finalqty = 0;
                long clipper_quantity = 0;
                initial = StockEntity.GetFieldValueAsLong("_QUANTITY");
                booked = StockEntity.GetFieldValueAsLong("_BOOKED_QUANTITY");
                used = StockEntity.GetFieldValueAsLong("_USED_QUANTITY");
                rest =  StockEntity.GetFieldValueAsLong("_REST_QUANTITY");

                Alma_Log.Write_Log(" qté initiale=" + initial + " qté reservé=" + booked + " qté utilidee=" + used + " qté initiale" + "qté en prod" + format_In_Production);

                //if (clipperqty - (initial - booked ) <= 0)
                object o = AF_ImportTools.SimplifiedMethods.GetDictionnaryValue(Line_Dictionnary, "_REST_QUANTITY");
                if (o != null)
                {
                    clipper_quantity = Convert.ToInt64(o);
                    //si pas de modificaitn autorisée on laisse la reste qty
                    if (Is_Alterable_Qty(StockEntity, clipper_quantity) == false)
                    //if (clipperqty - format_In_Production < 0)
                    {   ////
                        Alma_Log.Write_Log("qté calculée negative ou nulle : pas de modification du stock ");
                        finalqty = rest;

                    }
                    else
                    {

                        finalqty = clipper_quantity - booked;
                    }
                }
                Alma_Log.Write_Log("qté calculée= " + finalqty);
                return finalqty;


            }
            //
            catch (Exception ie)
            {

                Alma_Log.Write_Log("Erreur sur le calcul des quatité msg : " + ie.Message);

                return 0;
            }
            //
        }


        /// <summary>
        ///  cette methodeactive les tole omises
        /// </summary>
        /// <param name="stock">entity stock element</param>
        /// <returns>true/false si il y a eu la modification</returns>
        [Obsolete]
        public bool SetOmitted(IEntity stock)
        {
            string methodename = MethodBase.GetCurrentMethod().Name;
            try
            {
                bool rst = false;
                //-100000 on lance une requete sur les tole en cours de production 
                // on active l'omission uniquement si elle ne sont pas en prod. 
                if (Is_Alterable_Qty(stock, -100000))
                {
                    stock.SetFieldValue("AF_IS_OMMITED", true);
                    stock.SetFieldValue("_QUANTITY", 0);

                    stock.Save();
                    Alma_Log.Write_Log_Important(stock.GetFieldValueAsString("IDCLIP") + "-- OMITTED");
                    rst = true;
                }
                //sinon do nothing
                return rst;
            }
            catch (Exception ie)
            {
                System.Windows.Forms.MessageBox.Show(ie.Message, "erreur " + methodename);
                return false;
            }
        }


        /// <summary>
        ///  cette methonde rend les tole omisses
        /// </summary>
        /// <param name="stock"></param>
        /// <param name="Line_Dictonnary"></param>
        /// <returns>true/false si il y a eu la modification</returns>
        public bool SetOmitted_Sp5(IEntity stock)
        {
            string methodename = MethodBase.GetCurrentMethod().Name;
            try
            {
                bool rst = false;
                //-100000 on lance une requete sur les tole en cours de production 
                // on active l'omission uniquement si elle ne sont pas en prod. 
                if (Is_Alterable_Qty(stock, -100000))
                {
                    stock.SetFieldValue("AF_IS_OMMITED", true);
                   // stock.SetFieldValue("_QUANTITY", 0); plus utilisé depuis sp5
                    stock.SetFieldValue("_REST_QUANTITY", 0);

                    stock.Save();
                    Alma_Log.Write_Log_Important(stock.GetFieldValueAsString("IDCLIP") + "-- OMITTED");
                    rst = true;
                }
                //sinon do nothing
                return rst;
            }
            catch (Exception ie)
            {
                System.Windows.Forms.MessageBox.Show(ie.Message, "erreur " + methodename);
                return false;
            }
        }

        /// <summary>
        /// decrit la condition pour modifier les qunatitées sur un stock.
        /// ici clipper_qty - booked_qty<=0 --> renvoie false
        /// on peut forcer la modification en mettant -100000 dans les clipper quantity
        /// </summary>
        /// <param name="Stock">Ientity stock</param>
        /// <param name="Line_Dictionnary">dictionnary string objet</param>
        /// <returns>true/false</returns>
        public bool Is_Alterable_Qty(IEntity Stock, long Clipper_Quantity)
        {

            try
            {
                bool rst = true;
                string methodename = System.Reflection.MethodBase.GetCurrentMethod().Name;
                long booked_qty = Stock.GetFieldValueAsLong("_BOOKED_QUANTITY");
                string idclip = Stock.GetFieldValueAsString("IDCLIP");
                long dif = Clipper_Quantity - booked_qty;
                //long finalqty = Clipper_Quantity;
                //if (dif <= 0)
                if (dif < 0)
                    {
                    Alma_Log.Write_Log_Important(methodename + ":-----> stock clip " + idclip + ": qte clipper - quantité booked = " + dif.ToString());
                    rst = false;
                }
                //avec -100000 on recherche si le stock est en production
                if (Clipper_Quantity == -100000)
                {
                    Alma_Log.Write_Log_Important(methodename + ":-----> stock clip " + idclip + ": qté clip forcé ( CODE = -10000)  pour omission  " + dif.ToString());

                    if (booked_qty == 0)
                    {
                        Alma_Log.Write_Log_Important(methodename + ":-----> stock clip " + idclip + ": pas de qté reservée --> modification autorisée  ");

                        rst = true;
                    }
                    else
                    {
                        Alma_Log.Write_Log_Important(methodename + ":-----> stock clip " + idclip + ": pas de qté reservée --> modification interdite  ");

                        rst = false;
                    }
                }


                return rst;

            }
            catch (Exception ie)
            {
                System.Windows.Forms.MessageBox.Show(System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ie.Message);
                return false;
            }

            finally
            {

            }

        }


        /// <summary>
        /// renvoie l'entite matiere a partir du nom string 
        /// </summary>
        /// <param name="contextlocal">ientity context</param>
        /// <param name="line_dictionnary">dictionnary <string,object> line_dictionnary</param>
        /// <returns>entity de type matiere</returns>
        public IEntity GetMaterialEntity(IContext contextlocal, Dictionary<string, object> line_dictionnary)
        {
            IEntity material = null;

            try
            {

                //IEntityList materials = null;

                //verification simple par nom nuance*etat epaisseur en rgardnat une structure comme ceci
                //"SPC*BRUT 1.00" //attention pas de control de l'obsolecence pour le moment
                if (line_dictionnary.ContainsKey("_MATERIAL") && line_dictionnary.ContainsKey("THICKNESS"))
                {
                    //material = Material.getMaterial_Entity(contextlocal, line_dictionnary["_MATERIAL"].ToString(), Convert.ToDouble(line_dictionnary["THICKNESS"]));
                    //((double)line_Dictionnary["_WIDTH"]).ToString()
                    //material = Material.getMaterial_Entity(contextlocal, line_dictionnary["_MATERIAL"].ToString(), SimplifiedMethods.GetDoubleInvariantCulture(line_dictionnary["THICKNESS"].ToString()));
                    material = Material.getMaterial_Entity(contextlocal, line_dictionnary["_MATERIAL"].ToString(), ((double)line_dictionnary["THICKNESS"]));
                }


                return material;
            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log_Important(MethodBase.GetCurrentMethod().Name + ": Lecture impossible de la matiere :" + ie.Message);

                return material;
            }

        }

    }
    #endregion
    #region doonaction_retour_gp
    /// <summary>
    ///  action à la restauration d'un placement : suppression des stock cfao et suppression des fichiers
    /// </summary>
    public class Clipper_8_Before_Nesting_Restore_Event : BeforeNestingRestoreEvent
    {
        public override void OnBeforeNestingRestoreEvent(IContext context, BeforeNestingRestoreArgs args)
        {
            try
            {

                Execute(args.NestingEntity);

            }
            catch (Exception ie)
            {
                System.Windows.Forms.MessageBox.Show(System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ie.Message);

            }

            finally
            { }
        }

        public void Execute(IEntity nesting)
        {
            /*creatin du stock*/
            StockManager.DeleteNestingAssociatedStock(nesting);


        }

    }

    /// <summary>
    ///  action à l'envoi à la coupe
    ///  UNIQUEMENT A L ENVOIE
    /// </summary>
    public class Clipper_8_DoOnAction_AfterSendToWorkshop : AfterSendToWorkshopEvent
    {

        //cette fonction est lancée autant de fois qu'il y a de selection
        //la multiselection n'est pas controlée
        public override void OnAfterSendToWorkshopEvent(IContext contextlocal, AfterSendToWorkshopArgs args)
        {
            try
            {
                ///methode excecute creer le process complet de conversion des données almacam
                Execute(args.NestingEntity);
            }


            catch (Exception ie)
            {
                MessageBoxEx.ShowError(ie.Message);
            }
        }



        /// <summary>
        /// creation auto du fichier texte à  la cloture
        /// </summary>
        /// <param name="args"></param>
        public void Execute(IEntity nesting_to_cut)
        {
            try { 
            //recuperation des path
            Clipper_Param.GetlistParam(nesting_to_cut.Context);
            string export_gpao_path = Clipper_Param.GetPath("Export_GPAO");

            //ouverture des log//
            string timetag = string.Format("{0:d_M_yyyy_HH_mm_ss}", DateTime.Now);
            Alma_Log.Create_Log(Clipper_Param.GetVerbose_Log());
            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": retour gp tag :" + timetag);

                if(Clipper_Param.GetSheetAutoValidationMode())
                Alma_Log.Write_Log("ATTENTION LE MODE UTILISE : AUTO VALIDATION DES FORMATS A L ENVOI A LA COUPE N EST PAS RECOMMANDE");


             AF_ImportTools.SimplifiedMethods.NotifyMessage("export rp", "creation du fichier de retour");
            {
               
                var clipper_nest_infos = new Clipper_8_Nest_Infos();
                    // cette classe recerer une forme d'ancien nestinfos actcut pour un emeileur communicatin avec les ca
                    //on recupere la valeur par defaut du stock af normalement toujours 0
                    // cette qté sera ajoutées aux chutes crees
                    clipper_nest_infos.setdefaultAFQuantity(Clipper_Param.Parameters_Dictionnary.TryGetParam<long>("AF_ACTIVATE_QTY_ON_SENDTOWSHOP"));
                    //rempli les informations du nestinfos
                    // lance les calculs de ratios... ATTENTION CLASSE STANDARD import tools
                    //recuperation des valeurs par defaut des option cam.
                    clipper_nest_infos.Fill(nesting_to_cut,Clipper_Param.GetSheetAutoValidationMode());

                //verification de l'acces
                if (File_Tools.HasAccess(export_gpao_path)) {
                    //creation du fichier de retour pour le module de decoupe 
                    clipper_nest_infos.Export_NestInfosToFile(nesting_to_cut,export_gpao_path);
                }

                AF_ImportTools.SimplifiedMethods.NotifyMessage("export rp", "export terminé");
            }

            //fermeture des logs//
            

            }
            catch (Exception ie)
            {
                Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ":"  + ":" + ie.Message);
            }
            finally
            {
                //fermeture des logs//
                Alma_Log.Close_Log();
            }
        }





    }

    /// <summary>
    /// action apres envoie à la coupe 
    /// on supprime les toles non cfao si elle ne sont pas utilisées
    /// comme ca on enleve tous lien avec le stock almacam
    /// </summary>
    [Obsolete("PLUS UTILISE DEPUIS LA SP3 POUR CLIPPER")]
    public class Clipper_8_DoOnAction_After_Cutting_end : AfterEndCuttingEvent
    {
        //Clipper_DoOnAction_After_Cutting_end//
        public override void OnAfterEndCuttingEvent(IContext context, AfterEndCuttingArgs args)
        {
            {
                Execute(args.ToCutSheetEntity);
            }
        }


        public void Execute(IEntity ToCutSheet)
        {
            /*
            IEntity nesting = ToCutSheet.Context.EntityManager.GetEntity(ToCutSheet.GetFieldValueAsInt("_TO_CUT_NESTING"), "_NESTING");
            int nestingMultiplicity = nesting.GetFieldValueAsInt("_QUANTITY");
            StockManager.DeleteAlmaCamStock(ToCutSheet);*/
            //cloture//
            //supression du stock alma
            //IEntity nesting = ToCutSheet.Context.EntityManager.GetEntity(ToCutSheet.GetFieldValueAsInt("_TO_CUT_NESTING"), "_NESTING");
            //int nestingMultiplicity = nesting.GetFieldValueAsInt("_QUANTITY")
            Clipper_Param.GetlistParam(ToCutSheet.Context);
            int AF_ACTIVATE_QTY_ON_SENDTOWSHOP = (int)Clipper_Param.Parameters_Dictionnary.TryGetParam<long>("AF_ACTIVATE_QTY_ON_SENDTOWSHOP");

            StockManager.DeleteAlmaCamStock(ToCutSheet, AF_ACTIVATE_QTY_ON_SENDTOWSHOP);

        }
    }
    #endregion

    #region export clipper gpao : common tools

    /// <summary>
    /// creation de la classe local clipper nest_infos
    /// </summary>
    public class Clipper_8_Nest_Infos : Nest_Infos
    {

        //creation du distionnaire d'objet
        public override void Set_Stock_Specific_Fields(Tole tole)
        {
            string numatlot = tole.StockEntity.GetFieldValueAsString("NUMMATLOT");
            string numlot = tole.StockEntity.GetFieldValueAsString("NUMLOT");

            tole.Specific_Fields.Add<string>("NUMMATLOT", tole.StockEntity.GetFieldValueAsString("NUMMATLOT"));
            tole.Specific_Fields.Add<string>("NUMLOT", tole.StockEntity.GetFieldValueAsString("NUMLOT"));
        }
        public override void Set_Nested_Part_Specific_Fields(Nested_PartInfo part)
        {



            part.Specific_Fields.Add<string>("AFFAIRE", part.Part_To_Produce_IEntity.GetFieldValueAsString("AFFAIRE"));
            part.Specific_Fields.Add<string>("FAMILY", part.Part_To_Produce_IEntity.GetFieldValueAsString("FAMILY"));
            part.Specific_Fields.Add<string>("IDLNROUT", part.Part_To_Produce_IEntity.GetFieldValueAsString("IDLNROUT"));
            part.Specific_Fields.Add<string>("IDLNBOM", part.Part_To_Produce_IEntity.GetFieldValueAsString("IDLNBOM"));
            //on recherche les pieces fantomes
            if (part.Part_To_Produce_IEntity.GetFieldValueAsString("IDLNROUT") == string.Empty)
            {
                part.Part_IsGpao = false;
            }



        }
        ///export//

            ///extraction de la fiche atelier
        public string Get_WorkshopDocument(IContext context, IEntity NestingEntity, string outputpath)
        {
           
            try
            {
                string cnFileExtractDirectory = "";
                IMachineManager machineManager = new MachineManager();
                ICutMachine cutMachine = machineManager.GetCutMachine(context, NestingEntity.Id);
                ICutMachineResource cutMachineResource = cutMachine.GetCutMachineResource(true);

                //Extract the cnfile directory from the ressource of the machine
                if (string.IsNullOrEmpty(outputpath)) {
                    outputpath = cutMachineResource.GetSimpleParameterValueAsString("OUTPUT_DIRECTORY");
                AF_ImportTools.File_Tools.CreateDirectory(cnFileExtractDirectory);
                }
                else
                {
                    AF_ImportTools.File_Tools.CreateDirectory(outputpath);
                }
                //Extracted the workshop document of the nesting (named as the cn file of the nesting) as pdf in the cn file directory
                ActcutNestingManager actcutNestingManager = new ActcutNestingManager();
                actcutNestingManager.ExtractWorkshopDocumentAsPdf(context, NestingEntity, outputpath, NestingEntity.GetFieldValueAsString("_NAME"));
                return outputpath;
            }

            catch {//erreur sur le chemin renvoie un chaine vide
                        return  "";
                    }
        }
        
        public override void Export_NestInfosToFile(IEntity nesting_to_cut,string OutputDirectory)
        {
            // base.Export_NestInfosToFile(export_gpao_file);//

            /***/
            bool explodMultiplicity = Clipper_Param.Get_Multiplicity_Mode();
           
            //si la fonction est lancer par la méthode planning alors le fichier de sortie aura l'opiton .planning
            //ce pour ne pas polluer les problemes de produciton
            string extension = ".txt";
            string format = Clipper_Param.Get_string_format_double();
            string stringformatdouble = Clipper_Param.Get_string_format_double();

            ///recuperation des placements selectionnés
            ///

            switch (explodMultiplicity)
            {///normalement on conidere le premier placement
                case false:
                    {
                        MessageBox.Show(System.Reflection.MethodBase.GetCurrentMethod().Name + " Veuillez revoir les paramètres de Clipper Configuration : /r/n Le mode de multiplicité choisi dans les paramètres clipper  n'est pas géré. ");
                        break;
                    }
                case true:
                    {

                        //recuperer les infos des toles
                        foreach (Tole tole in Nesting_List_Nest_Infos_Tole_Nesting_Infos)
                        {
                            //on verifie les données du nestinfos
                            if (CheckNestingInfos() == true)
                            {
                                //definir le fichier de sortie
                                using (StreamWriter af_gpao_file = new StreamWriter(Path.Combine(OutputDirectory, tole.To_Cut_Sheet_Name + extension)))
                                { string af_gpao_file_name = Path.Combine(OutputDirectory, tole.To_Cut_Sheet_Name + extension);
                                    string Separator = ";";
                                    string[] HEADER = new string[50];
                                    string[] PART = new string[50];
                                    string[] OFFCUT = new string[50];
                                    List<string> DETAIL = new List<string>();
                                    List<string> CHUTE = new List<string>();


                                    // ecriture du fichier de sortie//
                                    // HEADER //
                                    //ecriture des entetes de nesting

                                    int iheader = 0;
                                    string headerline;
                                    tole.Specific_Fields.Get<string>("NUMMATLOT", out string NUMMATLOT);
                                    HEADER[iheader++] = "HEADER";
                                    HEADER[iheader++] = SimplifiedMethods.EmptyString(tole.Stock_Name);//nom
                                    HEADER[iheader++] = SimplifiedMethods.NumberToString(tole.Sheet_Length, format);//longeur
                                    HEADER[iheader++] = SimplifiedMethods.NumberToString(tole.Sheet_Width, format);//largeur
                                    HEADER[iheader++] = SimplifiedMethods.NumberToString(tole.Thickness, format);//epaisseur
                                    HEADER[iheader++] = SimplifiedMethods.EmptyString(tole.GradeName);//matiere
                                    HEADER[iheader++] = "1"; //mutliplicité forcée a 1 pour le moment
                                    HEADER[iheader++] = SimplifiedMethods.NumberToString((Nesting_TotalTime / 60), format); //SimplifiedMethods.NumberToString((tole.Calculus_GPAO_Parts_Total_Time/ 60),format);//temps de chargement// multiplicité toujour a 1
                                    HEADER[iheader++] = SimplifiedMethods.EmptyString(Nesting_CentreFrais_Machine);//centre de frais
                                    HEADER[iheader++] = "";//Get_WorkshopDocument(tole.Material_Entity.Context, nesting_to_cut);//";//chemin vers pdf fiche atelier
                                    HEADER[iheader++] = SimplifiedMethods.EmptyString(Nesting_Emf_File);//emf
                                    HEADER[iheader++] = SimplifiedMethods.EmptyString(NUMMATLOT);   //numero de lot
                                    HEADER[iheader++] = SimplifiedMethods.NumberToString(((Nesting_Sheet_loadingTimeEnd+ Nesting_Sheet_loadingTimeInit) / 60), format);//temps de dechargement // multiplicité toujour a 1
                                    HEADER[iheader++] = "";
                                    headerline = SimplifiedMethods.WriteTableToLine(HEADER, iheader, Separator);
                                    //
                                    //cas des toles nulles
                                    if(tole.List_Offcut_Infos.Count() == 0)
                                    { tole.StockEntity.SetFieldValue("AF_GPAO_FILE", af_gpao_file_name);
                                        tole.SheetEntity.Refresh(); // sinon plantage si on ecrit sur la meme tole en cas de mutliplicité
                                        tole.StockEntity.Save();
                                    }
                                    
                                    //
                                    // PART //

                                    foreach (Nested_PartInfo part in tole.List_Nested_Part_Infos)
                                    {
                                        int ipart = 0;
                                        string partline = "";
                                        PART[ipart++] = "DETAIL";
                                        part.Specific_Fields.Get<string>("IDLNROUT", out string idlnrout);
                                        part.Specific_Fields.Get<string>("IDLNBOM", out string idlnbom);

                                        double Part_Cutting_Time_Ratio = (part.Part_Time == 0 && tole.Calculus_GPAO_Parts_Total_Time==0) == false ? (part.Part_Time * part.Nested_Quantity / tole.Calculus_GPAO_Parts_Total_Time) : 0;

                                        double partweightcoef = part.Part_Balanced_Weight * part.Nested_Quantity / (tole.Sheet_Weight);


                                        PART[ipart++] = SimplifiedMethods.EmptyString(idlnrout);  //id unuqie clip  
                                        PART[ipart++] = SimplifiedMethods.EmptyString(idlnbom);  //id nomenclature
                                        PART[ipart++] = SimplifiedMethods.NumberToString(part.Nested_Quantity, format); // qté par placement (normalement)
                                        PART[ipart++] = SimplifiedMethods.NumberToString(part.Height, format); //hauteur 
                                        PART[ipart++] = SimplifiedMethods.NumberToString(part.Width, format); //largeur
                                        PART[ipart++] = SimplifiedMethods.NumberToString(partweightcoef, format);//poids ratio
                                        PART[ipart++] = SimplifiedMethods.NumberToString(part.Weight * 0.001, format);//poids reel en kg
                                        //PART[ipart++] = SimplifiedMethods.NumberToString(Part_Cutting_Time_Ratio, format);//poids ratio
                                        //correctif sur erreur li le ratio des temps de coupe est le meme que celui des poids
                                        //double ratio = ((part.Weight * part.Nested_Quantity) / Nesting__Gpao_Parts_Total_Weight) ;
                                        double ratio =  ((part.Weight * part.Nested_Quantity)/ Nesting__Gpao_Parts_Total_Weight)* Nesting_Multiplicity;
                                        PART[ipart++] = SimplifiedMethods.NumberToString(ratio, format);//temps de coupe//


                                        partline = SimplifiedMethods.WriteTableToLine(PART, ipart, Separator);
                                        DETAIL.Add(partline);


                                    }


                                    // OFFCUT //

                                    foreach (Tole offcut in tole.List_Offcut_Infos)
                                    {
                                        int ioffcut = 0;
                                        string offcutline = "";
                                        double longueur = offcut.Sheet_Length;
                                        double largeur = offcut.Sheet_Width;
                                        double surface = offcut.Sheet_Surface;

                                        string IsRectagular = (longueur * largeur == surface) == true ? "1" : "0";

                                        if (offcut.Sheet_Is_rotated)
                                        {
                                            longueur = offcut.Sheet_Width;
                                            largeur = offcut.Sheet_Length;
                                        }

                                        //on set les infos du fichier gp en sortie pour le supprimer avec les chutes si necessaire


                                        OFFCUT[ioffcut++] = "CHUTE";
                                        OFFCUT[ioffcut++] = SimplifiedMethods.NumberToString(longueur, format);
                                        OFFCUT[ioffcut++] = SimplifiedMethods.NumberToString(largeur, format);
                                        OFFCUT[ioffcut++] = SimplifiedMethods.NumberToString(offcut.Mutliplicity, format);
                                        OFFCUT[ioffcut++] = SimplifiedMethods.NumberToString((offcut.Sheet_Surface / Nesting_Surface), format);
                                        OFFCUT[ioffcut++] = IsRectagular;  //si la chute est care mettre 1
                                        OFFCUT[ioffcut++] = "";//chemin du dpr
                                        OFFCUT[ioffcut++] = offcut.Sheet_EmfFile;
                                        OFFCUT[ioffcut++] = SimplifiedMethods.NumberToString(offcut.Sheet_Weight/1000, format); //poids

                                        offcutline = SimplifiedMethods.WriteTableToLine(OFFCUT, ioffcut, Separator);
                                        CHUTE.Add(offcutline);
                                        //on set les infos du fichier gp en sortie pour le supprimer avec les chutes si necessaire
                                        if (offcut.StockEntity != null)
                                            offcut.StockEntity.SetFieldValue("AF_GPAO_FILE", af_gpao_file_name);
                                        offcut.StockEntity.Save();




                                        
                                        tole.Dispose();


                                    }

                                   
                                    ///ecriture
                                    #region ecriture_gpao_file
                                    ///ecriture du fichier
                                    //header
                                    af_gpao_file.WriteLine(headerline);
                                    //part
                                    foreach (string p in DETAIL)
                                    {
                                        af_gpao_file.WriteLine(p);

                                    }
                                    //offcut
                                    foreach (string c in CHUTE)
                                    {
                                        af_gpao_file.WriteLine(c);
                                    }
                                    #endregion



                                    

                                }
                            }










                        }





                        break;
                    }
            }




        }
        /// <summary>
        /// boucle de verification des données de sorties
        /// </summary>
        /// <returns></returns>
        public override bool CheckNestingInfos()
        {
            try
            {


                return true;

            }
            catch (Exception ie)
            {
                System.Windows.Forms.MessageBox.Show(System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ie.Message);
                return false;
            }

            finally
            { }

        }

    }

    /// <summary> reajustement pour retour gp
    /// on recupere avant un tole de meme matiere meme epaisseur pour calculer un ratio de poids
    /// on estime un temps de coupe
    /// on recupere les infos de pîece ..rang...
    ///Print #1, "ENGAPI;"+CodePiece+";"+Affaire+";"+Designation+";"+MonDPR+".emf"+";"+Format$(Poids,"#0.0#################")+";"+EN_RANG+";"+EN_PERE_PIECE
    ///'Print #1, "GAPIECE;"+cfrais+";;"+Format((Temps*60/100),"##0.0###############")
    ///'vb22122014 unité de temps en heure decimale (actctut est en minutes decimales)
    /// Print #1, "GAPIECE;"+cfrais+";;"+Format((Temps/60),"##0.0###############")
    ///	Print #1, "NOMENPIECE;"+CodeMat+";"+CStr(Xtole)+";"+CStr(YTole)+";"+Format(SurfPiece/SurfTole,"#0.0#################")+";"+Split(Capable," ")(0)+";"+Split(Capable," ")(1)
    ///ENGAPI;P101;3150;250 X 250;c:\alma\data\laser\formes\p102.dpr.emf;5,67555315;1;P101
    ///GAPIECE;HLASE;;0,0177742074004079
    ///NOMENPIECE;TL*S235*JR+AR*5;3200;1500;0,0304193889166667;610.78;239.06
    /// </summary>
    /// attention remonter les infos sur les nested parts
    /// <summary>
    /// impossible pour le moment car les pieces/geométries ne sont pas modifiables apres coups
    /// </summary>
    public class Clipper_8_RemonteeDt : IDisposable
    {
        public void Dispose()
        {
            ///purge
            GC.SuppressFinalize(this);
        }
        /// <summary>
        /// /// retoune la premiere tole dans une matiere donnée 
        /// </summary>
        /// <param name="contexlocal">contexte</param>
        /// <param name="part">entite reference ou piece</param>
        /// <param name="stock">entite du stock : contient l'entitté du stock trouvée</param>
        /// <param name="sheet">entitte sheet: contient la tole trouvée</param>
        /// <returns>True ou false selon si une tole est trouvée</returns>
        public bool GetRadomSheet(IContext contexlocal, IEntity part, out IEntity stock, out IEntity sheet)

        {   //on initialise sheet et stock
            sheet = null;
            stock = null;
            bool result = false;
            IEntityList sheets, stocks = null;


            try
            {
                //recherche d'element de stock
                sheets = contexlocal.EntityManager.GetEntityList("_SHEET", "_MATERIAL", ConditionOperator.Equal, part.GetFieldValueAsEntity("_MATERIAL").Id32);
                sheets.Fill(false);

                if (sheets.Count() > 0)
                {
                    foreach (IEntity sh in sheets.ToList<IEntity>())
                    {
                        if (sh.Status.ToString() == "Normal")
                        {
                            stocks = contexlocal.EntityManager.GetEntityList("_STOCK", "_SHEET", ConditionOperator.Equal, sh.Id32);
                            stock = AF_ImportTools.SimplifiedMethods.GetFirtOfList(stocks);

                            if (stocks.Count() > 0)
                            {
                                stock = stocks.FirstOrDefault();
                                if ((stock != null) & (stock.GetFieldValueAsLong("_QUANTITY") > 0))
                                {

                                    sheet = sh;
                                    break;
                                }

                            }
                        }


                    }

                }
                //ImportTools.SimplifiedMethods.GetFirtOfList(sheets);

                result = true;
                return result;
            }

            catch
            {

                Alma_Log.Write_Log(part.GetFieldValueAsString("_NAME") + ": no matierial found " + part.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsString("NAME"));
                Alma_Log.Write_Log_Important(part.GetFieldValueAsString("_NAME") + ":" + part.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsString("NAME") + " Pas de stock pour cette piece");
                return result;
            }
        }
        /// <summary>
        /// retourn true si la donnée a exportée sont valides
        /// </summary>
        /// <param name="contextlocal">icontext context</param>
        /// <param name="referenceToProduce">entité referenceToProduce</param>
        /// <returns></returns>
        public bool CheckDataIntegrity(IContext contextlocal, IEntity referenceToProduce)
        {
            return true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="Clipper_Machine"></param>
        /// <param name="Clipper_Centre_Frais"></param>
        /// <param name="CentreFrais_Dictionnary"></param>
        /// <returns>retourn la liste des machines et le centre de frais clipper : attention le centre de frais est une clé unique</returns>
        public Boolean Get_Clipper_Machine(IContext contextlocal, out IEntity Clipper_Machine, out IEntity Clipper_Centre_Frais, out Dictionary<string, string> CentreFrais_Dictionnary)
        {



            CentreFrais_Dictionnary = new Dictionary<string, string>();
            IEntityList machine_liste = null;
            //recuperation de la machine clipper et initialisation des listes
            //CentreFrais_Dictionnary = null;
            Clipper_Machine = null;
            Clipper_Centre_Frais = null;
            //CentreFrais_Dictionnary.Clear();
            //verification que toutes les machineS sont conformes pour une intégration clipper
            ///remplissage des machines et verification de la presence du centre de frais demandé par clipper
            machine_liste = contextlocal.EntityManager.GetEntityList("_CUT_MACHINE_TYPE");
            machine_liste.Fill(false);


            foreach (IEntity machine in machine_liste)
            {
                IEntity cf = machine.GetFieldValueAsEntity("CENTREFRAIS_MACHINE");
                ///creation du dictionnaire des machines installées   
                if (cf.DefaultValue != "" && machine.DefaultValue != "" && Clipper_Param.Get_Clipper_Machine_Cf() != null)
                {

                    CentreFrais_Dictionnary.Add(cf.DefaultValue, machine.DefaultValue);
                    if (cf.DefaultValue == Clipper_Param.Get_Clipper_Machine_Cf())
                    {
                        if (Clipper_Param.Get_Clipper_Machine_Cf() != "Undef clipper machine")
                        {
                            Clipper_Centre_Frais = cf;
                            Clipper_Machine = machine;
                        }
                        else
                        {
                            Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": Missing  clipper machine !!! ");
                            Alma_Log.Error("IL MANQUE LA MACHINE CLIPPER !!!", MethodBase.GetCurrentMethod().Name);
                            return false;//throw new Exception(machine.DefaultValue + " : Missing  cost center definition"); 
                        }

                    }

                }

                else
                { /*on log on arrete tout */
                    Alma_Log.Write_Log(MethodBase.GetCurrentMethod().Name + ": Missing  cost center definition on a machine !!! ");
                    Alma_Log.Error("IL MANQUE LE CENTRE DE FRAIS SUR L UNE DES MACHINES INSTALLEE !!!", MethodBase.GetCurrentMethod().Name);
                    return false;//throw new Exception(machine.DefaultValue + " : Missing  cost center definition"); 
                }
            }
            return true;


        }
        /// <summary>
        /// retourne le fichier d'echange clipper.
        ///exemple de creation du fichier de retour
        ///ENGAPI;P101;3150;250 X 250;c:\alma\data\laser\formes\p102.dpr.emf;5,67555315;1;P101
        ///GAPIECE;HLASE;;0,0177742074004079
        ///NOMENPIECE;TL*S235*JR+AR*5;3200;1500;0,0304193889166667;610.78;239.06
        /// </summary>
        /// <param name="contextlocal">contexte courant</param>

        public void Export_To_File(IContext contextlocal)
        {
            //recupere les path
            Clipper_Param.GetlistParam(contextlocal);
            string CsvExportPath = Clipper_Param.GetPath("EXPORT_Dt") + "\\DonnesTech.txt";
            //chargement de la liste de piece a retourner
            IEntitySelector select_to_produce_list = new EntitySelector();
            //on retourne les pieces marquées "sansdt"

            IEntityList sans_dt_filter = contextlocal.EntityManager.GetEntityList("_TO_PRODUCE_REFERENCE", "SANS_DT", ConditionOperator.Equal, true);
            sans_dt_filter.Fill(false);

            IDynamicExtendedEntityList references_sansdt = contextlocal.EntityManager.GetDynamicExtendedEntityList("_TO_PRODUCE_REFERENCE", sans_dt_filter);
            references_sansdt.Fill(false);


            select_to_produce_list.Init(contextlocal, references_sansdt);
            select_to_produce_list.MultiSelect = true;
            //select_to_produce_list.Init()

            //IEntityList machineList = Command.WorkOnEntityList; 

           
            
                using (StreamWriter csvfile = new StreamWriter(CsvExportPath, true))
            {


                if (select_to_produce_list.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    //on control si la matiere est la meme que la matiere precedement demandée

                    foreach (IEntity to_produce_ref in select_to_produce_list.SelectedEntity)
                    {
                        IEntity stock = null;
                        IEntity sheet = null;
                        //IEntity current_machine = null;
                        IEntity selected_reference = null;


                        IMachineManager machinemanager = new MachineManager();
                        selected_reference = to_produce_ref.GetFieldValueAsEntity("_REFERENCE");
                        //reference non vide et integrité
                        if (selected_reference != null && CheckDataIntegrity(contextlocal, to_produce_ref) == true)
                        {

                            //creation d'un part info
                            AF_ImportTools.PartInfo part_infos = new PartInfo();

                            part_infos.GetPartinfos(ref contextlocal, selected_reference);


                            //IEntity selected_reference =null;
                            //pour facilite l'ecriture du fichier de sortie on stock toutes les infos dans des listes d'objets
                            List<object> engapi = new List<object>();
                            List<object> gapiece = new List<object>();
                            List<object> nomenpiece = new List<object>();
                            string separator = ";";

                            // selected_reference = to_produce_ref.GetFieldValueAsEntity("_REFERENCE");


                            //recuperation d'une tole/stock dont la matiere est egale a celle de la  piece
                            if (GetRadomSheet(contextlocal, to_produce_ref.GetFieldValueAsEntity("_REFERENCE"), out stock, out sheet) != true)
                            {
                                Alma_Log.Write_Log_Important("aucun element de stock n' a ete trouve, seules les informations renseignées ou calculable vont etre renvoyee ");
                            };


                            //recuperation des infos de la piece
                            double surface = 0;
                            double parttime = part_infos.Quote_part_cyle_time;
                            //unité en m2
                            //surface = selected_reference.GetFieldValueAsDouble("_SURFACE") * 10E-6;
                            surface = part_infos.Surface * 10E-6;
                            string codepiece = SimplifiedMethods.ConvertNullStringToEmptystring(selected_reference.GetFieldValueAsString("_NAME"));
                            //unite en mm
                            //Double thickness = selected_reference.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsDouble("_THICKNESS");
                            double thickness = part_infos.Thickness;
                            //string matiere = selected_reference.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsString("_NAME");
                            string matiere = SimplifiedMethods.ConvertNullStringToEmptystring(part_infos.Material);
                            //double xdim = selected_reference.GetFieldValueAsDouble("_DIMENS1");
                            //double ydim = selected_reference.GetFieldValueAsDouble("_DIMENS2");
                            double xdim = part_infos.Width;
                            double ydim = part_infos.Height ;
                            //double poids = selected_reference.GetFieldValueAsDouble("_WEIGHT") * 10E-3;
                            double poids = part_infos.Weight;

                            //données clipper de la piece
                            string affaire = SimplifiedMethods.ConvertNullStringToEmptystring(to_produce_ref.GetFieldValueAsString("AFFAIRE"));
                            string description = SimplifiedMethods.ConvertNullStringToEmptystring(to_produce_ref.GetFieldValueAsString("_DESCRIPTION"));
                            //string emffile = part_infos.EmfFile;
                            string emffile;
                            if (to_produce_ref.GetFieldValueAsEntity("_REFERENCE") != null)
                            {
                                //emffile = SimplifiedMethods.ConvertNullStringToEmptystring(@to_produce_ref.GetFieldValueAsEntity("_REFERENCE").GetImageFieldValueAsLinkFile("_PREVIEW"));
                                emffile = SimplifiedMethods.GetPreview(to_produce_ref);
                            }
                            else { emffile = ""; }

                            string en_rang = SimplifiedMethods.ConvertNullStringToEmptystring(to_produce_ref.GetFieldValueAsString("EN_RANG"));
                            string en_pere_piece = SimplifiedMethods.ConvertNullStringToEmptystring(to_produce_ref.GetFieldValueAsString("EN_PERE_PIECE"));
                            string centrefrais = "";
                            if (to_produce_ref.GetFieldValueAsEntity("CENTREFRAIS") != null)
                            {
                                centrefrais = SimplifiedMethods.ConvertNullStringToEmptystring(to_produce_ref.GetFieldValueAsEntity("CENTREFRAIS").GetFieldValueAsString("_CODE"));
                            }
                            else { centrefrais = ""; }
                            //string centrefrais = SimplifiedMethods.ConvertEmptyStringToNullstring(to_produce_ref.GetFieldValueAsEntity("CENTREFRAIS").ToString());
                            string codematiere = SimplifiedMethods.ConvertNullStringToEmptystring(stock.GetFieldValueAsString("_NAME"));
                            //recuperation des infos de la tole
                            double xDim_Tole = sheet.GetFieldValueAsDouble("_LENGTH");
                            double yDim_Tole = sheet.GetFieldValueAsDouble("_WIDTH");
                            //double surface_Tole = sheet.GetFieldValueAsDouble("_SURFACE");
                            //pour le moment on considere que la surface de la tole est
                            double surface_Tole = xDim_Tole * yDim_Tole * 10E-6;
                            string nomdelaTole = SimplifiedMethods.ConvertNullStringToEmptystring(stock.GetFieldValueAsString("_NAME"));
                            //section engapi
                            engapi.Add("ENGAPI");
                            engapi.Add(codepiece);//nacleunik
                            engapi.Add(affaire);//numero_gamme
                            engapi.Add(description);
                            engapi.Add(emffile);//emf
                            engapi.Add(string.Format("{0:0.00000}", poids));
                            engapi.Add(en_rang);
                            engapi.Add(en_pere_piece);
                            //engapi.Add(indice_c);
                            //section gapiece, données clipper de la gamme
                            gapiece.Add("GAPIECE");
                            gapiece.Add(centrefrais);
                            //pour le moment ces valeurs sont nulles
                            gapiece.Add(string.Format("{0:0.00000}", parttime));
                            gapiece.Add("");
                            ///
                            nomenpiece.Add("NOMEPIECE");
                            nomenpiece.Add(codematiere);
                            nomenpiece.Add(xDim_Tole);
                            nomenpiece.Add(yDim_Tole);

                            if (surface_Tole > 0)
                            {
                                nomenpiece.Add(string.Format("{0:0.00000}", surface / surface_Tole));
                            }

                            nomenpiece.Add(xdim);
                            nomenpiece.Add(ydim);



                            string myline = "";
                            foreach (object o in engapi)
                            {
                                myline += o.ToString() + separator;
                            }
                            csvfile.WriteLine(myline);

                            //ecriture de gapiece
                            myline = "";
                            foreach (object o in gapiece)
                            {
                                myline += o.ToString() + separator;
                            }
                            csvfile.WriteLine(myline);
                            myline = "";
                            //ecriture de nomepiece
                            foreach (object o in nomenpiece)
                            {
                                myline += o.ToString() + separator;
                            }
                            csvfile.WriteLine(myline);
                            //on set le setfiled value a false

                            to_produce_ref.SetFieldValue("SANS_DT", false);
                            to_produce_ref.Save();
                            part_infos = null;

                        }
                    }

                }
                csvfile.Close();

            }


        }
        /// <summary>
        /// export un dossier technique
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="EngapiOnly">si false alors, la ligne de gamme n'est pas exportée</param>
        public void Export_Piece_To_File(IContext contextlocal, bool EngapiOnly)
        {
            //bool export ligne de gamme
            //bool export_ 
            //var answer= MessageBox.Show.   
            Clipper_Param.GetlistParam(contextlocal);

            EngapiOnly = false;
            //recupere les path

            Clipper_Param.GetlistParam(contextlocal);
            string CsvExportPath = Clipper_Param.GetPath("EXPORT_DT") + "\\DonnesTech.txt";
            //chargement de la liste de piece a retourner
            IEntitySelector select_machinable_part_list = new EntitySelector();
            IEntitySelector select_preparation_list = new EntitySelector();

            

            Dictionary<string, string> centre_frais_dictionnary = null;
            centre_frais_dictionnary = new Dictionary<string, string>();
            //machine clipper
            Get_Clipper_Machine(contextlocal, out IEntity clipper_machine, out IEntity centre_frais_clipper, out centre_frais_dictionnary);

           


            IEntityList sans_dt_filter = contextlocal.EntityManager.GetEntityList("_MACHINABLE_PART", "_ESTIMATION_IS_VALID", ConditionOperator.Equal, true);
            sans_dt_filter.Fill(false);


            IDynamicExtendedEntityList machinableparts = contextlocal.EntityManager.GetDynamicExtendedEntityList("_MACHINABLE_PART", sans_dt_filter);
            machinableparts.Fill(false);


            // select_reference_list.Init(contextlocal, references_sansdt);
            select_machinable_part_list.AllowSaveView=true;
            select_machinable_part_list.Init(contextlocal, machinableparts);
            select_machinable_part_list.MultiSelect = true;
            //
            MessageBox.Show("Seules les preparations valides pour l'export seront visibles", "Export Pieces 2d ", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            //GetPartinfos_FromMachinablePart
            //ecriture du fichier remplacement
            File_Tools.replaceFile(CsvExportPath, false);

            using (StreamWriter csvfile = new StreamWriter(CsvExportPath, true))
            {


                if (select_machinable_part_list.ShowDialog() == System.Windows.Forms.DialogResult.OK)

                {
                    //on control si la matiere est la meme que la matiere precedement demandée

                    foreach (IEntity machinable_part in select_machinable_part_list.SelectedEntity)
                    {
                        IEntity stock = null;
                        IEntity sheet = null;
                        //IEntity current_machine = null;

                        //IEntity to_produce_ref = null;
                        IMachineManager machinemanager = new MachineManager();

                        //reference non vide et integrité
                        if (machinable_part != null && CheckDataIntegrity(contextlocal, machinable_part) == true)
                        {
                            //

                            //IEntity current_machine = null;
                            string strresult = "";
                            //bool result = false;
                            long affaire;
                            string name = machinable_part.GetFieldValueAsEntity("IMPLEMENT__PREPARATION").GetFieldValueAsString("_NAME");
                            IEntity reference = machinable_part.GetImplementEntity("_PREPARATION").GetFieldValueAsEntity("_REFERENCE");

                            SimplifiedMethods.InputBox("Saissisez le numéro d'affaire", "Piece :" + name, ref strresult);

                            if (string.IsNullOrEmpty(strresult) == true || long.TryParse(strresult, out affaire) == false)//i now = 1 )
                            {
                                //throw new MissingAffaireExecption();
                                MessageBox.Show(name + " : Numero d'affaire incorrecte, cette piece sera ignorée.");
                                //Application.Exit();
                                continue;
                            }
                            else
                            {

                                //creation d'un part info

                                AF_ImportTools.PartInfo part_infos = new PartInfo();
                                part_infos.GetPartinfos_FromMachinablePart(ref contextlocal, machinable_part);


                                //pour facilite l'ecriture du fichier de sortie on stock toutes les infos dans des listes d'objets
                                List<object> engapi = new List<object>();
                                List<object> gapiece = new List<object>();
                                List<object> nomenpiece = new List<object>();
                                string separator = ";";

                                //recuperation d'une tole/stock dont la matiere est egale a celle de la  piece
                                if (GetRadomSheet(contextlocal, reference, out stock, out sheet) != true)
                                {
                                    Alma_Log.Write_Log_Important("aucun element de stock n' a ete trouve, seules les informations renseignées ou calculables vont etre renvoyees ");
                                };


                                //recuperation des infos de la piece
                                double surface = 0;
                                double parttime = part_infos.Quote_part_cyle_time + 0.1;
                                //unité en m2
                                //surface = selected_reference.GetFieldValueAsDouble("_SURFACE") * 10E-6;
                                surface = part_infos.Surface * 10E-6;
                                string codepiece = SimplifiedMethods.ConvertNullStringToEmptystring(reference.GetFieldValueAsString("_NAME"));
                                //unite en mm

                                double thickness = part_infos.Thickness;
                                //string matiere = selected_reference.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsString("_NAME");
                                string matiere = SimplifiedMethods.ConvertNullStringToEmptystring(part_infos.Material);
                                double feed = AF_ImportTools.Machine_Info.GetFeed(contextlocal, part_infos.DefaultMachineEntity, part_infos.MaterialEntity);

                                //
                                //double xdim = part_infos.Width;
                                //double ydim = part_infos.Height;


                                double xdim;
                                double ydim;
                                //longeur > largeur
                                if (part_infos.Width > part_infos.Height)
                                {
                                    xdim = part_infos.Width;
                                    ydim = part_infos.Height;

                                }
                                else
                                {
                                    ydim = part_infos.Width;
                                    xdim = part_infos.Height;
                                }



                                double poids = part_infos.Weight;

                                //données clipper de la piece

                                //string affaire = ""; //SimplifiedMethods.ConvertNullStringToEmptystring(reference.GetFieldValueAsString("AFFAIRE"));
                                string description = "Piece de dim : " + String.Format("{0:0.00}", xdim) + " mm" + " X " + String.Format("{0:0.00}", ydim) + " mm";//part_infos.Name;// SimplifiedMethods.ConvertNullStringToEmptystring(reference.GetFieldValueAsString("_REFERENCE"));


                                string emffile;
                                /// a revoir avec les grometrie par defaut
                                emffile = part_infos.EmfFile;//SimplifiedMethods.GetPreview(reference);
                                string en_rang = ""; // SimplifiedMethods.ConvertNullStringToEmptystring(reference.GetFieldValueAsString("EN_RANG"));
                                string en_pere_piece = ""; // SimplifiedMethods.ConvertNullStringToEmptystring(reference.GetFieldValueAsString("EN_PERE_PIECE"));
                                                           //string idlnbom = "";// SimplifiedMethods.ConvertNullStringToEmptystring(reference.GetFieldValueAsString("IDLNBOM"));

                                string referenceId = reference.Id.ToString(); ;  // reference.Id32.ToString();
                                string centrefrais = "";
                                centrefrais = SimplifiedMethods.ConvertNullStringToEmptystring(reference.GetFieldValueAsEntity("_DEFAULT_CUT_MACHINE_TYPE").GetFieldValueAsEntity("CENTREFRAIS_MACHINE").GetFieldValueAsString("_CODE"));

                                //string centrefrais = SimplifiedMethods.ConvertEmptyStringToNullstring(to_produce_ref.GetFieldValueAsEntity("CENTREFRAIS").ToString());
                                string codematiere = SimplifiedMethods.ConvertNullStringToEmptystring(stock.GetFieldValueAsString("_NAME"));
                                //recuperation des infos de la tole
                                double xDim_Tole = sheet.GetFieldValueAsDouble("_LENGTH");
                                double yDim_Tole = sheet.GetFieldValueAsDouble("_WIDTH");
                                //double surface_Tole = sheet.GetFieldValueAsDouble("_SURFACE");
                                //pour le moment on considere que la surface de la tole est
                                double surface_Tole = xDim_Tole * yDim_Tole * 10E-6;
                                string nomdelaTole = SimplifiedMethods.ConvertNullStringToEmptystring(stock.GetFieldValueAsString("_NAME"));
                                //section engapi

                                engapi.Add("ENGAPI");
                                engapi.Add(codepiece);
                                engapi.Add(affaire.ToString());
                                engapi.Add(description);
                                engapi.Add(emffile);
                                engapi.Add(string.Format("{0:0.00000}", poids / 1000));
                                //engapi.Add(idlnbom);
                                engapi.Add(en_rang);
                                engapi.Add(en_pere_piece);
                                engapi.Add(referenceId);
                                //var answer= MessageBox.Show.   
                                if (EngapiOnly == false)
                                {
                                    gapiece.Add("GAPIECE");
                                    gapiece.Add(centrefrais);
                                    //pour le moment ces valeurs sont nulles
                                    //String.Format(Clipper_Param.Get_string_format_double(), clipperpart.Part_Balanced_Weight * clipperpart.Nested_Quantity
                                    gapiece.Add(string.Format("{0:0.0000}", 0));
                                    gapiece.Add(string.Format("{0:0.0000}", parttime));
                                    //gapiece.Add("");
                                    ///
                                }
                                nomenpiece.Add("NOMENPIECE");
                                nomenpiece.Add(codematiere);
                                nomenpiece.Add(xDim_Tole);
                                nomenpiece.Add(yDim_Tole);


                                if (surface_Tole > 0)
                                {
                                    nomenpiece.Add(string.Format("{0:0.00000}", surface / surface_Tole));
                                }

                                nomenpiece.Add(xdim);
                                nomenpiece.Add(ydim);



                                string myline = "";
                                var last_o = engapi.LastOrDefault();
                                separator = ";";
                                foreach (object o in engapi)
                                {
                                    if (o.Equals(last_o))
                                    { separator = ""; }
                                    myline += o.ToString() + separator;
                                }
                                last_o = null;


                                csvfile.WriteLine(myline.Replace(",", "."));

                                //ecriture de gapiece uniquement si retour possible
                                //actuellement si la gamme existe clipper rajoute une nouvelle gamme
                                //GAMME
                                if (EngapiOnly == false)
                                {
                                    myline = "";
                                    last_o = gapiece.LastOrDefault();
                                    separator = ";";

                                    foreach (object o in gapiece)
                                    {
                                        if (o.Equals(last_o)) { separator = ""; }
                                        myline += o.ToString() + separator;
                                    }
                                    csvfile.WriteLine(myline.Replace(",", "."));
                                }

                                //ecriture de nomepiece
                                //NOMENPIECE
                                myline = "";
                                last_o = nomenpiece.LastOrDefault();
                                separator = ";";

                                foreach (object o in nomenpiece)
                                {
                                    if (o.Equals(last_o)) { separator = ""; }
                                    myline += o.ToString() + separator;
                                }

                                csvfile.WriteLine(myline.Replace(",", "."));
                                last_o = null;
                                //on set le setfiled value a false

                                //reference.SetFieldValue("REMONTER_DT", false);
                                //reference.Save();
                                part_infos = null;



                            }

                        



                        





                        }
                    }

                }
                csvfile.Close();

            }


        }


        /// <summary>
        /// export un dossier technique
        /// </summary>
        /// <param name="contextlocal"></param>
        /// <param name="EngapiOnly">si false alors, la ligne de gamme n'est pas exportée</param>
        public void Export_Part_To_Produce_To_File(IContext contextlocal, bool EngapiOnly)
        {

            Clipper_Param.GetlistParam(contextlocal);

            ///    EngapiOnly = true;



            //recupere les path

            Clipper_Param.GetlistParam(contextlocal);
            string CsvExportPath = Clipper_Param.GetPath("EXPORT_DT") + "\\DonnesTech.txt";
            //chargement de la liste de piece a retourner
            IEntitySelector select_reference_list = new EntitySelector();
            IEntitySelector select_preparation_list = new EntitySelector();

            //IEntity clipper_machine = null;
            //IEntity centre_frais_clipper = null;

            Dictionary<string, string> centre_frais_dictionnary = null;
            centre_frais_dictionnary = new Dictionary<string, string>();
            //machine clipper
            Get_Clipper_Machine(contextlocal, out IEntity clipper_machine, out IEntity centre_frais_clipper, out centre_frais_dictionnary);
            //on retourne les pieces marquées "sansdt"

            IEntityList Part_To_Produce_sans_dt_filter = contextlocal.EntityManager.GetEntityList("_TO_PRODUCE_REFERENCE", "SANS_DT", ConditionOperator.Equal, true);
            Part_To_Produce_sans_dt_filter.Fill(false);

            /*
            IDynamicExtendedEntityList references_sansdt = contextlocal.EntityManager.GetDynamicExtendedEntityList("_REFERENCE", sans_dt_filter);
            references_sansdt.Fill(false);
            */
            IDynamicExtendedEntityList Part_To_Produce_sansdt = contextlocal.EntityManager.GetDynamicExtendedEntityList("_TO_PRODUCE_REFERENCE", Part_To_Produce_sans_dt_filter);
            Part_To_Produce_sansdt.Fill(false);


            select_reference_list.Init(contextlocal, Part_To_Produce_sansdt);
            select_reference_list.MultiSelect = true;

            //ecriture remplacement du fichier
            File_Tools.replaceFile(CsvExportPath, false);

            MessageBox.Show("Seules les pieces selectionnées pour l'export seront visibles", "Export Pieces à porduire ", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            using (StreamWriter csvfile = new StreamWriter(CsvExportPath, true))
            {


                if (select_reference_list.ShowDialog() == System.Windows.Forms.DialogResult.OK)

                {
                    //on control si la matiere est la meme que la matiere precedement demandée

                    foreach (IEntity part_To_Produce in select_reference_list.SelectedEntity)
                    {
                        IEntity stock = null;
                        IEntity sheet = null;
                        //IEntity current_machine = null;
                        IEntity reference = null;
                        //IEntity to_produce_ref = null;
                        IMachineManager machinemanager = new MachineManager();
                        

                        if (part_To_Produce.GetFieldValueAsEntity("CENTREFRAIS") == null)
                        {
                            string msg = "CE PROGRAMME VA ETRE ARRETTE : \r\n Vous devez definir un centre de frais pour l'affaire " + part_To_Produce.GetFieldValueAsString("AFFAIRE");
                            MessageBox.Show(msg, "Export Dossier technique", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            Environment.Exit(0);
                        }

                        //reference non vide et integrité
                        if (part_To_Produce != null && CheckDataIntegrity(contextlocal, part_To_Produce) == true)
                        {
                            //recupoeration de la piece 2d correspondante
                            reference = part_To_Produce.GetFieldValueAsEntity("_REFERENCE");
                            //creation d'un part info
                            AF_ImportTools.PartInfo part_infos = new PartInfo();
                            //GetPartinfos(ref IContext contextlocal, IEntity Part, IEntity Centrefrais)
                            // part_infos.GetPartinfos(ref contextlocal, reference);
                            //a uniformiser


                            part_infos.GetPartinfos(ref contextlocal, reference, part_To_Produce.GetFieldValueAsEntity("CENTREFRAIS"));


                            //pour facilite l'ecriture du fichier de sortie on stock toutes les infos dans des listes d'objets
                            List<object> engapi = new List<object>();
                            List<object> gapiece = new List<object>();
                            List<object> nomenpiece = new List<object>();
                            string separator = ";";

                            // selected_reference = to_produce_ref.GetFieldValueAsEntity("_REFERENCE");


                            //recuperation d'une tole/stock dont la matiere est egale a celle de la  piece
                            if (GetRadomSheet(contextlocal, reference, out stock, out sheet) != true)
                            {
                                Alma_Log.Write_Log_Important("aucun element de stock n' a ete trouve, seules les informations renseignées ou calculables vont etre renvoyees ");
                            };


                            //recuperation des infos de la piece
                            double surface = 0;
                            double parttime = part_infos.AlmaCam_PartTime;//.Quote_part_cyle_time + 0.1;
                            //unité en m2
                            //surface = selected_reference.GetFieldValueAsDouble("_SURFACE") * 10E-6;
                            surface = part_infos.Surface * 10E-6;
                            string codepiece = SimplifiedMethods.ConvertNullStringToEmptystring(reference.GetFieldValueAsString("_NAME"));
                            //unite en mm

                            double thickness = part_infos.Thickness;
                            //string matiere = selected_reference.GetFieldValueAsEntity("_MATERIAL").GetFieldValueAsString("_NAME");
                            string matiere = SimplifiedMethods.ConvertNullStringToEmptystring(part_infos.Material);
                            double feed = AF_ImportTools.Machine_Info.GetFeed(contextlocal, part_infos.DefaultMachineEntity, part_infos.MaterialEntity);


                            double xdim;
                            double ydim;
                            //longeur > largeur
                            if (part_infos.Width > part_infos.Height)
                            {
                                 xdim = part_infos.Width;
                                 ydim = part_infos.Height;

                            }
                            else
                            {
                                 ydim = part_infos.Width;
                                 xdim = part_infos.Height;
                            }
                            //double poids = selected_reference.GetFieldValueAsDouble("_WEIGHT") * 10E-3;
                            double poids = part_infos.Weight;

                            //données clipper de la piece
                            string affaire = SimplifiedMethods.ConvertNullStringToEmptystring(part_To_Produce.GetFieldValueAsString("AFFAIRE"));
                            string description = SimplifiedMethods.ConvertNullStringToEmptystring(part_To_Produce.GetFieldValueAsString("_DESCRIPTION"));
                            //string emffile = part_infos.EmfFile;
                            string emffile;
                            /// a revoir avec les grometrie par defaut
                            //emffile = SimplifiedMethods.ConvertNullStringToEmptystring(@reference.GetImageFieldValueAsLinkFile("_PREVIEW"));
                            emffile = SimplifiedMethods.GetPreview(reference);
                            string en_rang = SimplifiedMethods.ConvertNullStringToEmptystring(part_To_Produce.GetFieldValueAsString("EN_RANG"));
                            string en_pere_piece = SimplifiedMethods.ConvertNullStringToEmptystring(part_To_Produce.GetFieldValueAsString("EN_PERE_PIECE"));
                            string idlnbom = SimplifiedMethods.ConvertNullStringToEmptystring(part_To_Produce.GetFieldValueAsString("IDLNBOM"));
                            string referenceId = reference.Id32.ToString();

                            string centrefrais = "";
                            centrefrais = SimplifiedMethods.ConvertNullStringToEmptystring(reference.GetFieldValueAsEntity("_DEFAULT_CUT_MACHINE_TYPE").GetFieldValueAsEntity("CENTREFRAIS_MACHINE").GetFieldValueAsString("_CODE"));

                            //string centrefrais = SimplifiedMethods.ConvertEmptyStringToNullstring(to_produce_ref.GetFieldValueAsEntity("CENTREFRAIS").ToString());
                            string codematiere = SimplifiedMethods.ConvertNullStringToEmptystring(stock.GetFieldValueAsString("_NAME"));
                            //recuperation des infos de la tole
                            double xDim_Tole = sheet.GetFieldValueAsDouble("_LENGTH");
                            double yDim_Tole = sheet.GetFieldValueAsDouble("_WIDTH");
                            //double surface_Tole = sheet.GetFieldValueAsDouble("_SURFACE");
                            //pour le moment on considere que la surface de la tole est
                            double surface_Tole = xDim_Tole * yDim_Tole * 10E-6;
                            string nomdelaTole = SimplifiedMethods.ConvertNullStringToEmptystring(stock.GetFieldValueAsString("_NAME"));
                            //section engapi

                            engapi.Add("ENGAPI");
                            engapi.Add(codepiece);
                            engapi.Add(affaire);
                            engapi.Add(description);
                            engapi.Add(emffile);
                            engapi.Add(string.Format("{0:0.00000}", poids / 1000));
                            //engapi.Add(idlnbom);
                            engapi.Add(en_rang);
                            engapi.Add(en_pere_piece);
                            engapi.Add(referenceId);


                            gapiece.Add("GAPIECE");
                            gapiece.Add(centrefrais);
                            //pour le moment ces valeurs sont nulles
                            gapiece.Add(string.Format("{0:0.00000}", 0));
                            gapiece.Add(string.Format("{0:0.00000}", parttime));
                           
                            ///

                            nomenpiece.Add("NOMENPIECE");
                            nomenpiece.Add(codematiere);
                            nomenpiece.Add(xDim_Tole);
                            nomenpiece.Add(yDim_Tole);


                            if (surface_Tole > 0)
                            {
                                nomenpiece.Add(string.Format("{0:0.00000}", surface / surface_Tole));
                            }

                            nomenpiece.Add(string.Format("{0:0.00000}", xdim));
                            nomenpiece.Add(string.Format("{0:0.00000}", ydim));



                            string myline = "";
                            var last_o = engapi.LastOrDefault();
                            separator = ";";
                            foreach (object o in engapi)
                            {
                                if (o.Equals(last_o))
                                { separator = ""; }
                                myline += o.ToString() + separator;
                            }
                            last_o = null;
                            csvfile.WriteLine(myline.Replace(",", "."));

                            //ecriture de gapiece uniquement si retour possible
                            //actuellement si la gamme existe clipper rajoute une nouvelle gamme
                            //on force EngapiOnly a true
                            EngapiOnly = false;
                            if (EngapiOnly == false)
                            {
                                myline = "";
                                last_o = gapiece.LastOrDefault();
                                separator = ";";

                                foreach (object o in gapiece)
                                {
                                    if (o.Equals(last_o)) { separator = ""; }
                                    myline += o.ToString() + separator;
                                }
                                csvfile.WriteLine(myline.Replace(",", "."));
                            }
                            //ecriture de nomepiece
                            myline = "";
                            last_o = nomenpiece.LastOrDefault();
                            separator = ";";

                            foreach (object o in nomenpiece)
                            {
                                if (o.Equals(last_o)) { separator = ""; }
                                myline += o.ToString() + separator;
                            }

                            csvfile.WriteLine(myline.Replace(",", "."));
                            last_o = null;
                            //on set le setfiled value a false

                            part_To_Produce.SetFieldValue("SANS_DT", false);
                            part_To_Produce.Save();
                            part_infos = null;

                        }
                    }

                }
                csvfile.Close();

            }


        }
    }



    /// <summary>
    /// 
    /// </summary>




    #endregion

    #region Integration

        /*
    public class Clipper_8_integrate : ScriptModelCustomization, IScriptModelCustomization
    {

        public override bool Execute(IContext context, IContext hostContext)
        {
            IEntityType entityType = context.Kernel.GetEntityType("_TO_PRODUCE_REFERENCE");
            IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType, null, "_ACTCUT", null);
            entityTypeFactory.Key = "_TO_PRODUCE_REFERENCE";
            entityTypeFactory.Name = "Pièces à produire";
            entityTypeFactory.DefaultDisplayKey = "_NAME";
            entityTypeFactory.ActAsEnvironment = false;

            {
                IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                fieldDescription.Key = "AFFAIRE";
                fieldDescription.Name = "*Affaire";
                fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                fieldDescription.Mandatory = false;
                fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
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
            return true;
        }




        public void Integrate_All(string databasemane)
        {


            List<Tuple<string, string>> CustoFileCsList = new List<Tuple<string, string>>();
            CustoFileCsList.Add(new Tuple<string, string>("Command", Properties.Resources.All_Commandes));
            CustoFileCsList.Add(new Tuple<string, string>("Entities", Properties.Resources.All_Entities));
            CustoFileCsList.Add(new Tuple<string, string>("FormulasAndEvents", Properties.Resources.All_Events));


            foreach (Tuple<string, string> CustoFileCs in CustoFileCsList)
            {
                string CommandCsPath = Path.GetTempPath() + CustoFileCs.Item1 + ".cs";
                File.WriteAllText(CommandCsPath, CustoFileCs.Item2);

                IModelsRepository modelsRepository = new ModelsRepository();
                ModelManager modelManager = new ModelManager(modelsRepository);
                modelManager.CustomizeModel(CommandCsPath, databasemane, true);
                File.Delete(CommandCsPath);
            }
            //LogMessage.Write("End Install Customisation");
            return;





        }
        public void Field_Add_PartToProduce(IContext localcontexte)
        {
            //recuperation des champs spé. de la base selectionnée
            IEntityType entityType = localcontexte.Kernel.GetEntityType("_TO_PRODUCE_REFERENCE");
            IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(localcontexte, 1, entityType, null, "_ACTCUT", null);
            entityTypeFactory.Key = "_TO_PRODUCE_REFERENCE";
            entityTypeFactory.Name = "Pièces à produire";
            entityTypeFactory.DefaultDisplayKey = "_NAME";
            entityTypeFactory.ActAsEnvironment = false;
            //foreach (IField f in entityTypeFactory.EditEntityType.GetStandardFieldList())
            //IWpmEnumerable ie = entityTypeFactory.EditEntityType.GetStandardFieldList();
            //recuperation  de la liste des champs spé.

        }

    }

    */
    #endregion



    #region exception
   


    /// <summary>
    /// il manque le numero d'affaire
    /// </summary>
    public class MissingAffaireException : Exception
    {

        public  MissingAffaireException()
        {
            if (AF_Clipper_Dll.Clipper_Param.GetVerbose_Log() == true)
            {

                MessageBox.Show("Numero d'affaire incorrecte.");
                Alma_Log.Write_Log_Important("Numero d'affaire incorrecte.");
            }
            Alma_Log.Write_Log("Numero d'affaire incorrecte.");
           
        }



    }
    #endregion
}







///recuperation des pipes pour communication interapplication
///https://msdn.microsoft.com/en-us/library/bb546102.aspx
///eviter les socket pour des raisons de securité
#endregion