using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Wpm.Schema.Kernel;
using Wpm.Implement.Manager;
using Wpm.Implement.ModelSetting;
using Alma.BaseUI.DescriptionEditor;


using Wpm.Implement.Manager;
using Wpm.Implement.Processor;
using Wpm.Implement.ComponentEditor;  // ouverture de fenetres de selection almacam




namespace Wpm.Implement.ModelSetting
{
    public partial class ImportUserCommandType : ScriptModelCustomization, IScriptModelCustomization
    {


        public override bool Execute(IContext context, IContext hostContext)
        {

            #region Export Dossier Technique
            {
                if (CommmandeExists(context, "CLIPPER_EXPORT_DT_DLL") == false)
                {
                    {
                        ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                        commandTypeFactory.Key = "CLIPPER_EXPORT_DT_DLL";
                        commandTypeFactory.Name = "Clipper  : Export Dossier Technique";
                        commandTypeFactory.PlugInFileName = @"AF_Clipper_Dll.dll";
                        commandTypeFactory.PlugInNameSpace = "AF_Clipper_Dll";
                        commandTypeFactory.PlugInClassName = "Clipper_8_Export_DT_Processor";
                        commandTypeFactory.Shortcut = Shortcut.None;
                        commandTypeFactory.WorkOnEntityTypeKey = "_REFERENCE";
                        commandTypeFactory.LargeImage = true;
                        if (File.Exists(@"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png"))
                        {
                            commandTypeFactory.ImageFile = @"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png";
                        }
                        else
                        {
                            commandTypeFactory.ImageFile = "";
                        }

                        if (!commandTypeFactory.UpdateModel())
                        {
                            foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                            {
                                hostContext.TraceLogger.TraceError(error.Message, true);
                            }
                            return false;
                        }

                    }
                }
            }
            #endregion

            #region Import OF
            {
                if (CommmandeExists(context, "IMPORT_OF") == false)
                {
                    {
                        ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                        commandTypeFactory.Key = "IMPORT_OF";
                        commandTypeFactory.Name = "Clipper  : Import OF";
                        commandTypeFactory.PlugInFileName = @"AF_Clipper_Dll.dll";
                        commandTypeFactory.PlugInNameSpace = "AF_Clipper_Dll";
                        commandTypeFactory.PlugInClassName = "Clipper_8_Import_OF_Processor";
                        commandTypeFactory.Shortcut = Shortcut.None;
                        commandTypeFactory.WorkOnEntityTypeKey = "_TO_PRODUCE_REFERENCE";
                        commandTypeFactory.LargeImage = true;
                        if (File.Exists(@"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png"))
                        {
                            commandTypeFactory.ImageFile = @"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png";
                        }
                        else
                        {
                            commandTypeFactory.ImageFile = "";
                        }

                        if (!commandTypeFactory.UpdateModel())
                        {
                            foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                            {
                                hostContext.TraceLogger.TraceError(error.Message, true);
                            }
                            return false;
                        }

                    }
                }
            }
            #endregion

            #region Import Stock
            {
                if (CommmandeExists(context, "IMPORT_STOCK") == false)
                {
                    {
                        ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                        commandTypeFactory.Key = "IMPORT_STOCK";
                        commandTypeFactory.Name = "Clipper  : Import Stock";
                        commandTypeFactory.PlugInFileName = @"AF_Clipper_Dll.dll";
                        commandTypeFactory.PlugInNameSpace = "AF_Clipper_Dll";
                        commandTypeFactory.PlugInClassName = "Clipper_8_Import_Stock_Processor";
                        commandTypeFactory.Shortcut = Shortcut.None;
                        commandTypeFactory.WorkOnEntityTypeKey = "_STOCK";
                        commandTypeFactory.LargeImage = true;
                        if (File.Exists(@"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png"))
                        {
                            commandTypeFactory.ImageFile = @"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png";
                        }
                        else
                        {
                            commandTypeFactory.ImageFile = "";
                        }

                        if (!commandTypeFactory.UpdateModel())
                        {
                            foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                            {
                                hostContext.TraceLogger.TraceError(error.Message, true);
                            }
                            return false;
                        }

                    }
                }
            }
            #endregion

            #region ImporterMatiere
            {
                if (CommmandeExists(context, "IMPORT_CLIPPER_MATERIAL") == false)
                {
                    {
                        ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                        commandTypeFactory.Key = "IMPORT_CLIPPER_MATERIAL";
                        commandTypeFactory.Name = "Clipper  : Import Matiere";
                        commandTypeFactory.PlugInFileName = @"AF_Import_ODBC_Clipper_AlmaCam.dll";
                        commandTypeFactory.PlugInNameSpace = "AF_Import_ODBC_Clipper_AlmaCam";
                        commandTypeFactory.PlugInClassName = "Clipper8_ImportMatiere_Processor";
                        commandTypeFactory.Shortcut = Shortcut.None;
                        commandTypeFactory.WorkOnEntityTypeKey = "_QUALITY";
                        commandTypeFactory.LargeImage = true;

                        if (File.Exists(@"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png"))
                        {
                            commandTypeFactory.ImageFile = @"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png";
                        }
                        else
                        {
                            commandTypeFactory.ImageFile = "";
                        }

                        if (!commandTypeFactory.UpdateModel())
                        {
                            foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                            {
                                hostContext.TraceLogger.TraceError(error.Message, true);
                            }
                            return false;
                        }

                    }
                }
            }
            #endregion

            #region Import Fournitures

            {
                if (CommmandeExists(context, "IMPORT_FOURNITURES") == false)
                {
                    {
                        ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                        commandTypeFactory.Key = "IMPORT_FOURNITURES";
                        commandTypeFactory.Name = "Clipper  : Importer Fournitures";
                        commandTypeFactory.PlugInFileName = @"AF_Import_ODBC_Clipper_AlmaCam.dll";
                        commandTypeFactory.PlugInNameSpace = "AF_Import_ODBC_Clipper_AlmaCam";
                        commandTypeFactory.PlugInClassName = "Clipper_Import_Fournitures_Divers_Processor";
                        commandTypeFactory.Shortcut = Shortcut.None;
                        commandTypeFactory.WorkOnEntityTypeKey = "_SIMPLE_SUPPLY";
                        commandTypeFactory.LargeImage = true;

                        if (File.Exists(@"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png"))
                        {
                            commandTypeFactory.ImageFile = @"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png";
                        }
                        else
                        {
                            commandTypeFactory.ImageFile = "";
                        }

                        if (!commandTypeFactory.UpdateModel())
                        {
                            foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                            {
                                hostContext.TraceLogger.TraceError(error.Message, true);
                            }
                            return false;
                        }

                    }
                }
            }

            #endregion

            #region CLIP Configuration



            {
                if (CommmandeExists(context, "CLIP_CONFIGURATION") == false)
                {
                    ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, true);
                    commandTypeFactory.Key = "CLIP_CONFIGURATION";
                    commandTypeFactory.Name = "Clipper  : CLIP Configuration";
                    commandTypeFactory.PlugInFileName = @"AF_Clipper_Dll.dll";
                    commandTypeFactory.PlugInNameSpace = "AF_Clipper_Dll";
                    commandTypeFactory.PlugInClassName = "ClipperIE_Global";
                    commandTypeFactory.Shortcut = Shortcut.None;
                    commandTypeFactory.LargeImage = true;

                    if (File.Exists(@"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png"))
                    {
                        commandTypeFactory.ImageFile = @"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png";
                    }
                    else
                    {
                        commandTypeFactory.ImageFile = "";
                    }

                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "IMPORT_CDA";
                        parameterDescription.Name = "Chemin complet vers le cahier d affaire ";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Import_OF\CAHIER_AFFAIRE.csv";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "IMPORT_DM";
                        parameterDescription.Name = "Chemin complet vers le stock";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Import_Stock\DISPO_MAT.csv";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "EXPORT_Rp";
                        parameterDescription.Name = "Dossier d export des retours GP";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Export_GPAO";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "EXPORT_Dt";
                        parameterDescription.Name = "Dossier d export des clotures";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Export_DT";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "IMPORT_AUTO";
                        parameterDescription.Name = "Activer l import Auto";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                        parameterDescription.DefaultValue = true;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "EMF_DIRECTORY";
                        parameterDescription.Name = "Emf Directory";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\EMF";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "MODEL_CA";
                        parameterDescription.Name = "Model Cahier Affaire";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"0#_NAME#string;1#AFFAIRE#string;2#THICKNESS#string;3#_MATERIAL#string;4#CENTREFRAIS#string;5#TECHNOLOGIE#string;6#FAMILY#string;7#IDLNROUT#string;8#CENTREFRAISSUIV#string;9#_FIRM#string;10#_QUANTITY#integer;11#QUANTITY#double;12#ECOQTY#string;13#STARTDATE#date;14#ENDDATE#date;15#PLAN#string;16#FORMATCLIP#string;17#IDMAT#string;18#IDLNBOM#string;19#NUMMAG#string;20#FILENAME#string;21#_DESCRIPTION#string;22#_CLIENT_ORDER_NUMBER#string;23#DELAI_INT#date;24#EN_RANG#string;25#EN_PERE_PIECE#string;26#ID_PIECE_CFAO#string";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "MODEL_DM";
                        parameterDescription.Name = "Model Fichier de Stock";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        //parameterDescription.DefaultValue = @"0#_NAME#string;1#_MATERIAL#string;2#_LENGTH#double;3#_WIDTH#double;4#THICKNESS#double;5#QTY_TOT#integer;6#_QUANTITY#integer;7#GISEMENT#string;8#NUMMAG#string;9#NUMMATLOT#string;10#NUMCERTIF#string;11#NUMLOT#string;12#NUMCOUL#string;13#IDCLIP#string;14#FILENAME#string";
                        //modification sp5
                        //parameterDescription.DefaultValue = @"0#_NAME#string;1#_MATERIAL#string;2#_LENGTH#double;3#_WIDTH#double;4#THICKNESS#double;5#QTY_TOT#integer;6#_REST_QUANTITY#integer;7#GISEMENT#string;8#NUMMAG#string;9#NUMMATLOT#string;10#NUMCERTIF#string;11#NUMLOT#string;12#NUMCOUL#string;13#IDCLIP#string;14#FILENAME#string;15#NAF_LOTIE#string";
                        //naf jamais envoyé si vide--> on vire le champs
                        parameterDescription.DefaultValue = @"0#_NAME#string;1#_MATERIAL#string;2#_LENGTH#double;3#_WIDTH#double;4#THICKNESS#double;5#QTY_TOT#integer;6#_REST_QUANTITY#integer;7#GISEMENT#string;8#NUMMAG#string;9#NUMMATLOT#string;10#NUMCERTIF#string;11#NUMLOT#string;12#NUMCOUL#string;13#IDCLIP#string;14#FILENAME#string";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "MODEL_PATH";
                        parameterDescription.Name = "Model Chemin d export";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"0#TECHNOLOGIE#string;1#ImportCda#string;0#ImportDM#string;2#ExportRp#string;3#ExportDT#string;4#Centredefrais#string;5#Destination_Path#string;6#Source_Path#string";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "APPLICATION1";
                        parameterDescription.Name = "Path to AlmaCamUser1";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"C:\AlmaCAM\Bin\AlmaCamUser1.exe";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }

                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "CLIPPER_MACHINE_CF";
                        parameterDescription.Name = "centre de frais machine clipper";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"CLIP";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "STRING_FORMAT_DOUBLE";
                        parameterDescription.Name = "format double";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"{0:0.00###}";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "ALMACAM_EDITOR_NAME";
                        parameterDescription.Name = "Nom de l editeur AlmaCam";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"Wpm.Implement.Editor.exe";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "SHEET_REQUIREMENT";
                        parameterDescription.Name = "Dossier d'export des besoins matieres";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                        parameterDescription.DefaultValue = @"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Export_Sheet_requirements";
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "EXPLODE_MULTIPLICITY";
                        parameterDescription.Name = "Explosion de fichiers de retours sur mutliplicite";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                        parameterDescription.DefaultValue = true;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "ACTIVATE_OMISSION";
                        parameterDescription.Name = "Activation des omissions";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                        parameterDescription.DefaultValue = true;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }

                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "AF_ACTIVATE_SHEET_ON_SENDTOWSHOP";
                        parameterDescription.Name = "Validation des chutes à l envoie a la coupe";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                        parameterDescription.DefaultValue = false;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }

                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "VERBOSE_LOG";
                        parameterDescription.Name = "Log verbeux";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                        parameterDescription.DefaultValue = true;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }

                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "AF_MULTIDIM_MODE";
                        parameterDescription.Name = "Favoriser le Multidim";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                        parameterDescription.DefaultValue = true;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }


                    if (!commandTypeFactory.UpdateModel())
                    {
                        foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                        {
                            hostContext.TraceLogger.TraceError(error.Message, true);
                        }
                        return false;
                    }
                }
            }

            #endregion
            
            #region Clipper  : Importer Tubes

            {
                if (CommmandeExists(context, "AF_IMPORT_TUBES") == false)
                {
                    ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                    commandTypeFactory.Key = "AF_IMPORT_TUBES";
                    commandTypeFactory.Name = "Clipper  : Importer Tubes";
                    commandTypeFactory.PlugInFileName = @"AF_Import_ODBC_Clipper_AlmaCam.dll";
                    commandTypeFactory.PlugInNameSpace = "AF_Import_ODBC_Clipper_AlmaCam";
                    commandTypeFactory.PlugInClassName = "Clipper8_ImportTubes_Processor";
                    commandTypeFactory.Shortcut = Shortcut.None;
                    commandTypeFactory.WorkOnEntityTypeKey = "_BARTUBE";
                    commandTypeFactory.LargeImage = true;

                    if (File.Exists(@"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png"))
                    {
                        commandTypeFactory.ImageFile = @"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png";
                    }
                    else
                    {
                        commandTypeFactory.ImageFile = "";
                    }


                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "AF_ROUND";
                        parameterDescription.Name = "Importer  Ronds";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                        parameterDescription.DefaultValue = true;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "AF_ROUND_TUBE";
                        parameterDescription.Name = "Importer Tubes";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                        parameterDescription.DefaultValue = true;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "AF_RECTANGLE";
                        parameterDescription.Name = "Importer  Rectangles";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                        parameterDescription.DefaultValue = true;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "AF_FLAT";
                        parameterDescription.Name = "Importer  Plats";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                        parameterDescription.DefaultValue = true;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "AF_FLAT_THREESOLD";
                        parameterDescription.Name = "Seuil de longueur pour reconnaitre les plats";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Double;
                        parameterDescription.DefaultValue = 300;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    {
                        IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                        parameterDescription.Key = "AF_CHECK_NEW_QUALITY";
                        parameterDescription.Name = "Importer les matieres avant d'importer les tubes";
                        parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                        parameterDescription.DefaultValue = true;
                        commandTypeFactory.ParameterList.Add(parameterDescription);

                    }
                    if (!commandTypeFactory.UpdateModel())
                    {
                        foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                        {
                            hostContext.TraceLogger.TraceError(error.Message, true);
                        }
                        return false;
                    }
                }
            }

            #endregion

            #region export gp
            {
                if (CommmandeExists(context, "AF_EXPORT_NESTING") == false)
                {
                    {
                        ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                        commandTypeFactory.Key = "AF_EXPORT_NESTING";
                        commandTypeFactory.Name = "Clipper  : Export des placements vers clipper";
                        commandTypeFactory.PlugInFileName = @"AF_Clipper_Dll.dll";
                        commandTypeFactory.PlugInNameSpace = "AF_Clipper_Dll";
                        commandTypeFactory.PlugInClassName = "Clipper_8_Export_GP_Processor";
                        commandTypeFactory.Shortcut = Shortcut.None;
                        commandTypeFactory.WorkOnEntityTypeKey = "_TO_CUT_NESTING";
                        commandTypeFactory.LargeImage = true;
                        if (File.Exists(@"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png"))
                        {
                            commandTypeFactory.ImageFile = @"C:\AlmaCAM\Bin\Icones\Clipper\" + commandTypeFactory.Key + ".png";
                        }
                        else
                        {
                            commandTypeFactory.ImageFile = "";
                        }

                        if (!commandTypeFactory.UpdateModel())
                        {
                            foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                            {
                                hostContext.TraceLogger.TraceError(error.Message, true);
                            }
                            return false;
                        }

                    }
                }
            }
            #endregion

            return true;
        }

        public Boolean CommmandeExists(IContext context, string commandekey)
        {

            Boolean rst = false;

            IList<CustomizedCommandItem> customizedItemList = GetCustomizedCommandTypeList(context);

            if (customizedItemList.Count > 0)
            {
                foreach (CustomizedCommandItem customizedCommandType in customizedItemList)
                {
                    if (customizedCommandType.Key == commandekey)
                    {

                        rst = true;
                    }
                }


            }



            return rst;



        }


        public static List<CustomizedCommandItem> GetCustomizedCommandTypeList(IContext DBContext)
        {
            List<CustomizedCommandItem> customizedCommandTypeList = new List<CustomizedCommandItem>();
            foreach (ISimpleCommandType simpleCommandType in DBContext.Kernel.SimpleCommandTypeList.Values)
            {
                if (simpleCommandType.OwnerLevel == 1 && simpleCommandType.StringTag == "FROM_WPM_WIZARD")
                    customizedCommandTypeList.Add(new CustomizedCommandItem(simpleCommandType.Name, simpleCommandType.Key, true, null));
            }
            foreach (ICommandType commandType in DBContext.Kernel.CommandTypeList.Values)
            {
                if (commandType.OwnerLevel == 1 && commandType.StringTag == "FROM_WPM_WIZARD")
                    customizedCommandTypeList.Add(new CustomizedCommandItem(commandType.Name, commandType.Key, false, commandType.WorkOn.Name));
            }
            return customizedCommandTypeList;
        }





    }
}
