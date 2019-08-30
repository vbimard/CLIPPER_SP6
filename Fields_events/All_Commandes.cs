using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Wpm.Schema.Kernel;
using Wpm.Implement.Manager;
using Wpm.Implement.ModelSetting;
using Alma.BaseUI.DescriptionEditor;

namespace Wpm.Implement.ModelSetting
{
    public partial class ImportUserCommandType : ScriptModelCustomization, IScriptModelCustomization
    {
        public override bool Execute(IContext context, IContext hostContext)
        {
            #region CLIP Configuration
            
            {
                ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, true);
                commandTypeFactory.Key = "CLIP_CONFIGURATION";
                commandTypeFactory.Name = "CLIP Configuration";
                commandTypeFactory.PlugInFileName = @"AF_Clipper_Dll.dll";
                commandTypeFactory.PlugInNameSpace = "AF_Clipper_Dll";
                commandTypeFactory.PlugInClassName = "ClipperIE_Global";
                commandTypeFactory.Shortcut = Shortcut.None;
                commandTypeFactory.LargeImage = true;
                commandTypeFactory.ImageFile = "";
                
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "IMPORT_CDA";
                    parameterDescription.Name = "Chemin complet vers le cahier d'affaire ( CAHIER_AFFAIRE.csv )";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                    parameterDescription.DefaultValue = @"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\CAHIER_AFFAIRE.csv";
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
                    parameterDescription.Name = "Dossier d'export des retours GP";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                    parameterDescription.DefaultValue = @"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Export_GPAO";
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "EXPORT_Dt";
                    parameterDescription.Name = "Dossier d'export des clotures";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                    parameterDescription.DefaultValue = @"C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Export_DT";
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "IMPORT_AUTO";
                    parameterDescription.Name = "Activer l' import Auto";
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
                    parameterDescription.DefaultValue = @"0#_NAME#string;1#AFFAIRE#string;2#THICKNESS#string;3#MATERIAL_CLIPPER#string;4#CENTREFRAIS#string;5#TECHNOLOGIE#string;6#FAMILY#string;7#IDLNROUT#string;8#CENTREFRAISSUIV#string;9#CUSTOMER#string;10#_QUANTITY#integer;11#QUANTITY#double;12#ECOQTY#string;13#STARTDATE#date;14#ENDDATE#date;15#PLAN#string;16#FORMATCLIP#string;17#IDMAT#string;18#IDLNBOM#string;19#NUMMAG#string;20#FILENAME#string;21#_DESCRIPTION#string;22#AF_CDE#string;23#DELAI_INT#date;24#EN_RANG#string;25#EN_PERE_PIECE#string";
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "MODEL_DM";
                    parameterDescription.Name = "Model Fichier de Stock";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.String;
                    parameterDescription.DefaultValue = @"0#_NAME#string;1#_MATERIAL#string;2#_LENGTH#double;3#_WIDTH#double;4#THICKNESS#double;5#QTY_TOT#integer;6#_QUANTITY#integer;7#GISEMENT#string;8#NUMMAG#string;9#NUMMATLOT#string;10#NUMCERTIF#string;11#NUMLOT#string;12#NUMCOUL#string;13#IDCLIP#string;14#FILENAME#string";
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "MODEL_PATH";
                    parameterDescription.Name = "Model Chemin d'export";
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
                    parameterDescription.Key = "VERBOSE_LOG";
                    parameterDescription.Name = "Log verbeux";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                    parameterDescription.DefaultValue = false;
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "CLIPPER_MACHINE_CF";
                    parameterDescription.Name = "centre de frais machien clipper";
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
                    parameterDescription.Name = "Nom de l'editeur AlmaCam";
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
                    parameterDescription.Name = "Explosion de fichiers de retours sur mutliplicité";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                    parameterDescription.DefaultValue = false;
                    commandTypeFactory.ParameterList.Add(parameterDescription);
                    
                }
                {
                    IParameterDescription parameterDescription = new ParameterDescription(context.Kernel.UnitSystem, true);
                    parameterDescription.Key = "ACTIVATE_OMISSION";
                    parameterDescription.Name = "Activation des omissions";
                    parameterDescription.ParameterDescriptionType = ParameterDescriptionType.Boolean;
                    parameterDescription.DefaultValue = false;
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
            
            #endregion
            #region Export Dossier Technique
            
            {
                ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                commandTypeFactory.Key = "CLIPPER_EXPORT_DT_DLL";
                commandTypeFactory.Name = "Export Dossier Technique";
                commandTypeFactory.PlugInFileName = @"AF_Clipper_Dll.dll";
                commandTypeFactory.PlugInNameSpace = "AF_Clipper_Dll";
                commandTypeFactory.PlugInClassName = "Clipper_Export_DT";
                commandTypeFactory.Shortcut = Shortcut.None;
                commandTypeFactory.WorkOnEntityTypeKey = "_REFERENCE";
                commandTypeFactory.LargeImage = true;
                
                if (!commandTypeFactory.UpdateModel())
                {
                    foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                    {
                        hostContext.TraceLogger.TraceError(error.Message, true);
                    }
                    return false;
                }
                
            }
            
            #endregion
            #region Import Fournitures
            
            {
                ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                commandTypeFactory.Key = "IMPORT_CLIPPER_FOURNITURES";
                commandTypeFactory.Name = "Import Fournitures";
                commandTypeFactory.PlugInFileName = @"AF_Import_ODBC_Clipper_AlmaCam.dll";
                commandTypeFactory.PlugInNameSpace = "AF_Import_ODBC_Clipper_AlmaCam";
                commandTypeFactory.PlugInClassName = "Clipper_Import_Fournitures_Divers_Processor";
                commandTypeFactory.Shortcut = Shortcut.None;
                commandTypeFactory.WorkOnEntityTypeKey = "_SIMPLE_SUPPLY";
                commandTypeFactory.LargeImage = true;
                
                if (!commandTypeFactory.UpdateModel())
                {
                    foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                    {
                        hostContext.TraceLogger.TraceError(error.Message, true);
                    }
                    return false;
                }
                
            }
            
            #endregion
            #region Import OF
            
            {
                ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                commandTypeFactory.Key = "IMPORT_OF";
                commandTypeFactory.Name = "Import OF";
                commandTypeFactory.PlugInFileName = @"AF_Clipper_Dll.dll";
                commandTypeFactory.PlugInNameSpace = "AF_Clipper_Dll";
                commandTypeFactory.PlugInClassName = "Clipper_8_Import_OF_Processor";
                commandTypeFactory.Shortcut = Shortcut.None;
                commandTypeFactory.WorkOnEntityTypeKey = "_TO_PRODUCE_REFERENCE";
                commandTypeFactory.LargeImage = true;
                commandTypeFactory.ImageFile = "";
                
                if (!commandTypeFactory.UpdateModel())
                {
                    foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                    {
                        hostContext.TraceLogger.TraceError(error.Message, true);
                    }
                    return false;
                }
                
            }
            
            #endregion
            #region Import Stock
            
            {
                ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                commandTypeFactory.Key = "IMPORT_STOCK";
                commandTypeFactory.Name = "Import Stock";
                commandTypeFactory.PlugInFileName = @"AF_Clipper_Dll.dll";
                commandTypeFactory.PlugInNameSpace = "AF_Clipper_Dll";
                commandTypeFactory.PlugInClassName = "Clipper_8_Import_Stock_Processor";
                commandTypeFactory.Shortcut = Shortcut.None;
                commandTypeFactory.WorkOnEntityTypeKey = "_STOCK";
                commandTypeFactory.LargeImage = true;
                commandTypeFactory.ImageFile = "";
                
                if (!commandTypeFactory.UpdateModel())
                {
                    foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                    {
                        hostContext.TraceLogger.TraceError(error.Message, true);
                    }
                    return false;
                }
                
            }
            
            #endregion
            #region ImporterMatiere
            
            {
                ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                commandTypeFactory.Key = "IMPORT_CLIPPER_MATERIAL";
                commandTypeFactory.Name = "ImporterMatiere";
                commandTypeFactory.PlugInFileName = @"AF_Import_ODBC_Clipper_AlmaCam.dll";
                commandTypeFactory.PlugInNameSpace = "AF_Import_ODBC_Clipper_AlmaCam";
                commandTypeFactory.PlugInClassName = "Clipper8_ImportMatiere_Processor";
                commandTypeFactory.Shortcut = Shortcut.None;
                commandTypeFactory.WorkOnEntityTypeKey = "_QUALITY";
                commandTypeFactory.LargeImage = true;
                commandTypeFactory.ImageFile = "";
                
                if (!commandTypeFactory.UpdateModel())
                {
                    foreach (ModelSettingError error in commandTypeFactory.ErrorList)
                    {
                        hostContext.TraceLogger.TraceError(error.Message, true);
                    }
                    return false;
                }
                
            }
            
            #endregion
            #region Sheet requirement
            
            {
                ICommandTypeFactory commandTypeFactory = new CommandTypeFactory(context, 1, null, false);
                commandTypeFactory.Key = "SHEET_REQUIREMENT";
                commandTypeFactory.Name = "Sheet requirement";
                commandTypeFactory.PlugInFileName = @"AF_Clipper_Dll.dll";
                commandTypeFactory.PlugInNameSpace = "AF_Clipper_Dll";
                commandTypeFactory.PlugInClassName = "Clipper_Requirement_Processor";
                commandTypeFactory.Shortcut = Shortcut.None;
                commandTypeFactory.WorkOnEntityTypeKey = "_TO_CUT_NESTING";
                commandTypeFactory.LargeImage = true;
                
                if (!commandTypeFactory.UpdateModel())
                {
                    foreach (ModelSettingError error in commandTypeFactory.ErrorList)
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
}
