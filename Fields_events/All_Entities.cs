using System;
using System.IO;
using System.Collections.Generic;
using Wpm.Schema.Kernel;
using Wpm.Implement.Manager;
using Wpm.Implement.ModelSetting;
using Alma.BaseUI.DescriptionEditor;

namespace Wpm.Implement.ModelSetting
{
    public partial class ImportUserEntityType : ScriptModelCustomization, IScriptModelCustomization
    {
        public override bool Execute(IContext context, IContext hostContext)
        {
            #region Tôles
            
            {
                IEntityType entityType = context.Kernel.GetEntityType("_SHEET");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType , null, "_SUPPLY", null);
                entityTypeFactory.Key = "_SHEET";
                entityTypeFactory.Name = "Tôles";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;
                
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "FILENAME";
                    fieldDescription.Name = "*File name";
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
                
            }
            
            #endregion
            #region Machines de coupe
            
            {
                IEntityType entityType = context.Kernel.GetEntityType("_CUT_MACHINE_TYPE");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType , null, "_RSC", null);
                entityTypeFactory.Key = "_CUT_MACHINE_TYPE";
                entityTypeFactory.Name = "Machines de coupe";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;
                
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "CENTREFRAIS_MACHINE";
                    fieldDescription.Name = "-centre de frais machine";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Lookup;
                    fieldDescription.LinkKey = "_CENTRE_FRAIS";
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
            #region Stock
            
            {
                IEntityType entityType = context.Kernel.GetEntityType("_STOCK");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType , null, "_SUPPLY", null);
                entityTypeFactory.Key = "_STOCK";
                entityTypeFactory.Name = "Stock";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;
                
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "QTE_TOT";
                    fieldDescription.Name = "*Qte_Tot";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Int;
                    fieldDescription.DefaultValue = 0;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "FILENAME";
                    fieldDescription.Name = "*FileName";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "NUMCERTIF";
                    fieldDescription.Name = "*Numero certif";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "IDCLIP";
                    fieldDescription.Name = "*Id_clip";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "NUMMATLOT";
                    fieldDescription.Name = "*Numero_Matiere_lotie";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "NUMMAG";
                    fieldDescription.Name = "*Numero_Magasin";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "GISEMENT";
                    fieldDescription.Name = "*Gisement";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "NUMLOT";
                    fieldDescription.Name = "*Numero lot";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "STOCK_IMPORT_NUMBER";
                    fieldDescription.Name = "*Import Number";
                    fieldDescription.Editable = FieldDescriptionEditableType.NoEditable;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "NAF_LOTIE";
                    fieldDescription.Name = "*Numero_Affaire";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_STOCK_RENAME";
                    fieldDescription.Name = "*Renommé ?";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Boolean;
                    fieldDescription.DefaultValue = false;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_STOCK_CFAO";
                    fieldDescription.Name = "*Stock CFAO";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Boolean;
                    fieldDescription.DefaultValue = false;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_STOCK_NAME";
                    fieldDescription.Name = "*Nom stock";
                    fieldDescription.Editable = FieldDescriptionEditableType.NoEditable;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_IS_OMMITED";
                    fieldDescription.Name = "*Is Omitted";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Boolean;
                    fieldDescription.DefaultValue = false;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_GPAO_FILE";
                    fieldDescription.Name = "*GpaoFile";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_NESTING_NAME";
                    fieldDescription.Name = "*nom_du_createur";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_TO_CUT_SHEET";
                    fieldDescription.Name = "*id de la tole a couper";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Int;
                    fieldDescription.DefaultValue = 0;
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
            #region Pièces 2D
            
            {
                IEntityType entityType = context.Kernel.GetEntityType("_REFERENCE");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType , null, "_REFERENCE", null);
                entityTypeFactory.Key = "_REFERENCE";
                entityTypeFactory.Name = "Pièces 2D";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;
                
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_PIECE_TOLERANCES";
                    fieldDescription.Name = "Tolérances";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AFFAIRE";
                    fieldDescription.Name = "*Affaire_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "REMONTER_DT";
                    fieldDescription.Name = "*Remonter_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Boolean;
                    fieldDescription.DefaultValue = false;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "FAMILY";
                    fieldDescription.Name = "*Famille_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "IDLNROUT";
                    fieldDescription.Name = "*Id gamme_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "CUSTOMER";
                    fieldDescription.Name = "*Client_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "PLAN";
                    fieldDescription.Name = "*Plan_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "IDMAT";
                    fieldDescription.Name = "*Matiere_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "IDLNBOM";
                    fieldDescription.Name = "*Nomenclature_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_CDE";
                    fieldDescription.Name = "*Commande_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "EN_RANG";
                    fieldDescription.Name = "*Rang_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "EN_PERE_PIECE";
                    fieldDescription.Name = "*Repere piece_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "FORMATCLIP";
                    fieldDescription.Name = "*Format_dt";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
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
                
            }
            
            #endregion
            #region Pièces à produire
            
            {
                IEntityType entityType = context.Kernel.GetEntityType("_TO_PRODUCE_REFERENCE");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType , null, "_ACTCUT", null);
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
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "FAMILY";
                    fieldDescription.Name = "*Famille";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "IDLNROUT";
                    fieldDescription.Name = "*Idlnrout";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "CENTREFRAIS";
                    fieldDescription.Name = "*Centre frais";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Lookup;
                    fieldDescription.LinkKey = "_CENTRE_FRAIS";
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "ECOQTY";
                    fieldDescription.Name = "*Ecoqty";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Int;
                    fieldDescription.DefaultValue = 0;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "STARTDATE";
                    fieldDescription.Name = "*Debut";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Date;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "ENDDATE";
                    fieldDescription.Name = "*Fin";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Date;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "IDLNBOM";
                    fieldDescription.Name = "*Nomenclature";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "PLAN";
                    fieldDescription.Name = "*Plan";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "IDMAT";
                    fieldDescription.Name = "*Id matiere";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "EN_RANG";
                    fieldDescription.Name = "*Rang";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "EN_PERE_PIECE";
                    fieldDescription.Name = "*Repere pieces";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "MATERIAL_CLIPPER";
                    fieldDescription.Name = "*Matiere clip";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "FORMATCLIP";
                    fieldDescription.Name = "*Format clip";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "NUMMAG";
                    fieldDescription.Name = "*Numéro mag";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "LEVQA";
                    fieldDescription.Name = "*Qualité";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "CENTREFRAISSUIV";
                    fieldDescription.Name = "*Centre de frais suivant";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "DELAIS_INT";
                    fieldDescription.Name = "*Délais interne";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "MACHINE_FROM_CF";
                    fieldDescription.Name = "*Machine suggérée par l of";
                    fieldDescription.Editable = FieldDescriptionEditableType.NoEditable;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "SANS_DT";
                    fieldDescription.Name = "*Remonter le Dossier Technique";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Boolean;
                    fieldDescription.DefaultValue = false;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "NEED_PREP";
                    fieldDescription.Name = "*A traiter";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Boolean;
                    fieldDescription.DefaultValue = false;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "OF_IMPORT_NUMBER";
                    fieldDescription.Name = "*Import Number";
                    fieldDescription.Editable = FieldDescriptionEditableType.NoEditable;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "GEOMETRY_FROM_OF";
                    fieldDescription.Name = "*Géometrie de l'of";
                    fieldDescription.Editable = FieldDescriptionEditableType.NoEditable;
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
                
            }
            
            #endregion
            return true;
        }
    }
}
