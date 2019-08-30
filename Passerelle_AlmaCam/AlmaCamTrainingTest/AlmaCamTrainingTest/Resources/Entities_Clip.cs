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


            //preparation du dictionnaire //
            // champs des toles//
            var sheet_Dictionnary = new Dictionary<string, NewField>();
            var machine_Dictionnary = new Dictionary<string, NewField>();
            var stock_Dictionnary = new Dictionary<string, NewField>();
            var part2d_Dictionnary = new Dictionary<string, NewField>();
            var part_to_produce_Dictionnary = new Dictionary<string, NewField>();
            var matiere_Dictionnary = new Dictionary<string, NewField>();
            //tole//
            var fielddesc = new NewField("FILENAME", "Clip filename", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            sheet_Dictionnary.Add("FILENAME", fielddesc);
            //centre frais//   
            fielddesc = new NewField("CENTREFRAIS_MACHINE", "Clip centre de frais", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Lookup, "_CENTRE_FRAIS");
            machine_Dictionnary.Add("CENTREFRAIS_MACHINE", fielddesc);
            //stock//			
            fielddesc = new NewField("FILENAME", "Clip Filename", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("FILENAME", fielddesc);
            fielddesc = new NewField("QTE_TOT", "Clip Quté", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Int, "");
            stock_Dictionnary.Add("QTE_TOT", fielddesc);
            fielddesc = new NewField("NUMCERTIF", "Clip certif.", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("NUMCERTIF", fielddesc);
            fielddesc = new NewField("IDCLIP", "Clip identification", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("IDCLIP", fielddesc);
            fielddesc = new NewField("NUMMATLOT", "Clip matiere lotie", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("NUMMATLOT", fielddesc);
            fielddesc = new NewField("NUMMAG", "Clip magasin", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("NUMMAG", fielddesc);
            fielddesc = new NewField("GISEMENT", "Clip gisement", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("GISEMENT", fielddesc);
            fielddesc = new NewField("NUMLOT", "Clip lotie", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("NUMLOT", fielddesc);
            fielddesc = new NewField("STOCK_IMPORT_NUMBER", "Clip numero import", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("STOCK_IMPORT_NUMBER", fielddesc);
            fielddesc = new NewField("NAF_LOTIE", "Clip affaire lotie", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("NAF_LOTIE", fielddesc);
            fielddesc = new NewField("AF_STOCK_RENAME", "Clip rename", FieldDescriptionEditableType.NoEditable, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Boolean, "");
            stock_Dictionnary.Add("AF_STOCK_RENAME", fielddesc);
            fielddesc = new NewField("AF_STOCK_NAME", "Clip nom du stock", FieldDescriptionEditableType.NoEditable, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("AF_STOCK_NAME", fielddesc);
            fielddesc = new NewField("AF_STOCK_CFAO", "Clip Cfao", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Boolean, "");
            stock_Dictionnary.Add("AF_STOCK_CFAO", fielddesc);
            fielddesc = new NewField("AF_IS_OMMITED", "Clip omitted", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Boolean, "");
            stock_Dictionnary.Add("AF_IS_OMMITED", fielddesc);
            fielddesc = new NewField("AF_GPAO_FILE", "Clip gpao file", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("AF_GPAO_FILE", fielddesc);
            fielddesc = new NewField("AF_NESTING_NAME", "Clip nesting name", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            stock_Dictionnary.Add("AF_NESTING_NAME", fielddesc);
            fielddesc = new NewField("AF_TO_CUT_SHEET", "Clip to cut sheet id", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Int, "");
            stock_Dictionnary.Add("AF_TO_CUT_SHEET", fielddesc);
            fielddesc = new NewField("AF_IS_MULTIDIM", "Miltidim", FieldDescriptionEditableType.NoEditable, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Boolean, "");
            stock_Dictionnary.Add("AF_IS_MULTIDIM", fielddesc);
            //part2d//
            fielddesc = new NewField("AF_PIECE_TOLERANCES", "Clip tolerance", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("AF_PIECE_TOLERANCES", fielddesc);
            fielddesc = new NewField("AFFAIRE", "Clip affaire", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("AFFAIRE", fielddesc);
            fielddesc = new NewField("REMONTER_DT", "Clip to cut sheet id", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Boolean, "");
            part2d_Dictionnary.Add("REMONTER_DT", fielddesc);
            fielddesc = new NewField("FAMILY", "Clip famille", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("FAMILY", fielddesc);
            fielddesc = new NewField("IDLNROUT", "Clip gamme", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("IDLNROUT", fielddesc);
            fielddesc = new NewField("CUSTOMER", "Clip customer", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("CUSTOMER", fielddesc);
            fielddesc = new NewField("PLAN", "Clip plan", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("PLAN", fielddesc);
            fielddesc = new NewField("IDMAT", "Clip mat", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("IDMAT", fielddesc);
            fielddesc = new NewField("IDLNBOM", "Clip nommenclature", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("IDLNBOM", fielddesc);
            fielddesc = new NewField("AF_CDE", "Clip commande", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("AF_CDE", fielddesc);
            fielddesc = new NewField("EN_RANG", "Clip rang", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("EN_RANG", fielddesc);
            fielddesc = new NewField("EN_PERE_PIECE", "Clip repere piece", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("EN_PERE_PIECE", fielddesc);
            fielddesc = new NewField("FORMATCLIP", "Clip format clip", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part2d_Dictionnary.Add("FORMATCLIP", fielddesc);

            //part2prod//
            fielddesc = new NewField("AFFAIRE", "Clip affaire", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("AFFAIRE", fielddesc);
            fielddesc = new NewField("FAMILY", "Clip famille", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("FAMILY", fielddesc);
            fielddesc = new NewField("IDLNROUT", "Clip gamme", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("IDLNROUT", fielddesc);
            fielddesc = new NewField("CENTREFRAIS", "Clip centre frais", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Lookup, "_CENTRE_FRAIS");
            part_to_produce_Dictionnary.Add("CENTREFRAIS", fielddesc);
            fielddesc = new NewField("ECOQTY", "Clip ecoqty", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("ECOQTY", fielddesc);
            fielddesc = new NewField("AF_STARTDATE", "Clip date debut", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Date, "");
            part_to_produce_Dictionnary.Add("STARTDATE", fielddesc);
            fielddesc = new NewField("ENDDATE", "Clip date fin", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Date, "");
            part_to_produce_Dictionnary.Add("ENDDATE", fielddesc);
            fielddesc = new NewField("IDLNBOM", "Clip nomenclature", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("IDLNBOM", fielddesc);
            fielddesc = new NewField("PLAN", "Clip plan", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("PLAN", fielddesc);
            fielddesc = new NewField("IDMAT", "Clip mat", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("IDMAT", fielddesc);
            fielddesc = new NewField("EN_RANG", "Clip en rang", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("EN_RANG", fielddesc);
            fielddesc = new NewField("EN_PERE_PIECE", "Clip repere piece", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("EN_PERE_PIECE", fielddesc);
            fielddesc = new NewField("MATERIAL_CLIPPER", "Clip material", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("MATERIAL_CLIPPER", fielddesc);
            fielddesc = new NewField("FORMATCLIP", "Clip format", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("FORMATCLIP", fielddesc);
            fielddesc = new NewField("NUMMAG", "Clip magasin", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("NUMMAG", fielddesc);
            fielddesc = new NewField("LEVQA", "Clip qualite", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("LEVQA", fielddesc);
            fielddesc = new NewField("CENTREFRAISSUIV", "Clip centre frais suivant", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("CENTREFRAISSUIV", fielddesc);
            fielddesc = new NewField("DELAIS_INT", "Clip delais int", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("DELAIS_INT", fielddesc);
            fielddesc = new NewField("MACHINE_FROM_CF", "Clip machine", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("MACHINE_FROM_CF", fielddesc);
            fielddesc = new NewField("SANS_DT", "Clip sans dt", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Boolean, "");
            part_to_produce_Dictionnary.Add("SANS_DT", fielddesc);
            fielddesc = new NewField("NEED_PREP", "Clip need prep", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Boolean, "");
            part_to_produce_Dictionnary.Add("NEED_PREP", fielddesc);
            fielddesc = new NewField("OF_IMPORT_NUMBER", "Clip numero import", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("OF_IMPORT_NUMBER", fielddesc);
            fielddesc = new NewField("GEOMETRY_FROM_OF", "Clip from of", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            part_to_produce_Dictionnary.Add("GEOMETRY_FROM_OF", fielddesc);

            //matiere//
            //fielddesc = new NewField("AF_DEFAULT_SHEET", "matiere par defaut", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.Link, "_SHEET");
            fielddesc = new NewField("AF_DEFAULT_SHEET", "Reference matiere par defaut", FieldDescriptionEditableType.AllSection, FieldDescriptionVisibilityType.AllSection, FieldDescriptionType.String, "");
            matiere_Dictionnary.Add("AF_DEFAULT_SHEET", fielddesc);
            //





            ////////////////////////////////////////////////////////////////////////////////////////////
            //sheet;
            #region Tôles
            {
                IEntityType entityType = context.Kernel.GetEntityType("_SHEET");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType, null, "_SUPPLY", null);
                entityTypeFactory.Key = "_SHEET";
                entityTypeFactory.Name = "Tôles";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;



                foreach (var item in sheet_Dictionnary)
                {

                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);

                    fieldDescription.Key = item.Key;
                    fieldDescription.Name = sheet_Dictionnary[item.Key].Name;
                    fieldDescription.Editable = sheet_Dictionnary[item.Key].Editable;
                    fieldDescription.Visible = sheet_Dictionnary[item.Key].Visible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = sheet_Dictionnary[item.Key].Type;

                    if (sheet_Dictionnary[item.Key].Type == FieldDescriptionType.Lookup)
                    {
                        fieldDescription.LinkKey = sheet_Dictionnary[item.Key].Link;
                    }


                    if (entityType.FieldList.ContainsKey(item.Key) == false)
                    {
                        entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    }
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

            ////////////////////////////////////////////////////////////////////////////////////////////
            //machine;

            #region machine
            {
                IEntityType entityType = context.Kernel.GetEntityType("_CUT_MACHINE_TYPE");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType, null, "_RSC", null);
                entityTypeFactory.Key = "_CUT_MACHINE_TYPE";
                entityTypeFactory.Name = "Machines de coupe";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;



                foreach (var item in machine_Dictionnary)
                {

                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);

                    fieldDescription.Key = item.Key;
                    fieldDescription.Name = machine_Dictionnary[item.Key].Name;
                    fieldDescription.Editable = machine_Dictionnary[item.Key].Editable;
                    fieldDescription.Visible = machine_Dictionnary[item.Key].Visible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = machine_Dictionnary[item.Key].Type;

                    if (machine_Dictionnary[item.Key].Type == FieldDescriptionType.Lookup)
                    {
                        fieldDescription.LinkKey = machine_Dictionnary[item.Key].Link;
                    }


                    if (entityType.FieldList.ContainsKey(item.Key) == false)
                    {
                        entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    }
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


            ////////////////////////////////////////////////////////////////////////////////////////////
            //stock;
            #region stock
            {
                IEntityType entityType = context.Kernel.GetEntityType("_STOCK");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType, null, "_SUPPLY", null);
                entityTypeFactory.Key = "_STOCK";
                entityTypeFactory.Name = "Stock";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;



                foreach (var item in stock_Dictionnary)
                {

                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);

                    fieldDescription.Key = item.Key;
                    fieldDescription.Name = stock_Dictionnary[item.Key].Name;
                    fieldDescription.Editable = stock_Dictionnary[item.Key].Editable;
                    fieldDescription.Visible = stock_Dictionnary[item.Key].Visible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = stock_Dictionnary[item.Key].Type;

                    if (stock_Dictionnary[item.Key].Type == FieldDescriptionType.Lookup)
                    {
                        fieldDescription.LinkKey = stock_Dictionnary[item.Key].Link;
                    }


                    if (entityType.FieldList.ContainsKey(item.Key) == false)
                    {
                        entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    }
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
            ////////////////////////////////////////////////////////////////////////////////////////////
            //sheet;
            #region p2d
            {
                IEntityType entityType = context.Kernel.GetEntityType("_REFERENCE");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType, null, "_REFERENCE", null);
                entityTypeFactory.Key = "_REFERENCE";
                entityTypeFactory.Name = "Pièces 2D";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;



                foreach (var item in part2d_Dictionnary)
                {

                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);

                    fieldDescription.Key = item.Key;
                    fieldDescription.Name = part2d_Dictionnary[item.Key].Name;
                    fieldDescription.Editable = part2d_Dictionnary[item.Key].Editable;
                    fieldDescription.Visible = part2d_Dictionnary[item.Key].Visible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = part2d_Dictionnary[item.Key].Type;

                    if (part2d_Dictionnary[item.Key].Type == FieldDescriptionType.Lookup)
                    {
                        fieldDescription.LinkKey = part2d_Dictionnary[item.Key].Link;
                    }


                    if (entityType.FieldList.ContainsKey(item.Key) == false)
                    {
                        entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    }
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
            ////////////////////////////////////////////////////////////////////////////////////////////
            //sheet;
            #region to_product_ref
            {
                IEntityType entityType = context.Kernel.GetEntityType("_TO_PRODUCE_REFERENCE");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType, null, "_ACTCUT", null);
                entityTypeFactory.Key = "_TO_PRODUCE_REFERENCE";
                entityTypeFactory.Name = "Pièces à produire";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;



                foreach (var item in part_to_produce_Dictionnary)
                {

                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);

                    fieldDescription.Key = item.Key;
                    fieldDescription.Name = part_to_produce_Dictionnary[item.Key].Name;
                    fieldDescription.Editable = part_to_produce_Dictionnary[item.Key].Editable;
                    fieldDescription.Visible = part_to_produce_Dictionnary[item.Key].Visible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = part_to_produce_Dictionnary[item.Key].Type;

                    if (part_to_produce_Dictionnary[item.Key].Type == FieldDescriptionType.Lookup)
                    {
                        fieldDescription.LinkKey = part_to_produce_Dictionnary[item.Key].Link;
                    }


                    if (entityType.FieldList.ContainsKey(item.Key) == false)
                    {
                        entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    }
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

            
            /////////////
            //matiere
            ////
            ////////////////////////////////////////////////////////////////////////////////////////////
            //machine;

            #region matiere
            {
                IEntityType entityType = context.Kernel.GetEntityType("_MATERIAL");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType, null, "_ACTCUT", null);
                entityTypeFactory.Key = "_MATERIAL";
                entityTypeFactory.Name = "Matière";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;



                foreach (var item in matiere_Dictionnary)
                {

                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);

                    fieldDescription.Key = item.Key;
                    fieldDescription.Name = matiere_Dictionnary[item.Key].Name;
                    fieldDescription.Editable = matiere_Dictionnary[item.Key].Editable;
                    fieldDescription.Visible = matiere_Dictionnary[item.Key].Visible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = matiere_Dictionnary[item.Key].Type;

                    if (matiere_Dictionnary[item.Key].Type == FieldDescriptionType.Lookup)
                    {
                        fieldDescription.LinkKey = matiere_Dictionnary[item.Key].Link;
                    }
                    if (matiere_Dictionnary[item.Key].Type == FieldDescriptionType.Link)
                    {
                        fieldDescription.LinkKey = matiere_Dictionnary[item.Key].Link;
                    }

                    if (entityType.FieldList.ContainsKey(item.Key) == false)
                    {
                        entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    }
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


}
