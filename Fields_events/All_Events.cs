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
    public partial class ImportCustomizedFormulaEventType : ScriptModelCustomization, IScriptModelCustomization
    {
        public override bool Execute(IContext context, IContext hostContext)
        {
            ITransaction transaction = context.CreateTransaction();
            IParameterValue parameterValue;
            #region Formule pour le champ "Référence" d'un devis
            
            {
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_QUOTE_REFERENCE_TYPE");
                parameterValue.SetValue(1);
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_QUOTE_REFERENCE_FORMULA");
                parameterValue.SetValue(@"using System;
using System.Diagnostics;

using Wpm.Implement.Manager;

namespace Actcut.QuoteModelManager
{
    public class MyQuoteReferenceFormula : QuoteReferenceFormula
    {
        public override string Execute(IEntity quoteEntity, bool copy, IEntity sourceQuoteEntity)
        {
            string reference = """";
			quoteEntity.SetFieldValue(""_ACCEPTANCE_PERIOD"",5);
			Int64 offset= quoteEntity.Context.ParameterSetManager.GetParameterValue(""_EXPORT"", ""_CLIPPER_QUOTE_NUMBER_OFFSET"").GetValueAsLong();
			Int64 iddevis=quoteEntity.GetFieldValueAsLong(""_INC_NO"")+offset;
				
            if (copy && sourceQuoteEntity != null && quoteEntity != null)
            {
                //string sourceReference = sourceQuoteEntity.GetFieldValueAsString(""_REFERENCE"");
            		reference = string.Format(""{0}"",iddevis.ToString());
                //reference = string.Format(""{0}-{1}"", quoteEntity.GetFieldValueAsLong(""_INC_NO""), sourceReference);
				//reference = string.Format(""{0}"", quoteEntity.GetFieldValueAsLong(""_INC_NO""), sourceReference);

            }
            else if (quoteEntity != null)
            {
                //reference = string.Format(""{0}"", quoteEntity.GetFieldValueAsLong(""_INC_NO""));
				reference = string.Format(""{0}"",iddevis.ToString());
             
            }
            return reference;
        }
    }
}
");
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_QUOTE_REFERENCE_FILENAME");
                parameterValue.SetValue(null);
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_QUOTE_REFERENCE_NAMESPACE");
                parameterValue.SetValue(null);
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_QUOTE_REFERENCE_CLASS");
                parameterValue.SetValue(null);
                parameterValue.Save(transaction);
            }
            
            #endregion
            #region Evénement après l'envoi à la coupe
            
            {
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_AFTER_SEND_TO_WORKSHOP_EVENT_TYPE");
                parameterValue.SetValue(2);
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_AFTER_SEND_TO_WORKSHOP_EVENT_FILENAME");
                parameterValue.SetValue(@"AF_Clipper_Dll.dll");
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_AFTER_SEND_TO_WORKSHOP_EVENT_NAMESPACE");
                parameterValue.SetValue(@"AF_Clipper_Dll");
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_AFTER_SEND_TO_WORKSHOP_EVENT_CLASS");
                parameterValue.SetValue(@"Clipper_8_DoOnAction_AfterSendToWorkshop");
                parameterValue.Save(transaction);
            }
            
            #endregion
            #region Evénement après la fin de la coupe
            
            {
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_AFTER_END_CUTTING_EVENT_TYPE");
                parameterValue.SetValue(2);
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_AFTER_END_CUTTING_EVENT_FILENAME");
                parameterValue.SetValue(@"AF_Clipper_Dll.dll");
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_AFTER_END_CUTTING_EVENT_NAMESPACE");
                parameterValue.SetValue(@"AF_Clipper_Dll");
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_AFTER_END_CUTTING_EVENT_CLASS");
                parameterValue.SetValue(@"Clipper_8_DoOnAction_After_Cutting_end");
                parameterValue.Save(transaction);
            }
            
            #endregion
            #region Evénement avant la restauration d'un placement
            
            {
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_BEFORE_NESTING_RESTORE_EVENT_TYPE");
                parameterValue.SetValue(2);
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_BEFORE_NESTING_RESTORE_EVENT_FILENAME");
                parameterValue.SetValue(@"AF_Clipper_Dll.dll");
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_BEFORE_NESTING_RESTORE_EVENT_NAMESPACE");
                parameterValue.SetValue(@"AF_Clipper_Dll");
                parameterValue.Save(transaction);
                parameterValue = context.ParameterSetManager.GetParameterValue("_EVENT_FORMULA_HANDLER", "_BEFORE_NESTING_RESTORE_EVENT_CLASS");
                parameterValue.SetValue(@"Clipper_8_Before_Nesting_Restore_Event");
                parameterValue.Save(transaction);
            }
            
            #endregion
            transaction.Save();
            return true;
        }
    }
}
