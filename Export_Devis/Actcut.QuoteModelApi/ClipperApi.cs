using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;

using Alma.BaseUI.Utils;

using Wpm.Schema.Kernel;
using Wpm.Implement.Manager;
using Wpm.Implement.ComponentEditor;

using Actcut.QuoteModel;
using Actcut.QuoteModelManager;
using Actcut.ActcutModelManagerUI;

namespace Actcut.ActcutClipperApi
{
    public class ClipperApi : IClipperApi
    {
        IContext _Context = null;
        bool _UserOk = false;

        #region IClipperApi Membres

        public bool Init(IContext context)
        {
            _Context = context;
            _UserOk = (_Context.UserId > 0);
            return true;
        }

        public bool Init(string databaseName, string user)
        {
            ModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(databaseName);

            if (_Context != null)
            {
                Licence.InitLicence(_Context.Kernel, null);
                _UserOk = SetUser(_Context, user);
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool ExportQuote(long quoteNumber, string orderNumber, string exportFile)
        {
            if (_Context != null)
            {
                IQuoteManagerUI quoteManagerUI = new QuoteManagerUI();
                IEntity quoteEntity = quoteManagerUI.GetQuoteEntity(_Context, quoteNumber);
                bool ret = quoteManagerUI.AccepQuote(_Context, quoteEntity, orderNumber, exportFile);
                return (ret ? true : false);
            }
            else
            {
                return false;
            }
        }
        public bool GetQuote(out long quoteNumberReference)
        {
            quoteNumberReference = -1;
            if (_Context != null)
            {
                IEntity quoteEntity = null;
                IEntitySelector entitySelector = new EntitySelector();
                entitySelector.Init(_Context, _Context.Kernel.GetEntityType("_QUOTE_SENT"));
                entitySelector.MultiSelect = false;
                entitySelector.ShowPropertyBox = false;
                if (entitySelector.ShowDialog() == DialogResult.OK)
                    quoteEntity = entitySelector.SelectedEntity.FirstOrDefault();

                if (_UserOk) _Context.SaveUserModel();

                if (quoteEntity != null)
                {
                    string quoteReference = quoteEntity.GetFieldValueAsString("_REFERENCE");
                    long.TryParse(quoteReference, out quoteNumberReference);
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        #endregion

        private bool SetUser(IContext context, string user)
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
