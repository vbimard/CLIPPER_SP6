using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
//using Alma.BaseUI.Utils;
using Wpm.Schema.Kernel;
//using Wpm.Implement.Manager;
//using Wpm.Implement.ComponentEditor;

using Actcut.QuoteModel;
using Actcut.QuoteModelManager;
using Actcut.ActcutModelManagerUI;
using Wpm.Implement.Manager;
using Wpm.Implement.ComponentEditor;

namespace AF_Actcut.ActcutClipperApi
{
    public class ClipperApi : IClipperApi,IDisposable
    {
        IContext _Context = null;
        bool _UserOk = false;

        #region IClipperApi Membres

        public bool ConnectAlmaCamContext(IContext context)
        {
            _Context = context;
            _UserOk = (_Context.UserId > 0);
            return true;
        }

        public  void Dispose()
        {
            _Context = null;
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// supprime le contexte ouvert par almacam et laisse la gestion memoire au garbage collector
        /// false si oas de context ouvert
        /// </summary>
        /// <returns>true/false</returns>
        public bool ExitAlmaCam()
        {

            if (_Context != null)
            {
                Dispose();
                return true;
            }
            else
            {
              ///aucune session ouverte
                return false;
            }
           
        }

       
        /// <summary>
        /// connexion a la base almacam
        /// </summary>
        /// <param name="databaseName">nom de la base</param>
        /// <param name="user">utilisateur</param>
        /// <returns></returns>
        public bool ConnectAlmaCamDatabase(string databaseName, string user)
        {

            
            ModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(databaseName);
            

            if (_Context != null)
            {
                Licence.InitLicence(_Context.Kernel, null);
                if (SetUser(_Context, user)) { return true; } else { MessageBox.Show("User Inconnu"); return false; }//_UserOk = SetUser(_Context, user);
         
            }
            else
            {
                return false;
            }
        }

       
        public bool ExportQuote(long quoteNumber, string orderNumber, string exportFile)

        {
            bool ret = false;
            if (_Context != null)
            {
                IQuoteManagerUI quoteManagerUI = new QuoteManagerUI();

                
                IEntity quoteEntity = quoteManagerUI.GetQuoteEntity(_Context, quoteNumber);
                if (Control_quote_Integrity(quoteEntity))
                {
                    ret = quoteManagerUI.AccepQuote(_Context, quoteEntity, orderNumber, exportFile);
                }
            }
            return ret;
        }

        /// <summary>
        /// boite de dialogue quote de selection des devis
        /// </summary>
        /// <param name="quoteNumberReference"></param>
        /// <returns></returns>
        public bool SelectQuoteUI(out long quoteNumberReference)
        {
            bool rst=false;
            quoteNumberReference = -1;
            long quoteid=0;
            try { 
            
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
                { //control de l'integrité des donnees
                        rst = Control_quote_Integrity(quoteEntity);
                        if (rst == true)
                        {
                            quoteid = quoteEntity.Id32;
                            string quoteref = "0";
                            quoteref = quoteEntity.GetFieldValueAsString("_REFERENCE");
                            quoteNumberReference = Convert.ToInt64(quoteref);
                        }
                    //Convert.ToInt64( quoteEntity.GetFieldValueAsString("_REFERENCE"));
                                                
                       

                    return rst;
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

            //catch (Exception ie) { return rst; }
            catch (FormatException ex ) {
                string text = "Impossible de faire d'import sur ce devis.\r\nLa reference Almacam du devis " + quoteid + 
                    " contient des caracteres alphanumeriques.";
                MessageBox.Show(text, "SelectQuoteUI");
                return rst; }
        }

        public bool Control_quote_Integrity(IEntity quoteEntity)
        {
            bool rst=true; //valide par defaut
            quoteEntity = quoteEntity.Context.EntityManager.GetEntity(quoteEntity.Id32, "_QUOTE_REQUEST"); //AF_ImportTools.SimplifiedMethods.GetFirtOfList(quotes);
            ITransaction transaction = quoteEntity.Context.CreateTransaction();
            IQuote iquote = new Quote(transaction, quoteEntity);

            try
            {

                ///vrefication des pieces
                ///
                bool sansDoublons = true;
                List<String> ListSansDuplication = new List<String>();
                foreach (IEntity partEntity in iquote.QuotePartList)
                {

                    if (!ListSansDuplication.Contains(partEntity.GetFieldValueAsString("_REFERENCE")))
                        ListSansDuplication.Add(partEntity.GetFieldValueAsString("_REFERENCE"));
                    else
                    {
                        sansDoublons = false;
                        string text = "Impossible de faire d'import du devis " + quoteEntity.Id32 + ".\r\nCe devis contient plusieurs pieces du meme nom (reference) ";
                        MessageBox.Show(text, "Control_quote_Integrity");
                        break;
                    }

                }
                ///verification client simple message
                string risk = quoteEntity.GetFieldValueAsEntity("_FIRM").GetFieldValueAsString("_FINANCIAL_RISK");
                if (risk != "0") { MessageBox.Show("Attention !! le  client " + quoteEntity.GetFieldValueAsEntity("_FIRM").GetFieldValueAsString("_NAME") + " possede des conditions de blocage sur Clipper", "Control_quote_Integrity"); }
               
                ///verification du format de l'id du devis, il ne doi pas contenir de caracteres alphanumerique
                //quoteid = quoteEntity.Id32;
                string quoteref = "0";
                quoteref = quoteEntity.GetFieldValueAsString("_REFERENCE");
                Convert.ToInt64(quoteref);



                rst = rst && sansDoublons;

                return rst;

            }

            catch (FormatException ex)
            {
                string text = "Impossible de faire d'import sur ce devis.\r\nLa reference Almacam du devis selectionné contient des caracteres alphanumeriques.";
                MessageBox.Show(text, "Control_quote_Integrity");
                return rst;
            }
            catch { return rst; }
            finally { }
            
        }



        /// <summary>
        /// boite de dialogue quote de selection des devis
        /// </summary>
        /// <param name="quoteNumberReference"></param>
        /// <returns></returns>
        public bool GetQuoteList(out string JsonQuoteList)
        {
            JsonQuoteList = "";
            if (_Context != null)
            {

                IEntityList quoteEntityList = null;
                quoteEntityList = _Context.EntityManager.GetEntityList("_QUOTE_SENT");
                
                

                if (quoteEntityList != null)
                {
                    //string quoteReference = quoteEntityList.GetFieldValueAsString("_REFERENCE");
                    //long.TryParse(quoteReference, out JsonQuoteList);
                    foreach (IEntity quoteentity in quoteEntityList)
                    {

                        string reference = quoteentity.GetFieldValueAsString("_REFERENCE");
                        //string user = quoteentity.GetFieldValueAsString("_REFERENCE");
                        //string date = quoteentity.GetFieldValueAsString("_REFERENCE");

                    }
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
