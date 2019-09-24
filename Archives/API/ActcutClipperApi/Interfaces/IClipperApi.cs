using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

using Alma.NetKernel.Tools;
using Alma.BaseUI.ErrorMessage;

using Wpm.Implement.Manager;
using Wpm.Schema.Kernel;

namespace AF_Actcut.ActcutClipperApi
{
    public interface IClipperApi
    {
        /// <summary>
        /// se connecte au context almacam
        /// </summary>
        /// <param name="Context"></param>
        /// <returns></returns>
        bool ConnectAlmaCamContext(IContext Context);
        /// <summary>
        /// ouvre l'acces à la base : methode à lancer au demarrage de clipper 
        /// ou a la premiere ouverture de la fenetre de commande
        /// </summary>
        /// <param name="DatabaseName">nom de la base almacam</param>
        /// <param name="User">nom de l'utilisateur</param>
        /// <returns>true/false</returns>
       // bool Init(string DatabaseName, string User);
        bool ConnectAlmaCamDatabase(string DatabaseName, string User);

        /// <summary>
        /// Export le devis au format texte 
        /// Permet la saisie de la commande 
        /// Accepte le devis dans almacam
        ///  
        /// </summary>
        /// <param name="QuoteNumber"></param>
        /// <param name="OrderNumber"></param>
        /// <param name="ExportFile"></param>
        /// <returns></returns>
        bool ExportQuote(long QuoteNumber, string OrderNumber, string ExportFile);
        /// <summary>
        /// ouvre la fenetre de selection des devis en cours dans almacam
        /// </summary>
        /// <param name="QuoteNumberReference">retourne le numero du devis selectionné</param>
        /// <returns>true/false</returns>
        bool SelectQuoteUI(out long QuoteNumberReference);
        /// <summary>
        /// recupere une liste Json des devis en cours
        /// </summary>
        /// <param name="JsonQuoteList">String Json</param>
        /// <returns>true/false</returns>
        bool GetQuoteList(out string JsonQuoteList);
         ///pensez a mettre ‘CMDNULL’ dans le ordernumber si vous appelez la methode  ExportQuote(10, 0, « »);







    }
}
