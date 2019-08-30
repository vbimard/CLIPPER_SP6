using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.Odbc;
using System.IO;
using Wpm.Implement.Processor;
using AF_ImportTools;
using Wpm.Implement.Manager;
using Wpm.Schema.Kernel;
using System.Diagnostics;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Win32;
using System.Runtime.Serialization;
using System.Data.Common;
using Actcut.CommonModel;
using AF_Clipper_Dll;

namespace AF_Import_ODBC_Clipper_AlmaCam

{

    #region  commandes ou boutons intefacés sur almacam
    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <summary>
    /// creation des commandes d'imports
    /// </summary>
    //section des commandes
    /// <summary>
    /// bouton d'import des matieres
    /// </summary>
    public class Clipper8_ImportMatiere_Processor : CommandProcessor
    {
        /// <summary>
        /// main methode
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {

            try
            {
                //verification de la connexion obdc
                if (OdbcTools.Test_Connexion())
                {
                    Import(Context);
                    AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT  des Matieres ", "Import terminé.");
                }
                return base.Execute();
            }
            catch
            {
                return base.Execute();
            }

        }

        /////////////////////////////////////////////////////////////////
        /// <summary>
        /// import des entité clipper
        /// </summary>
        /// <param name="contextcontextlocal"></param>
        /// 
        public void Import(IContext contextcontextlocal)
        {

            //TextWriterTraceListener logFile;
            try
            {
                //creation des logs
                using (Clipper_Import_Matiere matiere = new Clipper_Import_Matiere(contextcontextlocal))
                {
                    try
                    {

                        Cursor.Current = Cursors.WaitCursor;

                        matiere.Import();

                        Cursor.Current = Cursors.Default;


                    }
                    catch { Cursor.Current = Cursors.Default; }
                }


            }
            catch (Exception ie)
            {
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
    }

    /// <summary>
    /// bouton d'import des toles
    /// attention l'import des toles est fait par l'import du stock clipper pour le moment.
    /// </summary>
    public class Clipper_Import_Toles_Processor : CommandProcessor
    {
        /// <summary>
        /// main methode
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {

            try
            {

                if (OdbcTools.Test_Connexion())
                {
                    Import(Context);
                    AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT  des Matieres ", "Import terminé.");
                }
               
                return base.Execute();
            }
            catch
            {
                return base.Execute();
            }

        }
        public bool Import(IContext contextlocal)
        {
            try
            {
                using (Clipper_Article_Toles_MonoDim Toles_MonoDim = new Clipper_Article_Toles_MonoDim(contextlocal))
                {
                    try
                    {


                        Cursor.Current = Cursors.WaitCursor;


                        Toles_MonoDim.Read();
                        Toles_MonoDim.Write();
                        Toles_MonoDim.Close();

                        Cursor.Current = Cursors.Default;
                    }
                    catch { Cursor.Current = Cursors.Default; }
                }

                return true;
            }
            catch { return false; }
        }
    }
    /// <summary>
    /// bouton d'import des vis
    /// </summary>
    public class Clipper_Import_Fournitures_Divers_Processor : CommandProcessor
    {



        /// <summary>
        /// constructeur de l'import des tube, tu stock et des matieres.........
        /// import_matiere,  tube_rond,  rond,  tube_rectangle,  tube_carre, tube_flat,  Tube_Speciaux
        /// </summary>
        /// <param name="import_matiere">true ou false pour declencher l'import</param>
        /// <param name="tube_rond">true ou false pour declencher l'import</param>
        /// <param name="rond">true ou false pour declencher l'import</param>
        /// <param name="tube_rectangle">true ou false pour declencher l'import</param>
        /// <param name="tube_carre">true ou false pour declencher l'import</param>
        /// <param name="tube_flat">true ou false pour declencher l'import</param>
        /// <param name="Tube_Speciaux">true ou false pour declencher l'import</param>
        /// <param name="Fourniture">true ou false pour declencher l'import</param>
        /// readonly possible
        private Boolean Import_Matiere = false;
        private Boolean Tube_Rond = false;
        private Boolean Tube_Speciaux = false;
        private Boolean Rond = false;
        private Boolean Tube_Rectangle = false;
        private Boolean Tube_Carre = false;
        private Boolean Tube_Flat = false;
        private Boolean Fourniture = true;

        public Clipper_Import_Fournitures_Divers_Processor(bool import_matiere, bool tube_rond, bool rond, bool tube_rectangle, bool tube_carre, bool tube_flat, bool tube_speciaux, bool fourniture)
        {

            Import_Matiere = import_matiere;
            Tube_Rond = tube_rond;
            Rond = rond;
            Tube_Rectangle = tube_rectangle;
            Tube_Carre = tube_carre;
            Tube_Flat = tube_flat;
            Tube_Speciaux = tube_speciaux;
            Fourniture = fourniture;


        }
        public Clipper_Import_Fournitures_Divers_Processor()
        {

        }
        public void Import(IContext contextlocal)
        {
            //creation des logs
            ///TextWriterTraceListener logFile;
            try
            {
                //detection du contexte
                Cursor.Current = Cursors.WaitCursor;
                //creation des logs
                //creation du listener
                ////
                if (Import_Matiere)
                {
                    using (Clipper_Import_Matiere Update_Material = new Clipper_Import_Matiere(contextlocal))
                    {
                        try
                        {

                            //Update_Material.Almacam_Update_Material(contextlocal);
                            Update_Material.Import();
                            Update_Material.Dispose();
                            //Update_Material.Close();
                            Cursor.Current = Cursors.Default;

                        }
                        catch { Cursor.Current = Cursors.Default; }
                    }

                }

                if (Tube_Rond)
                {
                    using (Clipper_Import_Tube_Rond tubesronds = new Clipper_Import_Tube_Rond(contextlocal))
                    {
                        try
                        {

                            tubesronds.ReadTubes();
                            tubesronds.WriteTubes();
                            tubesronds.Close();
                            tubesronds.Dispose();


                        }
                        catch (Exception ie)
                        {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }
                }
                if (Rond)
                {
                    using (Clipper_Import_Rond ronds = new Clipper_Import_Rond(contextlocal))
                    {
                        try
                        {


                            ronds.ReadTubes();
                            ronds.WriteTubes();
                            ronds.Close();

                        }
                        catch (Exception ie)
                        {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }

                    }
                }
                if (Tube_Rectangle)
                {

                    using (Clipper_Import_Tube_Rectangle tubesrectangle = new Clipper_Import_Tube_Rectangle(contextlocal))
                    {
                        try
                        {
                            tubesrectangle.ReadTubes();
                            tubesrectangle.WriteTubes();
                            tubesrectangle.Close();

                        }
                        catch (Exception ie)
                        {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }


                    }
                }
                ///toles --> flat
                if (Tube_Flat)
                {
                    using (Clipper_Import_Tube_Flat tubesflats = new Clipper_Import_Tube_Flat(contextlocal))
                    {
                        try
                        {
                            tubesflats.ReadTubes();
                            tubesflats.WriteTubes();
                            tubesflats.Close();
                            tubesflats.Dispose();
                        }
                        catch (Exception ie) { MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error); }


                    }
                }
                if (Tube_Speciaux)
                {
                    using (Clipper_Import_Tube_Speciaux tubesspeciaux = new Clipper_Import_Tube_Speciaux(contextlocal))
                    {
                        try
                        {

                            tubesspeciaux.ReadTubes();
                            tubesspeciaux.WriteTubes();
                            tubesspeciaux.Close();
                            tubesspeciaux.Dispose();
                        }
                        catch (Exception ie)
                        {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }

                    }
                }
                if (Fourniture)
                {
                    using (Clipper_Import_Fournitures_Divers fournitures = new Clipper_Import_Fournitures_Divers(contextlocal))
                    {
                        try
                        {

                            fournitures.Read();
                            fournitures.Write();
                            fournitures.Close();
                            fournitures.Dispose();
                        }
                        catch (Exception ie)
                        {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }

                    }
                }


            }
            catch
            {

            }

        }

        /// <summary>
        /// main methode
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {

            try
            {
                if (OdbcTools.Test_Connexion())
                {
                    Import(Context);
                    AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT des fournitures ", "Import terminé.");
                }               
                
                return base.Execute();
            }
            catch (Exception ie)
            {
                AF_ImportTools.SimplifiedMethods.NotifyMessage("ERREUR IMPORT  des fournitures ", ie.Message);
                return false;
            }
            finally
            {

            }

        }




        

    }

    /// </summary>
    ///bouton import des tubes
    public class Clipper8_ImportTubes_Processor : CommandProcessor
    {



        /// <summary>
        /// constructeur de l'import des tube, tu stock et des matieres.........
        /// import_matiere,  tube_rond,  rond,  tube_rectangle,  tube_carre, tube_flat,  Tube_Speciaux
        /// </summary>
        /// <param name="import_matiere">true ou false pour declencher l'import</param>
        /// <param name="tube_rond">true ou false pour declencher l'import</param>
        /// <param name="rond">true ou false pour declencher l'import</param>
        /// <param name="tube_rectangle">true ou false pour declencher l'import</param>
        /// <param name="tube_carre">true ou false pour declencher l'import</param>
        /// <param name="tube_flat">true ou false pour declencher l'import</param>
        /// <param name="Tube_Speciaux">true ou false pour declencher l'import</param>
        /// <param name="Fourniture">true ou false pour declencher l'import</param>
        /// 
        private Boolean Import_Matiere = true;
        private Boolean Tube_Rond = true;
        private  Boolean Tube_Speciaux = true;
        private Boolean Rond = true;
        private Boolean Tube_Rectangle = true;
        private Boolean Tube_Carre = false;
        private Boolean Tube_Flat = true;
        private Boolean Fourniture = false;

        public Clipper8_ImportTubes_Processor(bool import_matiere, bool tube_rond, bool rond, bool tube_rectangle, bool tube_carre, bool tube_flat, bool tube_speciaux, bool fourniture)
        {

            Import_Matiere = import_matiere;
            Tube_Rond = tube_rond;
            Rond = rond;
            Tube_Rectangle = tube_rectangle;
            Tube_Carre = tube_carre;
            Tube_Flat = tube_flat;
            Tube_Speciaux = tube_speciaux;
            Fourniture = fourniture;


        }

        public Clipper8_ImportTubes_Processor()
        {
        }
        public void Import(IContext contextlocal)
        {
            //creation des logs
            ///TextWriterTraceListener logFile;
            try
            {
                //detection du contexte
                Cursor.Current = Cursors.WaitCursor;
                //creation des logs
                //creation du listener
                ////
                if (Import_Matiere)
                {
                    using (Clipper_Import_Matiere Update_Material = new Clipper_Import_Matiere(contextlocal))
                    {
                        try
                        {

                            //Update_Material.Almacam_Update_Material(contextlocal);
                            Update_Material.Import();
                            Update_Material.Dispose();
                            //Update_Material.Close();
                            Cursor.Current = Cursors.Default;

                        }
                        catch { Cursor.Current = Cursors.Default; }
                    }

                }

                if (Tube_Rond)
                {
                    using (Clipper_Import_Tube_Rond tubesronds = new Clipper_Import_Tube_Rond(contextlocal))
                    {
                        try
                        {

                            tubesronds.ReadTubes();
                            tubesronds.WriteTubes();
                            tubesronds.Close();
                            tubesronds.Dispose();


                        }
                        catch (Exception ie)
                        {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }
                }
                if (Rond)
                {
                    using (Clipper_Import_Rond ronds = new Clipper_Import_Rond(contextlocal))
                    {
                        try
                        {


                            ronds.ReadTubes();
                            ronds.WriteTubes();
                            ronds.Close();

                        }
                        catch (Exception ie)
                        {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }

                    }
                }
                if (Tube_Rectangle)
                {

                    using (Clipper_Import_Tube_Rectangle tubesrectangle = new Clipper_Import_Tube_Rectangle(contextlocal))
                    {
                        try
                        {
                            tubesrectangle.ReadTubes();
                            tubesrectangle.WriteTubes();
                            tubesrectangle.Close();

                        }
                        catch (Exception ie)
                        {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }


                    }
                }
                ///toles --> flat
                if (Tube_Flat)
                {
                    using (Clipper_Import_Tube_Flat tubesflats = new Clipper_Import_Tube_Flat(contextlocal))
                    {
                        try
                        {
                            tubesflats.ReadTubes();
                            tubesflats.WriteTubes();
                            tubesflats.Close();
                            tubesflats.Dispose();
                        }
                        catch (Exception ie) { MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error); }


                    }
                }
                if (Tube_Speciaux)
                {
                    using (Clipper_Import_Tube_Speciaux tubesspeciaux = new Clipper_Import_Tube_Speciaux(contextlocal))
                    {
                        try
                        {

                            tubesspeciaux.ReadTubes();
                            tubesspeciaux.WriteTubes();
                            tubesspeciaux.Close();
                            tubesspeciaux.Dispose();
                        }
                        catch (Exception ie)
                        {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }

                    }
                }
                if (Fourniture)
                {
                    using (Clipper_Import_Fournitures_Divers fournitures = new Clipper_Import_Fournitures_Divers(contextlocal))
                    {
                        try
                        {

                            fournitures.Read();
                            fournitures.Write();
                            fournitures.Close();
                            fournitures.Dispose();
                        }
                        catch (Exception ie)
                        {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }

                    }
                }


            }
            catch
            {

            }

        }
        /// <summary>
        /// main methode
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {

            try
            {
                if (OdbcTools.Test_Connexion())
                {
                    Import(Context);
                    AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT des tubes ", "Import terminé.");
                }

               
                return base.Execute();
            }
            catch(Exception ie)
            {
                AF_ImportTools.SimplifiedMethods.NotifyMessage("ERREUR IMPORT des tubes ",  ie.Message  );
                return false;
            }
            finally
            {
                
            }

        }

       
    }
    //////////////////////////////////////////////////////////////////////////////////////////
    /// <summary>
    /// 

    #region enum, class objets

    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <summary>
    /// routine d'import satndards, matiere, tubes.. toles, vis......
    /// </summary>
    //section des commandes
    public enum TypeTube
    {
        Aucun = 0,
        Rond = 4,
        Round = 5,
        Rectangle = 10,
        Flat = 3,
        Speciaux = 6
    }

    /// <summary>
    /// routine d'import satndards, matiere, tubes.. toles, vis......
    /// </summary>
    //section des commandes
    public enum SheetType
    {
        ToleNeuve = 0,
        Chute = 1,

    }

    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <summary>
    /// routine d'import satndards, matiere, tubes.. toles, vis......
    /// </summary>
    //section des commandes
    public enum UniteGestion
    {
        u = 1,
        le_dix = 10,
        le_cent = 100,

    }

    #endregion

    #region interface
    interface IArticle
    {
        /// <summary>
        /// verification de l'integrité des données articles
        /// </summary>
        /// <returns></returns>
        bool Check_Article_Integrity();
        /// <summary>
        /// lecture des donnees cam et clip
        /// </summary>
        void Read();

        /// <summary>
        /// ecriture dans cam et eventuelement maj
        /// </summary>
        void Write();
       
        /// <summary>
    }

    interface IArticleTube
    {
        /// <summary>
        /// verificaiton de l'integrité des données de tube
        /// </summary>
        /// <returns></returns>
        bool Check_Tube_Integrity(double longeurTube, IEntity quality);
        void ReadTubes();
        /// <summary>
        /// ecriture/mise a jour dans la base almacam des tubes 
        /// </summary>
        void WriteTubes();
        
    }



    #endregion

    #region classes statics 

    public static class OdbcTools
    {
        private static OdbcConnection connection_test = null;
        private static bool rst = false;
        //test de la connexion odbc paramétrée dans 
        public static bool Test_Connexion()
        {       ///premier paramertre dsn = clipper dsn= data source name
            //var DbConnection =new OdbcConnection("DSN=" + DSN);


            try
            {
                OdbcConnection connection_test; //= new OdbcConnection("DSN=" + DSN);

                var jstools = new JsonTools();
                //recup parametre dsn //
                /************/
                //le dsn est le nom de la connextion odbc paramétre par l'installateur clipper
                //cette info est paramétrable dans le fichier json
                AF_ImportTools.SimplifiedMethods.NotifyMessage("OBDC Test", "Json connexion...");
                var DSN = jstools.getJsonStringParametres("dsn");
                /************/


                if (AlmaCamTool.Is_Odbc_Exists(DSN))

                {
                    connection_test = new OdbcConnection("DSN=" + DSN);
                    connection_test.Open();
                    //System.Windows.Forms.MessageBox.Show("Connexion ok : " + DSN);
                    rst = true;
                }
                else { throw new Missing_Obdc_Exception(DSN); }

                return rst;

            }

            catch (Exception ie)
            {
                System.Windows.Forms.MessageBox.Show("ERREUR : " + ie.Message.ToString());
                return rst;
            }
            finally
            {
                if (connection_test != null)
                {
                    connection_test.Close();
                    connection_test = null;

                }

                //return rst;

            }


        }


    }


    #endregion
    //////////////////////////////////////////////////////////////////////////////////////////
    ///json   
    #region json
    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <summary>
    /// gestion des fichier json (partira dans import tools plus tard
    /// </summary>
    //section JSON
    public class JsonTools
    {
        //recupee le fichier json dans le dossier de la dll
        private string PATH = Directory.GetCurrentDirectory() + "\\" + Properties.Resources.jsonimportclipper;//"import-clipper.json";

        public string getJsonStringParametres(string ParameterName)
        {

            try
            {
                string returnvalue = "";
                using (StreamReader file = File.OpenText(this.PATH))
                {


                    using (JsonTextReader reader = new JsonTextReader(file))
                    {
                        JObject o = (JObject)JToken.ReadFrom(reader);
                        returnvalue = (string)o.SelectToken(ParameterName);
                    }


                }

                return returnvalue;
            }


            catch (Exception ie)
            {

                MessageBox.Show(ie.Message, "JsonTools Class", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;


            }

        }


    }
    #endregion

    //////////////////////////////////////////////////////////////////////////////////////////
    /// <summary>
    ///   class Clipper_Material.. class de stockage des matieres
    ///   les codes articles mutlidim sont recuperés dans le champs code article matiere
    /// </summary>
    public class Clipper_Material : Clipper_Article, IDisposable
    {
        //initialisation des parametres
        //definiton de l'objetc material
        //public string Coarti = "UNDEF"; //code etat (nuance epaisseur)
        //public double Prixart=0;// prixart
        //public double Densite=1;//
        //public string Nuance= "UNDEF";
        //public string Etat = "";//code nuance
        //public int Type = 0;
        //public bool IsMultiDim = false;
        //public string Quality; //quality
        //public double Thickness = 0;  //DIM1
        //public string Materialname = "UNDEF";
        public string Comments = "";
        public string Uidkey;

        /// <summary>
        /// retourne un nom normalisé pour la nuance Nuance*Etat;
        /// </summary>
        /// <returns></returns>
        public string GetNuance() { return Nuance + "*" + Etat; }
        /// <summary>
        /// reourne l'epaissseur matiere
        /// </summary>
        /// <returns></returns>
        public double GetThickess() { return Thickness; }

        //construit le nom matiere en concatenant nuance*etat
        /// <summary>
        /// update material name
        /// </summary>
        public void SetMaterialName()
        {   //setuid de l'objet
            Uidkey = GetMaterial_Uidkey(GetQualityName(Nuance, Etat), Thickness, Type);
            if (Etat != string.Empty)
            {
                Quality = Nuance + "*" + Etat;
                Materialname = Nuance + "*" + Etat + " " + Thickness + " mm";
            }
            else
            {
                Quality = Nuance;
                Materialname = Nuance + " " + Thickness + " mm";
            }

            //cle unique pour eviter les doublons
            //Uidkey = Nuance.Trim() + "*" + Etat.Trim() + " " + Thickness;


            // return Materialname ;
        }
        /// <summary>
        /// creer l'uidmatiere interne a la dll (structutrée) quality*thickness
        /// </summary>
        /// <param name="qualityname">nuance*etant generalement</param>
        /// <param name="Thickness">epaisseure</param>
        /// <returns></returns>
        public static string GetMaterial_Uidkey(string qualityname, double Thickness, int type)
        {
            string uidkey;

            uidkey = qualityname + "*" + Thickness.ToString() + "_" + type;

            return uidkey;

            // return Materialname ;
        }
        // <summary>
        /// recupere le nomde la qualité Nuance*Etat*epaisseur
        /// </summary>
        /// <returns></returns>
        public string GetQualityName(string Nuance, string Etat)
        {

            string qualityname;
            //cle unique pour eviter les doublons non stockée
            if (Etat != string.Empty)
            { qualityname = Nuance.Trim() + "*" + Etat.Trim(); }
            else
            { qualityname = Nuance.Trim(); }

            return qualityname;
        }
        /// <summary>
        /// recupere le nom matiere Nuance*Etat*epaisseur
        /// </summary>
        /// <returns></returns>
        public string GetMaterialName() { return Materialname; }

        /*
        /// <summary>
        /// dispose
        /// </summary>
        public void Dispose()
        {
            //Dispose(true);
            GC.SuppressFinalize(this);
        }*/

    }
    /// <summary>
    ///   class Clipper_IMport matiere.. derive de clipper article
    /// </summary>
    public class Clipper_Import_Matiere : IDisposable
    {
        private string DSN;
        private string SQL;
        private string SQL_NUANCE_ETAT;
        private JsonTools JSTOOLS;
        private OdbcDataReader TABLE_ARTICLEM_TOLE;
        private OdbcConnection DbConnection;
        private OdbcCommand DbCommand;
        private IContext ContextLocal;
        private bool _FAVORISE_MULTIDIM = true; // en prevision sp6
        IList<Clipper_Material> CLIPPER_MATERIAL_LIST = new List<Clipper_Material>(); //liste des matieres clipper
        //IList<Flat> CLIPPER_TOLE_LIST = new List<Flat>();    //liste des toles clipper
        IList<Clipper_Material> CLIPPER_TOLE_LIST = new List<Clipper_Material>();    //liste des toles clipper
        //recuperation des parametres
        //creation du listener
        TextWriterTraceListener logFile = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + Properties.Resources.ImportMatiereLog);
        public IList<Clipper_Material> GetMaterial_List() { return CLIPPER_MATERIAL_LIST; }

        public void Dispose()
        {
            //Dispose(true);
            GC.SuppressFinalize(this);
        }
        //constructeur
        #region constructeur
        public Clipper_Import_Matiere(IContext contextlocal)
        {
            this.ContextLocal = contextlocal;
            logFile.Write("debut import matiere ");
            Init();
        }


        #endregion
        /// <summary>
        /// init est une methode qui se connect connect  en odbc et 
        /// lance une commande odbc su le dsn indiqué dans le fichier json
        /// </summary>
        private void Init()

        {    ///premier paramertre dsn = clipper dsn= data source name
              
            try
            {
                if (OdbcTools.Test_Connexion())
                {
                    AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT MATIERE", "import des matieres en court...");
                    logFile.WriteLine("initialisatin de la methode");
                    //lecture du fichier json
                    logFile.WriteLine("lecture du fichier Json");
                    // Material matiere = new material(); //
                    JSTOOLS = new JsonTools();
                    //recup parametre dsn //
                    /************/
                    //le dsn est le nom de la connextion odbc paramétre par l'installateur clipper
                    //cette info est paramétrable dans le fichier json
                    DSN = JSTOOLS.getJsonStringParametres("dsn");
                    /************/
                    logFile.WriteLine("DSN demandé" + DSN);

                    if (AlmaCamTool.Is_Odbc_Exists(DSN))

                    {
                        DbConnection = new OdbcConnection("DSN=" + DSN);
                        DbConnection.Open();
                        DbCommand = DbConnection.CreateCommand();
                        logFile.WriteLine("etat de la connexion " + DbConnection.State.ToString());
                        Get_Clipper_Material(ContextLocal);
                    }
                    else { throw new Missing_Obdc_Exception(DSN); }
                }
            }

            catch (Exception ie)
            {
                logFile.WriteLine("ERREUR : " + ie.Message.ToString());
            }
            finally
            {
                //DbConnection.Close();
                this.Dispose();
            }


        }
       
        /// <summary>
        ///  recupere dans une table CLIPPER_MATERIAL_LIST
        ///  la requete est sql.nuance_etat
        /// </summary>
        private void Get_Clipper_Material(IContext contextlocal)
        {

           

            Clipper_Param.GetlistParam(contextlocal);

            //IList <Material> materialist = new List<Material>();
            //string prefix; // prefixe article matiere
            //string grade;//code nuance
            //string name; //code etat (nuance epaisseur)
            //double price; // prix
            //double densité;
            //double Thickness;
            int ii = 0;
            logFile.WriteLine("paramétrage de la requete du fichier json ");
            logFile.WriteLine("monodim = " + JSTOOLS.getJsonStringParametres("multidim"));
            //recuperation des nuance etat--> 304L*DKP
            //on recherche toutes les matieres mono et multidim


            this.SQL_NUANCE_ETAT = JSTOOLS.getJsonStringParametres("sql.nuance_etat");
            //coupel nuance etat
            //Material matiere = new material();
            //getJsonParametres();
            this.DbCommand.CommandText = this.SQL_NUANCE_ETAT;
            logFile.WriteLine("requete utilisée = " + this.SQL_NUANCE_ETAT);
            // requete type matiere recuperer dans la section nuance_etat //
            TABLE_ARTICLEM_TOLE = DbCommand.ExecuteReader();

            while (TABLE_ARTICLEM_TOLE.Read())
            {
                ii++;
                //Clipper_Material material = new Clipper_Material();
                var material = new Clipper_Material();
                //Clipper_Article current_article = new Clipper_Article();
                material.Nuance = TABLE_ARTICLEM_TOLE["CODENUANCE"].ToString().Trim();
                material.Etat = TABLE_ARTICLEM_TOLE["CODEETAT"].ToString().Trim();
                material.PRIXART = Convert.ToDouble(TABLE_ARTICLEM_TOLE["PRIXART"]);
                material.Densite = Convert.ToDouble(TABLE_ARTICLEM_TOLE["DENSITE"]);
                material.Type = Convert.ToInt32(TABLE_ARTICLEM_TOLE["TYPE"]);  //type article 3,4,5
                material.IsMultiDim = Convert.ToBoolean(TABLE_ARTICLEM_TOLE["MULTIDIM"]);


                //attention pour les tole monodim l'epaisseur est stock en dim3.
                //recuperation des codes multidim par defaut

                if (material.IsMultiDim)
                {
                    material.Thickness = Convert.ToDouble(TABLE_ARTICLEM_TOLE["EPAISSEUR"]);
                    //uniquement code mutli dim
                    //material.COARTI = TABLE_ARTICLEM_TOLE["COARTI"].ToString().Trim();
                    //material.PRIXART = 0;

                }

                else
                {

                    material.Thickness = Convert.ToDouble(TABLE_ARTICLEM_TOLE["DIM3"]);
                    material.COARTI = string.Empty;
                }

                material.COARTI = TABLE_ARTICLEM_TOLE["COARTI"].ToString().Trim();
                //on voit dans la liste le code article //
                material.Comments = TABLE_ARTICLEM_TOLE["MATERIAL_COMMENTS"].ToString().Trim();
                //issue du tube ou non//
                //le controle des matieres est obligatoire
                if (TABLE_ARTICLEM_TOLE["TYPE"].ToString().Trim() == "3" && CheckDataintegrity(material) == true)

                {   //plats
                    material.Comments = TABLE_ARTICLEM_TOLE["MATERIAL_COMMENTS"].ToString().Trim();
                    //creation de la liste  des articiel tole mono et multidims
                    //on ajoute toutes les toles clip en dso differentes de nulless
                    //current_article
                    CLIPPER_TOLE_LIST.Add(material);

                }

                else
                {
                    //tube et autres
                    material.Comments = "Matiere_Tube " + TABLE_ARTICLEM_TOLE["MATERIAL_COMMENTS"].ToString().Trim();

                }
                material.SetMaterialName();
                //teste d'integrite sur la matiere
                //if (CheckDataintegrity(material))

                //{
                // on controle les doublons ici //
                if (CLIPPER_MATERIAL_LIST.Where(m => m.Uidkey == material.Uidkey).Count() == 0)
                { CLIPPER_MATERIAL_LIST.Add(material); }
                else
                {
                    logFile.WriteLine("doublon detecter sur la matiere, la matiere ne sera pas ajouté " + material.Uidkey);
                }

                



                //  }

            };
            logFile.WriteLine("reconstitution de la table des matieres terminés : " + this.CLIPPER_MATERIAL_LIST.Count().ToString() + " trouvées ");


        }
        /// <summary>
        /// fermeture de la connexion odbc
        /// </summary>

        public void Close()
        {

            // reader.Close();

            DSN = "";
            SQL = "";

            if (TABLE_ARTICLEM_TOLE != null)
            {
                TABLE_ARTICLEM_TOLE.Close();

            }

            DbConnection.Dispose();
            DbCommand.Dispose();
            CLIPPER_MATERIAL_LIST = null;
            DbCommand.Dispose();
            DbConnection.Close();
            //destroy list

            CLIPPER_TOLE_LIST = null;
            CLIPPER_MATERIAL_LIST = null;

            //
            logFile.Close();

        }

        /// <summary>
        /// phase d'import
        /// ecrit les nouvelles matieres dans almacam
        /// </summary>
        /// <param name="contextlocal"></param>

        public void Import()
        {


            try
            {
                //  logFile.WriteLine("mise a jour des matieres dans la base " + contextlocal.Model.DatabaseName);
                //ecrirture de la liste des matiere//
                //recuperation de la liste des matieres almacam
                IEntityList qualityentitylist = ContextLocal.EntityManager.GetEntityList("_QUALITY");
                IEntityList materialentitylist = ContextLocal.EntityManager.GetEntityList("_MATERIAL");
                IList<string> qualities_To_Create = new List<string>();
                qualityentitylist.Fill(false);
                IList<IEntity> qualitylist = new List<IEntity>();
                //select distincte supprime les doublons
                qualitylist = qualityentitylist.Distinct().ToList();
                bool newdatabase = qualitylist.Count == 0;
                ////liste des qualité de materiaux
                IDictionary<long, string> updatedstringqualitylist = new Dictionary<long, string>();

                //////
                ///////detection des qualités a creer//// 
                logFile.WriteLine("creation de la liste des qualités de clipper ");
                foreach (Clipper_Material m in CLIPPER_MATERIAL_LIST)
                {
                    if (qualities_To_Create.Contains(m.Quality) == false)
                    { qualities_To_Create.Add(m.Quality); }
                }
                logFile.WriteLine(qualities_To_Create.Count().ToString() + " qualités trouvées ");

                ///////////
                ///creation des qualités dans la base si necessaire
                //////////            
                foreach (string quality in qualities_To_Create)
                {

                    // Find material           
                    IEntity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();

                    if (currentQuality == null)
                    {
                        //creation de la nuance et sauvegarde// ou retour de la qualité courente
                        currentQuality = ContextLocal.EntityManager.CreateEntity("_QUALITY");
                        currentQuality.SetFieldValue("_NAME", quality);
                        logFile.WriteLine("creation de la nouvelle qualité  " + quality);
                        currentQuality.Save();

                    }




                }

                logFile.WriteLine(" qualités crées ");
                qualityentitylist.Fill(false);
                qualitylist = qualityentitylist.ToList();

                ///////////
                ///creation des matieres assocées au qualités
                //mise a jour de la liste des matiere
                // on ne prends que les matiere de type 3  (tole/plat)         

                foreach (Clipper_Material m in CLIPPER_MATERIAL_LIST)
                {

                    // Find material           
                    IEntity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(m.Quality)).FirstOrDefault();
                    //update
                    currentQuality.SetFieldValue("_NAME", m.Quality);
                    currentQuality.SetFieldValue("_DENSITY", (m.Densite) / 1000);
                    //calcul du prix moyen toutes epaisseurs confondues
                    //currentQuality.SetFieldValue("_BUY_COST", (CLIPPER_MATERIAL_LIST.Average(p => p.PRIXART)));
                    //currentQuality.SetFieldValue("_OFFCUT_COST", (CLIPPER_MATERIAL_LIST.Average(p => p.PRIXART) / 1000));
                    currentQuality.SetFieldValue("_COMMENTS", m.Comments);
                    currentQuality.Save();
                    //on rempli la liste des qualités

                    if (currentQuality != null)
                    {
                        if (!updatedstringqualitylist.ContainsKey(currentQuality.Id))
                        {
                            updatedstringqualitylist.Add(currentQuality.Id, m.Quality);
                        }
                    }

                }


                //création des matieres   ---> pas d'unicité (voir dans la requete)
                //on rempli une liste de keypair de type uidkey,ientity matiere pour detecter les doublons uidkey est une cle interne a l'obkey clipper matierial
                //
                logFile.WriteLine("creation des matieres assocées aux qualités  ");
                foreach (var currentstringquality in updatedstringqualitylist)
                {
                    {   //recuperation de la lisre de matieres associée a la qualité choisie

                        IEntityList almacam_materialentitylist = ContextLocal.EntityManager.GetEntityList("_MATERIAL", "_QUALITY", ConditionOperator.Equal, currentstringquality.Key);
                        almacam_materialentitylist.Fill(false);

                        //construction de la liste des keypair (voir pour mettre un distionnaire si trop lent)
                        IList<IEntity> almacam_material_list = new List<IEntity>();
                        almacam_material_list = almacam_materialentitylist.ToList();
                        //detection des doublons
                        List<KeyValuePair<string, IEntity>> almacam_material_UiKey = new List<KeyValuePair<string, IEntity>>();
                        foreach (Entity e in almacam_material_list)
                        {
                            string quality = e.GetFieldValueAsEntity("_QUALITY").GetFieldValueAsString("_NAME");
                            double thickness = e.GetFieldValueAsDouble("_THICKNESS");
                            //Clipper_Material.GetMaterial_Uidkey(quality, thickness);
                            almacam_material_UiKey.Add(new KeyValuePair<string, IEntity>(quality + "*" + thickness, e));
                        }

                        //on se limite aux matiere des plats/toles
                        foreach (Clipper_Material m in this.CLIPPER_MATERIAL_LIST.Where(t => t.Type == 3))
                        {
                            if (m.Quality == currentstringquality.Value && m.Type == 3)
                            {

                                IEntity currentmaterial = null;

                                //IEntity currentmaterial = almacam_material_list.Where(q => q.GetFieldValueAsString("_NAME").Equals(m.GetMaterialName())).FirstOrDefault();
                                //si la clé est retrouvé alors c'est un update sinon c'est une creation
                                if (almacam_material_UiKey.Where(kvp => kvp.Key == m.Uidkey.Split('_')[0]).Count() > 0)
                                {
                                    currentmaterial = almacam_material_UiKey.First(kvp => kvp.Key == m.Uidkey.Split('_')[0]).Value;
                                }


                                //creation
                                if (currentmaterial == null)
                                {
                                    //creation de la matiere et sauvegarde//
                                    currentmaterial = ContextLocal.EntityManager.CreateEntity("_MATERIAL");
                                    logFile.WriteLine("creation de la matieres   " + m.GetMaterialName());
                                    currentmaterial.Save();

                                }

                                //update and save
                                logFile.WriteLine("mise à jour de la matiere :   " + m.GetMaterialName());

                                currentmaterial.SetFieldValue("_NAME", m.GetMaterialName());
                                currentmaterial.SetFieldValue("_QUALITY", currentstringquality.Key);
                                currentmaterial.SetFieldValue("_THICKNESS", m.Thickness);
                                currentmaterial.SetFieldValue("_BUY_COST", (m.PRIXART) / 1000);
                                //on rempli le code sur les multidim uniquement
                                if (m.IsMultiDim) { 
                                currentmaterial.SetFieldValue("_CLIPPER_CODE_ARTICLE", m.COARTI);
                                }
                                

                            IEntityList currentstocks = ContextLocal.EntityManager.GetEntityList("_STOCK", "_NAME", ConditionOperator.Equal, m.COARTI);
                                currentstocks.Fill(false);


                                if (currentstocks.Count() > 0)
                                {
                                    IEntity currentSheet = currentstocks.FirstOrDefault().GetFieldValueAsEntity("_SHEET");
                                    currentmaterial.SetFieldValue("AF_DEFAULT_SHEET", currentSheet.GetFieldValueAsString("_REFERENCE"));
                                    currentSheet = null;
                                }

                                currentmaterial.SetFieldValue("_COMMENTS", m.Comments);
                                currentmaterial.Save();

                                //on set le nom standard
                                CommonModelBuilder.ComputeMaterialName(currentmaterial.Context, currentmaterial.GetFieldValueAsEntity("_QUALITY"), currentmaterial);
                                currentmaterial.Save();
                            }
                        }







                        //pas utils mais au cas ou
                        almacam_materialentitylist = null;
                        almacam_material_list = null;


                    }
                }




                ///Update material price
                ///mise a jour des prix matiere sur les toles multi et monodim//
                foreach (var t in CLIPPER_TOLE_LIST)
                {


                    //pour le momelnt on ne trie pas sur les toles deja traitées.List<long> TreatedSheetList = new List<long>();
                    // Find material           
                    IEntityList currentstocks = ContextLocal.EntityManager.GetEntityList("_STOCK", "_NAME", ConditionOperator.Equal, t.COARTI);
                    currentstocks.Fill(false);

                    if (currentstocks.Count() > 0)
                    {
                        //sheet

                        //IEntity currentstock = currentstocks.FirstOrDefault();
                        foreach (var currentstock in currentstocks)
                        {


                            IEntity currentSheet = currentstock.GetFieldValueAsEntity("_SHEET");
                            currentSheet.SetFieldValue("_BUY_COST", t.PRIXART);
                            currentSheet.SetFieldValue("_AS_SPECIFIC_COST", !t.IsMultiDim);
                            currentSheet.Save();

                            //stock
                            currentstock.SetFieldValue("AF_IS_MULTIDIM", t.IsMultiDim);

                            currentstock.Save();
                            currentSheet = null;

                        }

                    }






                }




                //*//

                logFile.Flush();

                this.Close();
            }
            catch (Exception ie)
            {
                logFile.WriteLine("Erreur import matiere non definie : "+ ie.Message);
                //MessageBox.Show(ie.Message);
            }
            finally
            {

                logFile.WriteLine("Import ended ");
            }

        }
        /// <summary>
        /// verificaiton des matieres
        /// </summary>
        /// <param name="material"></param>
        /// <returns></returns>
        public bool CheckDataintegrity(Clipper_Material material)
        {
            bool integrity = true;

            return integrity;
        }

        /// <summary>
        /// verification si la matier existe dans la base cam
        /// </summary>
        /// <param name="material"></param>
        /// <returns></returns>
        public bool Material_Exists(Clipper_Material material)
        {
            bool integrity = true;

            return integrity;
        }



    }
   

   

    /// //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    /// </summary>
    /// concerne l'import des articles
    /// 
    /// <summary>
    ///   classe de base clipper article
    /// </summary>
    public class Clipper_Article :IArticle
    {
        //string prefix; // prefixe article matiere
        //creation du listener

        public IContext contextlocal;
        public long AMCLEUNIK;
        public string InternalName; // nom interne
        //fournisseur

        public string FOURN;
        public string UniteGest = "u"; //unite de gestion : u ou le cent , le dix...
        public string UnitePrix = "u";
        public string COARTI = "UNDEF"; //code etat (nuance epaisseur)
        public double PRIXART;// prixart
        public int Type = 0;  // type d'article --> famille 

        /*-- famille.DIMENSIONS=3 tole/plat
                    -- famille.DIMENSIONS=17 tube rectangulaire
                    -- famille.DIMENSIONS=16 tube carre
                    -- famille.DIMENSIONS=4 rond
                    -- famille.DIMENSIONS=5 tube rond
        */

        public double Densite;//
        public IEntity Material;
        public string Quality; //quality
        public string Nuance;
        public string Etat = "";
        public double Thickness = 0;  //DIM1
        public string Materialname = "UNDEF";
        public bool IsMultiDim = false;

        //codification
        public string Name = "UNDEF";
        public string COFA = "UNDEF";
        public string DESA1 = "";
        public string DESA2 = "";
        private string DSN;
        //dimensions
        //public Dictionary<string, double> DIM;
        /// <summary>
        /// conexion
        /// </summary>
        public OdbcDataReader TABLE_ARTICLEM;//TABLE_ARTICLEM;
        public OdbcConnection DbConnection;
        public OdbcCommand DbCommand;
        public JsonTools JSTOOLS;

        public void Dispose()
        {
            TABLE_ARTICLEM = null;
            //TABLE_ARTICLEM_ROND = null;
            DSN = null;
        }

        IList<Clipper_Material> TUBE_LIST = new List<Clipper_Material>();

        /// <summary>
        /// constructeur
        /// </summary>
        public Clipper_Article()
        {
        }
        /// <summary>
        /// verification de l'integrité des données
        /// </summary>
        /// <returns></returns>
        public virtual bool Check_Article_Integrity()
        {
            return true;
        }
        /// <summary>
        /// lecture des donnees cam et clip
        /// </summary>
        public virtual void Read()
        {


        }
        /// <summary>
        /// ecriture dans cam et eventuelement maj
        /// </summary>
        public virtual void Write()
        {


        }
        /// <summary>
        /// maj dans cam 
        /// </summary>
        public virtual void Update()
        {


        }

        /// <summary>
        /// creation de la connexion odbc
        /// </summary>
        public void Odbc_Connexion()
        {    ///premier paramertre dsn = clipper dsn= data source name
            try
            {

                //Material matiere = new material();
                JSTOOLS = new JsonTools();
                //recup parametre dsn
                DSN = JSTOOLS.getJsonStringParametres("dsn");

                if (AlmaCamTool.Is_Odbc_Exists(DSN))
                {

                    //recupe requete
                    DbConnection = null;
                    DbCommand = null;
                    this.DbConnection = new OdbcConnection("DSN=" + DSN);
                    this.DbConnection.Open();
                    this.DbCommand = DbConnection.CreateCommand();
                  
                }

            }
            catch (Exception ex)
            {
                string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
                MessageBox.Show(methode, ex.Message);
                Environment.Exit(0);
            }
        }
        //permet d'ecrire dans le log almacam
        public void Odbc_Connexion(IContext contextlocal)
        {    ///premier paramertre dsn = clipper dsn= data source name
            try
            {

                //Material matiere = new material();
                JSTOOLS = new JsonTools();
                //recup parametre dsn
                DSN = JSTOOLS.getJsonStringParametres("dsn");

                if (AlmaCamTool.Is_Odbc_Exists(DSN))
                {

                    //recupe requete
                    DbConnection = null;
                    DbCommand = null;
                    this.DbConnection = new OdbcConnection("DSN=" + DSN);
                    this.DbConnection.Open();
                    this.DbCommand = DbConnection.CreateCommand();
                    contextlocal.TraceLogger.TraceInformation("Creation du lien ODBC..");
                    //logFile.WriteLine("connexion ok    : " + DSN);
                }

            }
            catch (Exception ex)
            {
                string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
                MessageBox.Show(methode, ex.Message);
                Environment.Exit(0);
            }
        }
        /// <summary>
        /// creation de la connexion ole
        /// </summary>
        public void Ole_Connexion()
        {    ///premier paramertre dsn = clipper dsn= data source name
            try
            {




            }
            catch (Exception ex)
            {
                string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
                MessageBox.Show(methode, ex.Message);
                Environment.Exit(0);
            }
        }
        /// <summary>
        /// recuperation ou conversion numeric de donnees sql, retourn 0 si texte vide
        /// </summary>
        /// <param name="fieldname"></param>
        /// <returns></returns>
        public string GetSqlNumericValue(string fieldname)
        {
            try
            {
                //string value;
                if (TABLE_ARTICLEM[fieldname].ToString().Trim() == string.Empty) { return "0"; }
                else { return TABLE_ARTICLEM[fieldname].ToString().Trim(); }

            }
            // catch { return "0"; }
            catch (Exception ex)
            {
                string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
                MessageBox.Show(methode, ex.Message + fieldname + "Not found ");
                return "0";

            }

        }

        /// <summary>
        /// recuperation ou conversion numeric de donnees sql, retourn 0 si texte vide
        /// </summary>
        /// <param name="fieldname"></param>
        /// <returns></returns>
        public int GetSqlIntValue(string fieldname)
        {
            try
            {
                //string value;
                if (TABLE_ARTICLEM[fieldname].ToString().Trim() == string.Empty) { return 0; }
                else
                {
                    string temp = TABLE_ARTICLEM[fieldname].ToString().Trim();
                    return Convert.ToInt32(temp);

                }

            }
            // catch { return "0"; }
            catch (Exception ex)
            {
                string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
                MessageBox.Show(methode, ex.Message);
                return 0;

            }

        }

        public List<string> CheckDoublonsInStringList(List<string> Listacontroler)
        {
            List<String> temps = new List<String>();
            List<String> doublons = new List<String>();
            try
            {


                foreach (string s in Listacontroler)
                {
                    if (!temps.Contains(s)) { temps.Add(s); }
                    else { doublons.Add(s); }



                }
                temps.Clear();
                return doublons;
            }
            catch { return null; }


        }
        public List<string> CheckDoublonsInkeypPairList(List<KeyValuePair<string, long>> Listkeypairacontroler)
        {
            List<String> temps = new List<String>();
            List<String> doublons = new List<String>();

            try
            {


                foreach (var v in Listkeypairacontroler)
                {
                    if (!temps.Contains(v.Key))
                    { temps.Add(v.Key); }
                    else
                    {
                        doublons.Add(v.Key);
                    }
                }
                temps.Clear();
                return doublons;
            }
            catch { return null; }


        }
        public void Close()
        {

            DSN = "";
            //SQL = "";
            TABLE_ARTICLEM.Close();
            DbConnection.Dispose();
            DbCommand.Dispose();
            DbCommand.Dispose();
            DbConnection.Close();

            //logFile.WriteLine("connexion    : Fin" );
            //logFile.Close();
            //logFile.Flush();
            //logFile.Dispose();


        }

        #region Matiere
        public IEntity GetGrade(IContext contextlocal, string nuance, string etat)
        {
            try
            {
                //en cas de lenteru on pourrais stocker toutes les gardes poour eviter de faire troo de requetes
                //nous verrons
                IEntityList grades;
                IEntity grade;
                string qualityname;
                if (etat == "") { qualityname = nuance; } else { qualityname = nuance + "*" + etat; }

                grades = contextlocal.EntityManager.GetEntityList("_QUALITY", "_NAME", ConditionOperator.Equal, qualityname);
                grades.Fill(false);

                if (grades.Count() > 0)
                {
                    grade = grades.FirstOrDefault();


                }
                else
                {
                    grade = null;

                }


                return grade;


            }
            catch (Exception ie) { MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error); return null; }
        }
        /// <summary>
        /// converti le prix en fonction de l'unité de gestion
        /// </summary>
        /// <param name="unitegestion">par longeur ou par kg ou par u ou par le 100...</param>
        /// <param name="price"></param>
        /// <returns></returns>
        public double GetPrice(string unitegestion, double price,double longeur,double poids)
        {
            try
            {

                switch (unitegestion)
                {
                    case "le cent":
                        price = price / 100;
                        break;

                    case "le dix":
                        price = price / 10;
                        break;
                    default:
                        price = price/1;
                        break;
                }


                return Math.Round(price,5);


            }
            catch { return 0; }
        }

    }

    #endregion

    #region class des fourniture (vis, ecrous.... cartons)

    /// <summary>
    /// propriété specifiques founitures divers
    /// </summary>
    public class Founiture_Divers : Clipper_Article, IDisposable
    {

    }

    #endregion


    #region classe des toles (vis, ecrous.... cartons)
    /// <summary>
    /// recuperation des toles monodims
    /// </summary>
    public class Clipper_Article_Toles_MonoDim : Clipper_Article, IDisposable
    {

        //private string entityType = "_SHEET";
        public List<Flat> Clipper_articles_Stock;
        //public Dictionary<string, IEntity> Almacam_Stock;
        public Dictionary<string, IEntity> Almacam_Stock;
        /* public JsonTools JSTOOLS; */
        //IList<Clipper_Material> TOLES_MONO_DIM_LIST = new List<Clipper_Material>();

        public Clipper_Article_Toles_MonoDim(IContext context)
        {
            this.contextlocal = context;
            Odbc_Connexion();
            //getAlmacam_Existing_Format_List(entityType);
            Almacam_Stock = getAlmacam_Stock_Format_List();
        }


        TextWriterTraceListener logFilesheetMonodim = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "SHEET_MONO_" + Properties.Resources.Import_Sheet_Mono);



        /// <summary>
        /// retourn une liste d'exclusion des formats
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="entity_uniquestring_field"></param>
        /// <returns></returns>
        public List<string> getAlmacam_Existing_Format_List(string entitype)
        {
            try
            {

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                exclusionlist.Fill(false);
                List<string> exclusion = new List<string>();


                foreach (IEntity ex in exclusionlist)
                {
                    string uid = ex.GetFieldValueAsString("_REFERENCE").Trim();
                    if (uid != string.Empty)
                    {
                        if (exclusion.Contains(uid) == false)
                        { exclusion.Add(uid); }
                    }


                }

                return exclusion;
            }

            catch
            {

                return null;
            }


        }

        /// <summary>
        /// retourne la liste des toles et leur format
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="entity_uniquestring_field"></param>
        /// <returns>dicionnaire string entity sheet </returns>
        public Dictionary<string, IEntity> getAlmacam_Stock_Format_List()
        {
            try
            {
                string entitype = "_STOCK";
                Dictionary<string, IEntity> sheetlist = new Dictionary<string, IEntity>();

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                exclusionlist.Fill(false);
                //List<string> exclusion = new List<string>();

                //recuperation du format de chaque entity de stock
                foreach (IEntity ex in exclusionlist)
                {
                    string uid = ex.GetFieldValueAsString("_NAME").Trim();
                    IEntity e = ex.GetFieldInternalValueAsEntity("_SHEET");
                    uid += "_" + e.Id;
                    if (uid != string.Empty)
                    {
                        IEntity rst = null;
                        sheetlist.TryGetValue(uid, out rst);
                        if (rst == null)
                        { sheetlist.Add(uid, e); }
                    }


                }



                // Create the query.
                // The first line could also be written as "var studentQuery ="





                return sheetlist;





            }

            catch
            {

                return null;
            }


        }


        /// <summary>
        /// rempli la table Clipper_articles_Stock
        /// pour le moment seules les toles neuves et format commerciaux sont recupere dans cette list
        /// object tole article
        ///flat.COARTI 
        ///flat.Nuance
        ///flat.Etat
        ///flat.PRIXART
        ///bool multidim
        ///flat.MultiDim
        //recuperation des dimenssions
        ///flat.Longueur
        ///flat.Largeur
        ///flat.Epaisseur
        ///flat.Densite
        ///flat.COFA
        ///flat.InternalName = nuance_etat + "*" + flat.Longueur + "*" + flat.Largeur + "*" + flat.Epaisseur;
        ///flat.type= SheetType.Chute;
        ///flat.type = SheetType.ToleNeuve; //forcement tole neuve
        ///flat.commercial = true;
        ///
        /// </summary>
        public override void Read()
        {


            Clipper_articles_Stock = new List<Flat>();


            try
            {

                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.All_Flat");
                //logFileTubeRond.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                int ii = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    //Flat flat = new Flat();
                    var flat = new Flat();
                    flat.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();

                    flat.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    flat.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    string nuance_etat = flat.Etat == "" ? flat.Nuance : flat.Nuance + "*" + flat.Etat;

                    flat.PRIXART = Convert.ToDouble(GetSqlNumericValue("PRIXART"));
                    bool multidim = TABLE_ARTICLEM["MULTIDIM"].ToString() == "0" ? false : true;
                    flat.IsMultiDim = multidim;
                    //recuperation des dimenssions
                    if (flat.IsMultiDim)
                    {
                        flat.Longueur = Convert.ToDouble(GetSqlNumericValue("LOMULTIDIM"));
                        flat.Largeur = Convert.ToDouble(GetSqlNumericValue("LAMULTIDIM"));
                        flat.Epaisseur = Convert.ToDouble(GetSqlNumericValue("EPMULTIDIM"));
                    }
                    else
                    {
                        flat.Longueur = Convert.ToDouble(GetSqlNumericValue("LOMONODIM"));
                        flat.Largeur = Convert.ToDouble(GetSqlNumericValue("LAMONODIM"));
                        flat.Epaisseur = Convert.ToDouble(GetSqlNumericValue("EPMONODIM"));
                    }

                    flat.Densite = Convert.ToDouble(GetSqlNumericValue("DENSITE"));
                    flat.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim();

                    flat.InternalName = nuance_etat + "*" + flat.Longueur + "*" + flat.Largeur + "*" + flat.Epaisseur;

                    //int type = TABLE_ARTICLEM["CHUTE"].ToString() == "N" ? 0 : 1;
                    if (TABLE_ARTICLEM["CHUTE"].ToString() == "O")
                    {
                        flat.type = SheetType.Chute;
                    }
                    else if (TABLE_ARTICLEM["CHUTE"].ToString() == "1")
                    {
                        flat.type = SheetType.ToleNeuve; //forcement tole neuve
                        flat.commercial = true;
                    }
                    else
                    {
                        flat.type = SheetType.ToleNeuve;
                    }

                    //condition d'ajout dans la liste ==> format commercial multidim
                    //format monodim  ce qui assure un code article unique
                    if (flat.IsMultiDim = false || flat.commercial)
                    {
                        this.Clipper_articles_Stock.Add(flat);

                    }
                    logFilesheetMonodim.Write("sheet : " + flat.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();

            }

            catch (Exception ie)
            {
                Clipper_articles_Stock = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFilesheetMonodim.Close();

            }

















        }
        /// <summary>
        /// ecriture des tole dans almacam
        /// </summary>
        /// 

        public override void Write()
        {


            try
            {

                UpdatePrice();

            }

            catch (Exception ie)
            {
                Clipper_articles_Stock = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFilesheetMonodim.Close();

            }



        }
        public void UpdatePrice()
        {

            try
            {
                //this.Clipper_articles_Stock; //liste du stock clip
                //this.Almacam_Stock; liste du stock cam
                if (Clipper_articles_Stock.Count > 0 & Clipper_articles_Stock.Count() > 0)
                {
                    foreach (Flat flat in Clipper_articles_Stock)
                    {
                        var sheetlist = new Dictionary<string, IEntity>();

                        sheetlist = Almacam_Stock.Where(p => p.Key.Contains(flat.COARTI + "_")).ToDictionary(p => p.Key, p => p.Value);
                        foreach (IEntity e in sheetlist.Values)
                        {
                            double poids = e.GetFieldValueAsDouble("_WEIGHT") / 1000;
                            e.SetFieldValue("_BUY_COST", flat.PRIXART * poids);
                            e.SetFieldValue("_COMMENTS", "Prix = " + flat.PRIXART + " €/kg");
                            e.SetFieldValue("_AS_SPECIFIC_COST", true);
                            e.Save();
                        }

                        sheetlist = null;
                    }
                    //        flat.COARTI
                    //srtring coda = Clipper_articles_Stock[flat.COARTI]


                }




            }
            catch
            {


            }




        }
        //public virtual void DimInterpretor() { }
       // public virtual void Close() { Clipper_articles_Stock = null; }
    }
    /// <summary>
    /// recuperation des toles multidim
    /// </summary>
    public class Flat : Clipper_Article, IDisposable
    {

        //private List<TubeRond>listeTubeRond;
        public SheetType type; // = SheetType.ToleNeuve;

        public bool commercial = false;
        public double Longueur = 0;
        public double Largeur = 0;
        public double Epaisseur = 0;




    }
    #endregion
       
    #region Class des Tubes

    /// <summary>
    ///   class Clipper_Article_Tube.. derive de clipper article
    /// </summary>
    public class Clipper_Article_Tube : Clipper_Article, IDisposable, IArticleTube
    {
       
        //creation du listener
        //tools
        public string Section_Key; //"_SECTION_CIRCLE" or....
        public List<string> SectionExclusion = new List<string>();
        public List<string> TubeExclusion = new List<string>();
        public List<KeyValuePair<string, long>> TubeExclusion_WithId = new List<KeyValuePair<string, long>>();
        public List<KeyValuePair<string, long>> SectionQualityExclusion_WithId = new List<KeyValuePair<string, long>>();
        //pas de dictionnaire pour le moment
        public List<KeyValuePair<string, long>> SpecificTubeList_WithId = new List<KeyValuePair<string, long>>();
        private List<string> AlmaCam_Material_List = new List<string>();

        public void CloseImport()
        {
            if (SectionExclusion != null) { SectionExclusion.Clear(); }

            if (TubeExclusion != null)
            { TubeExclusion.Clear(); }
            if (SpecificTubeList_WithId != null)
            { TubeExclusion.Clear(); }
            if (TubeExclusion_WithId != null)
            { TubeExclusion_WithId.Clear(); }
            if (SectionQualityExclusion_WithId != null)
            { SectionQualityExclusion_WithId.Clear(); }
            if (AlmaCam_Material_List != null)
            { AlmaCam_Material_List.Clear(); }


           // AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT TUBE", "Import des " + Section_Key +" terminé.");


        }

        /* public JsonTools JSTOOLS;*/

        IList<Clipper_Material> TUBE_LIST = new List<Clipper_Material>();

        public Clipper_Article_Tube()
        {
            //verification de la connexion obdc
            Odbc_Connexion();
        }


        public void GetAlmaCamMateriallist()
        {

            IEntityList AlmaCam_Material_List_entities = contextlocal.EntityManager.GetEntityList("_MATERIAL");
            AlmaCam_Material_List_entities.Fill(false);

            //AlmaCam_Material_List=AlmaCam_Material_List_entities.ToDictionary(x => x.GetFieldValueAsString("_NAME"), x );
            AlmaCam_Material_List_entities = null;
        }
        /// <summary>
        /// retourn une liste d'exclusion des sections ou des tubes
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="entity_uniquestring_field"></param>
        /// <returns></returns>
        public List<string> GetSectionExclusionList(string entitype)
        {
            try
            {

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                exclusionlist.Fill(false);
                List<string> exclusion = new List<string>();

                // Find material           
                //Entity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();
                foreach (IEntity ex in exclusionlist)
                {
                    string uid = ex.GetImplementEntity("_SECTION").GetFieldValueAsString("_NAME").Trim();
                    ex.GetImplementEntity(Section_Key);
                    if (exclusion.Contains(uid) == false)
                    { exclusion.Add(uid); }

                }


                return exclusion;
            }

            catch
            {

                return null;
            }


        }

        /// <summary>
        /// retourn une liste d'exclusion des sections ou des tubes
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="entity_uniquestring_field"></param>
        /// <returns></returns>
        public List<string> GetExclusionList(string entitype, string uid_field)
        {
            try
            {

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                //exclusionlist[0].GetFieldValueAsEntity("_SECTION").ImplementedEntityType.Key
                exclusionlist.Fill(false);
                List<string> exclusion = new List<string>();

                // Find material           
                //Entity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();
                foreach (IEntity ex in exclusionlist)
                {

                    string uid = ex.GetFieldValueAsString(uid_field).Trim();
                    if (uid != null)
                    {
                        ex.GetImplementEntity(Section_Key);
                        if (exclusion.Contains(uid) == false)
                        { exclusion.Add(uid); }
                    }
                    else
                    {
                        MessageBox.Show("Certains articles tubes ont des reference nulles, merci de les remplir");
                    }
                }


                return exclusion;
            }

            catch
            {

                return null;
            }


        }

        //traduit du code article les inforamtions en decomposant le code article
        //-->>peu fiable 
        private TubeSpe getTubeInfosFrom_Coda(string coda)
        {

            try
            {

                TubeSpe sp = new TubeSpe();
                string[] infos = null;

                if (coda != null || coda != string.Empty)
                    //recuperation de la section spe du tube

                    infos = coda.Split('*');
                long dim = infos.Count() - 1;
                sp.Section = infos[0].Trim();
                sp.Longueur = Convert.ToDouble(infos[dim]);
                //section
                if (sp.Section.Contains("IPN")) { sp.Section_Spe_Key = "_SECTION_IPN"; }
                else if (sp.Section.Contains("IPE")) { sp.Section_Spe_Key = "_SECTION_IPE"; }
                else if (sp.Section.Contains("UPN")) { sp.Section_Spe_Key = "_SECTION_UPN"; }
                else if (sp.Section.Contains("UPE")) { sp.Section_Spe_Key = "_SECTION_UPE"; }
                else if (sp.Section == "L") { sp.Section_Spe_Key = "_SECTION_L"; }
                else if (sp.Section.Contains("LROUND")) { sp.Section_Spe_Key = "_SECTION_LROUND"; }
                else { sp.Section_Spe_Key = string.Empty; }
                sp.COARTI = coda;
                return sp;

            }

            catch (Exception ie)
            {
                MessageBox.Show(ie.Message); return null;


            }

        }
        /// <summary>
        /// retrourne la liste des entité barre a excluire de tout ajout avec les id des entités barre
        /// </summary>
        public List<KeyValuePair<string, long>> GetExclusionList_WithId(string entitype, string uid_field)
        {
            try
            {
                var list = new List<KeyValuePair<string, long>>();
                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                //exclusionlist[0].GetFieldValueAsEntity("_SECTION").ImplementedEntityType.Key
                exclusionlist.Fill(false);
                List<KeyValuePair<string, long>> exclusion = new List<KeyValuePair<string, long>>();

                // Find material           
                //Entity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();
                foreach (IEntity ex in exclusionlist)
                {

                    string uid = ex.GetFieldValueAsString(uid_field).Trim();
                    if (uid != null)
                    {

                        KeyValuePair<string, long> v = new KeyValuePair<string, long>(null, 0);
                        v = exclusion.SingleOrDefault(x => x.Key == uid);
                        ex.GetImplementEntity(Section_Key);

                        if (v.Key == null)
                        {
                            KeyValuePair<string, long> newtube = new KeyValuePair<string, long>(uid, ex.Id);
                            exclusion.Add(newtube);


                        }
                    }
                    else
                    {
                        MessageBox.Show("Certains articles tubes ont des reference nulles, merci de les remplir");
                    }
                }


                return exclusion;
            }

            catch
            {

                return null;
            }


        }
        /// <summary>
        /// rempli la liste d'exeption, du tuebe spé pour limité le nombre de requetes
        /// et accellerer le processus
        /// </summary>
        /// <returns></returns>
        public List<KeyValuePair<string, long>> getSpecificTubeList()
        {
            try
            {
                var list = new List<KeyValuePair<string, long>>();
                IEntityList specificTubeList;

                specificTubeList = contextlocal.EntityManager.GetEntityList("_BARTUBE");
                //pas de dictionnaire pour le moment
                specificTubeList.Fill(false);
                var specificsectionkeyList = new List<string>() { "_SECTION_IPN", "_SECTION_IPE", "_SECTION_UPN", "_SECTION_UPE", "_SECTION_L", "_SECTION_LROUND" };


                List<KeyValuePair<string, long>> tubelist = new List<KeyValuePair<string, long>>();


                foreach (IEntity tube in specificTubeList)
                {


                    {

                        KeyValuePair<string, long> v = new KeyValuePair<string, long>(null, 0);
                        string entitytypekey = tube.GetFieldValueAsEntity("_SECTION").ImplementedEntityType.Key;
                        string tubereference = tube.GetFieldValueAsString("_REFERENCE");

                        if (specificsectionkeyList.Contains(entitytypekey))
                        {
                            KeyValuePair<string, long> newtube = new KeyValuePair<string, long>(tubereference, tube.Id);
                            tubelist.Add(newtube);
                        }
                    }


                }


                //check de doublons
                var doublons = new List<string>();
                doublons = CheckDoublonsInkeypPairList(tubelist);
                if (doublons.Count() > 0)
                {
                    MessageBox.Show("Doublons detectées dans les tubes spécifiques, certaines données ne pourrons pas etre mise à jour.");
                }
                doublons.Clear();
                return tubelist;
            }

            catch
            {

                return null;
            }


        }
        public IEntity getTubeFromSpecifiTubeList(string Coda)
        {
            try
            {
                //dans le cas ou on a une exclusion --> on est en  mise a jour
                KeyValuePair<string, long> v = new KeyValuePair<string, long>(null, 0);
                IEntity barreEntity = null;

                if (SpecificTubeList_WithId.Count != 0)
                {

                    //pas de gestion de double ici
                    v = SpecificTubeList_WithId.SingleOrDefault(x => x.Key == Coda);
                    if (v.Key != null)
                    {

                        barreEntity = contextlocal.EntityManager.GetEntity(v.Value, "_BARTUBE");
                    }


                }

                return barreEntity;
            }

            catch { return null; }
            finally { }


        }
        
        //recupération de l'existant
        public long getExistingSection(string section_name, string section_key)
        {

            try
            {
                //en cas de lenteur on pourrais stocker toutes les gardes poour eviter de faire troo de requetes
                //nous verrons

                IExtendedEntityList sections;
                IExtendedEntity section = null;
                long section_id = 0;

                if (section_name != "")
                {
                    //sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", name);

                    sections = contextlocal.EntityManager.GetExtendedEntityList(contextlocal.Kernel.GetEntityType(section_key));
                    sections.Fill(false);

                    foreach (IExtendedEntity s in sections)
                    {
                        object o = s.GetFieldValue(section_key + "\\IMPLEMENT__SECTION\\_NAME");
                        if (o.ToString() == section_name)
                        {
                            section_id = s.Id;
                            section = s;
                            break;

                        }
                    }





                    return section_id;

                }
                else
                { return 0; }



            }
            catch (Exception ie) { MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error); return 0; }


        }
        public IEntity getExistingSectionEntity(string section_name, string section_key)
        {

            try
            {
                //en cas de lenteur on pourrais stocker toutes les gardes poour eviter de faire troo de requetes
                //nous verrons

                IExtendedEntityList sections;
                IExtendedEntity section = null;
                //long section_id = 0;

                if (section_name != "")
                {
                    //sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", name);

                    sections = contextlocal.EntityManager.GetExtendedEntityList(contextlocal.Kernel.GetEntityType(section_key));
                    sections.Fill(false);

                    foreach (IExtendedEntity s in sections)
                    {
                        object o = s.GetFieldValue(section_key + "\\IMPLEMENT__SECTION\\_NAME");
                        if (o.ToString() == section_name)
                        {
                            //section_id = s.Id;
                            section = s;
                            break;

                        }
                    }





                    return section.Entity;

                }
                else
                { return null; }



            }
            catch (Exception ie) { MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error); return null; }


        }
        public long getExistingImplementedSection(string section_name, string section_key)
        {

            try
            {
                //en cas de lenteur on pourrais stocker toutes les gardes poour eviter de faire troo de requetes
                //nous verrons

                IExtendedEntityList sections;
                //IExtendedEntity section = null;
                long section_id = 0;

                if (section_name != "")
                {
                    //sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", name);

                    sections = contextlocal.EntityManager.GetExtendedEntityList(contextlocal.Kernel.GetEntityType(section_key));
                    sections.Fill(false);

                    foreach (IExtendedEntity s in sections)

                    {
                        object o = s.GetFieldValue(section_key + "\\IMPLEMENT__SECTION\\_NAME");
                        object id = s.GetFieldValue(section_key + "\\IMPLEMENT__SECTION");
                        if (o.ToString() == section_name)
                        {
                            section_id = (long)id;

                            //section = s;
                            break;

                        }

                    }





                    return section_id;

                }
                else
                { return 0; }



            }
            catch (Exception ie) { MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error); return 0; }


        }


        /// <summary>
        /// creation des section
        /// si les entités existent alors elle sont retournées sinon une nouvelle entité est cree
        /// </summary>
        /// <param name="section_name"></param>
        /// <param name="section_key"></param>
        /// <param name="created">retourne true si l'entitté est nouvelement cree a l'appel de la methode</param>
        /// <returns></returns>
        public IEntity Create_Section_If_Not_Exists(string section_name, string section_key, out bool created)
        {
            created = false;

            try
            {
                IEntity sectionEntity = null;
                if (SectionExclusion.Contains(section_name) == false)

                {
                    string key = Guid.NewGuid().ToString();
                    sectionEntity = contextlocal.EntityManager.CreateEntity(Section_Key);
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_KEY", key);
                    created = true;

                }
                else
                {
                    sectionEntity = getExistingSectionEntity(section_name, Section_Key);
                }
                return sectionEntity;
            }

            catch { return null; }
            finally { }



        }
        public IEntity Create_Barre_If_Not_Exists(string barre_name, string section_key, IEntity quality, IEntity sectionentity, out bool created)
        {
            created = false;
            try
            {
                //dans le cas ou on a une exclusion --> on est en  mise a jour
                KeyValuePair<string, long> v = new KeyValuePair<string, long>(null, 0);
                IEntity barreEntity = null;

                if (TubeExclusion_WithId.Count != 0)
                {

                    //if (TubeExclusion.Contains(tuberond.COARTI) == false)
                    //v = TubeExclusion_WithId.SingleOrDefault(x => x.Key == tuberond.COARTI);
                    v = TubeExclusion_WithId.SingleOrDefault(x => x.Key == barre_name);
                    if (v.Key == null)
                    {
                        //creation
                        barreEntity = contextlocal.EntityManager.CreateEntity("_BARTUBE");
                        barreEntity.SetFieldValue("_REFERENCE", barre_name);
                        barreEntity.SetFieldValue("_SECTION", sectionentity);
                        barreEntity.SetFieldValue("_QUALITY", quality);
                        barreEntity.Save();
                        created = true;
                    }
                    else
                    {//mise a jour 

                        barreEntity = contextlocal.EntityManager.GetEntity(v.Value, "_BARTUBE");
                    }


                }//creation initiale
                else
                {


                    barreEntity = contextlocal.EntityManager.CreateEntity("_BARTUBE");
                    barreEntity.SetFieldValue("_REFERENCE", barre_name);
                    barreEntity.SetFieldValue("_SECTION", sectionentity);
                    barreEntity.SetFieldValue("_QUALITY", quality);
                    barreEntity.Save();
                }

                return barreEntity;
            }

            catch { return null; }
            finally { }




        }
        public IEntity Create_Section_Quality_If_Not_Exists(IEntity section, IEntity quality)
        {

            try
            {
                IEntity s = null;
                IEntityList sections = contextlocal.EntityManager.GetEntityList("_SECTION_QUALITY", LogicOperator.And, "_SECTION", ConditionOperator.Equal, section.GetImplementEntity("_SECTION").Id, "_QUALITY", ConditionOperator.Equal, quality.Id);
                sections.Fill(false);

                if (sections.Count > 0)
                {
                    s = sections.FirstOrDefault();
                }//creation
                else
                {
                    //string key = Guid.NewGuid().ToString();
                    //update prix article// prix au mettre /1000 pour obtenir le prix au mm

                    s = contextlocal.EntityManager.CreateEntity("_SECTION_QUALITY");
                    s.SetFieldValue("_SECTION", section.GetImplementEntity("_SECTION"));
                    s.Save();
                }
                //s = contextlocal.EntityManager.CreateEntity(Section_Key); }


                return s;
            }
            catch (Exception er) { MessageBox.Show(er.Message); return null; }
        }

        //check integrity
        ///prevoir d'enlever tous les parametres ancienement //CheckTubeIntegrity(double longeurTube, IEntity quality)
        public virtual bool Check_Tube_Integrity(double longeurTube, IEntity quality)
        {
            bool rst = true;
            try
            {

                if (longeurTube == 0) { rst = rst & false; }
                if (quality == null) { rst = rst & false; }

                return rst;
            }
            catch (Exception ie) { MessageBox.Show(ie.Message); return rst; }


        }
        /// <summary>
        /// lecture des tubes de la base clipper
        /// </summary>
        public virtual void ReadTubes() { }
        /// <summary>
        /// ecriture des tubes dans almacam
        /// </summary>
        public virtual void WriteTubes() { }

       
    }
    /// <summary>
    /// class TubeRond.. derive de clipper article
    /// class tube rond
    /// on ajoute longeur et diametre
    /// </summary>
    public class TubeRond : Clipper_Article, IDisposable
    {


        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Rond;
        public double Diametre;
        public double Longueur;
        public double epaisseur;


    }
    /// <summary>
    /// class TubeRec.. derive de clipper article
    /// </summary>
    public class TubeRec : Clipper_Article, IDisposable
    {

        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Rectangle;
        public double Largeur;
        public double Hauteur;
        public double Longueur;
        public double epaisseur;

    }

    public class TubeSpe : Clipper_Article, IDisposable
    {

        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Speciaux;
        public double Longueur;
        public string Section;
        public string Section_Spe_Key;


    }

    /// <summary>
    /// recuperation des tubes ronds
    /// </summary>

    /// <summary>
    /// recuperation des tubes ronds dans une liste de d'objet tubes ronds
    /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
    /// les methodes virtuelles read : lise le contenu de lcipper avec le lien odbc 
    /// les methodes virtuelles write : ecrive les tubes selon la typologie
    /// </summary>
    public class Clipper_Import_Tube_Rond : Clipper_Article_Tube, IDisposable
    {
        
        private TypeTube type = TypeTube.Rond;
        public double Diametre;
        public double Longueur;
        private List<TubeRond> List_Tube_Ronds;

        TextWriterTraceListener logFileTubeRond = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "TUBE_ROND_" + Properties.Resources.ImportTubeLog);

        /// <summary>
        /// constructeur obligatoire pour ajuster les elements a controler
        /// </summary>
        /// <param name="context"></param>
        public Clipper_Import_Tube_Rond(IContext context)
        {

   

                this.contextlocal = context;
                Section_Key = "_SECTION_CIRCLE";
                SectionExclusion = GetSectionExclusionList(Section_Key);
                TubeExclusion = GetExclusionList("_BARTUBE", "_REFERENCE");
                TubeExclusion_WithId = GetExclusionList_WithId("_BARTUBE", "_REFERENCE");
                AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT TUBE", "Import des section " + Section_Key);
            
        }



        public bool CheckspecificTubeIntegrity(TubeRond tr)
        {
            bool rst;
            //verification matiere
            logFileTubeRond.Write(tr.COARTI);
            rst = Check_Tube_Integrity(tr.Longueur, tr.Material);
            //cas du diamtre null 
            if (tr.Diametre * tr.epaisseur == 0)
            {
                logFileTubeRond.Write("longeur ou epaisseur nulle");
                rst = rst & false;
            }
            //cas du diamtre null 
            //if (tr.epaisseur == 0) { rst = rst & false; }
            if (tr.Diametre - tr.epaisseur < 0) { rst = rst & false; }


            return rst;
        }

        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond de la base clipper
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void ReadTubes()
        {

            try
            {
                //creation de la liste des tube ronds
                //listeTubeRond = new List<TubeRond>();
                this.List_Tube_Ronds = new List<TubeRond>();
                //recuperation des tube rectangluaires
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.TubeRond");
                logFileTubeRond.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                int ii = 0;
                this.DbCommand.CommandText = sql_tube_rec;
                long tubecaptured = 0;
                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    var tuberond = new TubeRond();
                    tuberond.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    tuberond.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim();
                    tuberond.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    tuberond.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                   
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM_TUBE["EPAISSEUR"]);
                    tuberond.Densite = Convert.ToDouble(GetSqlNumericValue("DENSITE"));
                    tuberond.IsMultiDim = Convert.ToBoolean(TABLE_ARTICLEM["MULTIDIM"]);

                    if (tuberond.IsMultiDim)
                    {
                        //multidim
                        tuberond.Longueur = Convert.ToDouble(GetSqlNumericValue("DIMENSIO_DIM1"));
                        tuberond.Diametre = Convert.ToDouble(GetSqlNumericValue("DIM2"));
                        tuberond.epaisseur = Convert.ToDouble(GetSqlNumericValue("DIM1"));
                        tuberond.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + tuberond.Longueur;

                    }
                    else
                    {//monodim
                        tuberond.Longueur = Convert.ToDouble(GetSqlNumericValue("DIM1"));
                        tuberond.Diametre = Convert.ToDouble(GetSqlNumericValue("DIM3"));
                        tuberond.epaisseur = Convert.ToDouble(GetSqlNumericValue("DIM2"));
                        tuberond.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    }


                    tuberond.Material = tuberond.GetGrade(contextlocal, tuberond.Nuance, tuberond.Etat);
                    //tuberond.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + tuberond.Longueur;
                    //modification prix
                    double price = Convert.ToDouble(GetSqlNumericValue("PRIXART"));
                    string unit = TABLE_ARTICLEM["UnitePrix"].ToString().Trim();
                    tuberond.PRIXART = GetPrice(unit, price, tuberond.Longueur, 0);

                    if (CheckspecificTubeIntegrity(tuberond))
                    {
                        tubecaptured++;
                        this.List_Tube_Ronds.Add(tuberond);
                    }

                    //logFileTubeRond.Write("Tube : " + tuberond.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();

                logFileTubeRond.Write("Tube : " + tubecaptured + " capturés");
                //return listeTubeRond;
            }

            catch (Exception ie)
            {
                List_Tube_Ronds = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFileTubeRond.Close();

                /*return listeTubeRond;*/
            }

        }

        /// <summary>
        /// ecriture/mise a jour dans la base almacam des tubes 
        /// </summary>
        public override void WriteTubes()
        {


            if (List_Tube_Ronds.Any())
            {

                foreach (TubeRond tuberond in List_Tube_Ronds)
                {

                    //creation de la section (type...)//
                    IEntity sectionEntity;
                    IEntity section;
                    IEntity barreEntity;
                    IEntity sectionQuality;

                    string section_name = String.Format("Tube_Rond*{0}*{1}", tuberond.Diametre, tuberond.epaisseur);
                    string description = String.Format("Tube Rond Diamete={0} mm Epaisseur={1} mm", tuberond.Diametre, tuberond.epaisseur);
                    //recuperation de la section
                    //TubeExclusion_WithId
                    //creation du tube


                    //creation conditionnelle de la section
                    bool created = false;//true si nouvelle section
                    sectionEntity = Create_Section_If_Not_Exists(section_name, Section_Key, out created);
                    //cas du tube nouvelement cree
                    if (created)
                    {
                        //string key = Guid.NewGuid().ToString();
                        //sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_KEY", key);
                        sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", section_name);// +"x"+ep.ToString());;
                        sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_STANDARD", true);
                        sectionEntity.SetFieldValue("_P_D", tuberond.Diametre);
                        sectionEntity.SetFieldValue("_P_T", tuberond.epaisseur);
                        //descpription //
                        sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_DESCRIPTION", description);

                        //sectionEntity.SetFieldValue("_P_T", this.ep);           
                        sectionEntity.Complete = true;

                        logFileTubeRond.Write("Tube : " + tuberond.COARTI + " sauvegarde");
                        sectionEntity.Save();

                        //creation de la qualité
                        tuberond.Material = tuberond.GetGrade(contextlocal, tuberond.Nuance, tuberond.Etat);
                        //section = sectionEntity.GetImplementEntity("_SECTION");
                        sectionQuality = Create_Section_Quality_If_Not_Exists(sectionEntity, tuberond.GetGrade(contextlocal, tuberond.Nuance, tuberond.Etat));
                        sectionQuality.SetFieldValue("_QUALITY", tuberond.Material.Id);
                        sectionQuality.Save();
                    }

                    else
                    //si elle existe deja on recupere la section //
                    {

                        sectionEntity = getExistingSectionEntity(section_name, Section_Key);
                        sectionQuality = Create_Section_Quality_If_Not_Exists(sectionEntity, tuberond.GetGrade(contextlocal, tuberond.Nuance, tuberond.Etat));


                    }

                    created = false;
                    string barrename;
                    if (tuberond.IsMultiDim == true)
                    {
                        barrename = tuberond.COARTI + "*" + tuberond.Longueur;
                    }
                    else { barrename = tuberond.COARTI; }
                    //creation de la barre
                    IEntity quality = tuberond.GetGrade(contextlocal, tuberond.Nuance, tuberond.Etat);
                    barreEntity = Create_Barre_If_Not_Exists(barrename, Section_Key, quality, sectionEntity.GetImplementEntity("_SECTION"), out created);
                    if (created)
                    {
                        if (barreEntity != null)
                        {
                            logFileTubeRond.Write("section " + description + " sauvegardée");
                            barreEntity.SetFieldValue("_QUALITY", tuberond.Material.Id32);
                            barreEntity.SetFieldValue("_LENGTH", tuberond.Longueur);
                            barreEntity.Save();

                        }
                    }

                    if (barreEntity != null)
                    {   //on set les valeurs
                        if (tuberond.IsMultiDim == false)
                        {  //monodim on declare les 
                            barreEntity.SetFieldValue("_AS_SPECIFIC_COST", true);
                            barreEntity.SetFieldValue("_BUY_COST", tuberond.PRIXART);
                        }
                        else
                        {
                            //recuperation de lid de section dans la table des sections
                            section = sectionEntity.GetImplementEntity("_SECTION");

                            if (tuberond.IsMultiDim)
                            {   //update prix article// prix au mettre /1000 pour obtenir le prix au mm
                                sectionQuality.SetFieldValue("_SECTION", section.Id);
                                sectionQuality.SetFieldValue("_BUY_COST", tuberond.PRIXART / 1000);
                                sectionQuality.Save();
                            }
                            //ATTENTION recuperation des infos de l'implemented sectio

                        }

                        barreEntity.Save();

                    }


                    logFileTubeRond.Write("section " + description + "lng " + tuberond.Longueur + " sauvegardée");
                }

                List_Tube_Ronds.Clear();

            }

            else
            {
                logFileTubeRond.Write("fail to import tube rond : no tube found");

            }

            this.CloseImport();
           

        }


     
    }


    /// <summary>
    /// recuperation des ronds dans une liste de d'objet rond
    /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
    /// les methodes virtuelles read : lise le contenu de lcipper avec le lien odbc 
    /// les methodes virtuelles write : ecrive les tubes selon la typologie
    /// </summary>
    public class Clipper_Import_Rond : Clipper_Article_Tube, IDisposable
    {
        
        private TypeTube type = TypeTube.Rond;
        public double Diametre;
        public double Longueur;
        private List<TubeRond> List_Ronds;

        TextWriterTraceListener logFileRond = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_ROND_" + Properties.Resources.ImportTubeLog);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public Clipper_Import_Rond(IContext context)
        {


         
                this.contextlocal = context;
                Section_Key = "_SECTION_PLAIN_ROUND";
                SectionExclusion = GetSectionExclusionList(Section_Key);
                TubeExclusion = GetExclusionList("_BARTUBE", "_REFERENCE");
                TubeExclusion_WithId = GetExclusionList_WithId("_BARTUBE", "_REFERENCE");
                AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT TUBE", "Import des sections " + Section_Key);
         
            /*
            this.contextlocal = context;
            Section_Key = "_SECTION_PLAIN_ROUND";
            SectionExclusion = GetSectionExclusionList(Section_Key);
            TubeExclusion = GetExclusionList("_BARTUBE", "_REFERENCE");
            */

        }


        /// <summary>
        /// verification de l'integrité du rond
        /// </summary>
        /// <param name="tr"></param>
        /// <returns></returns>
        public bool CheckspecificTubeIntegrity(TubeRond tr)
        {
            bool rst;
            //verification matiere
            logFileRond.Write(tr.COARTI);
            rst = Check_Tube_Integrity(tr.Longueur, tr.Material);
            //cas du diamtre null 
            if (tr.Diametre == 0)
            {
                logFileRond.Write("diametre null ");
                rst = rst & false;
            }
            //cas du diamtre null 



            return rst;
        }
        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void ReadTubes()
        {
            //creation de la liste des tube ronds
            //List<TubeRond> listeTubeRond;//= new List<TubeRond>();

            try
            {
                //creation de la liste des tube ronds
                //listeTubeRond = new List<TubeRond>();
                this.List_Ronds = new List<TubeRond>();
                //recuperation des tube rectangluaires
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.Rond");
                logFileRond.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                long ii = 0;
                long tubecaptured = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    TubeRond rond = new TubeRond();
                    rond.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    rond.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim();
                    rond.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    rond.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM_TUBE["EPAISSEUR"]);
                    rond.Densite = Convert.ToDouble(GetSqlNumericValue("DENSITE"));
                    rond.IsMultiDim = Convert.ToBoolean(TABLE_ARTICLEM["MULTIDIM"]);

                    if (rond.IsMultiDim)
                    {
                        //multidim
                        rond.Longueur = Convert.ToDouble(GetSqlNumericValue("DIMENSIO_DIM1"));
                        rond.Diametre = Convert.ToDouble(GetSqlNumericValue("DIM1"));

                        rond.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + rond.Longueur;

                    }
                    else
                    {//monodim
                        rond.Longueur = Convert.ToDouble(GetSqlNumericValue("DIM1"));
                        rond.Diametre = Convert.ToDouble(GetSqlNumericValue("DIM2"));

                        rond.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    }


                    rond.Material = rond.GetGrade(contextlocal, rond.Nuance, rond.Etat);

                    //modification prix
                    double price = Convert.ToDouble(GetSqlNumericValue("PRIXART"));
                    string unit = TABLE_ARTICLEM["UnitePrix"].ToString().Trim();
                    rond.PRIXART = GetPrice(unit, price, rond.Longueur, 0);

                    //tuberond.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + tuberond.Longueur;

                    if (CheckspecificTubeIntegrity(rond))
                    {
                        tubecaptured++;
                        this.List_Ronds.Add(rond);
                    }



                };

                logFileRond.Write("Tube : " + tubecaptured + " capturés");
                TABLE_ARTICLEM.Close();
                //logFileRond.Flush();
                //logFileRond.Close();
                //return listeTubeRond;
            }

            catch (Exception ie)
            {
                List_Ronds = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFileRond.Close();

                /*return listeTubeRond;*/
            }

        }


        //public void create(IContext contextlocal,string section_name, double diam, double ep, double lng, double cost)

        /// <summary>
        /// ecriture/mise a jour dans la base almacam des tubes 
        /// </summary>
        public override void WriteTubes()
        {


            if (List_Ronds.Any())
            {

                foreach (TubeRond rond in List_Ronds)
                {

                    //creation de la section (type...)//
                    IEntity sectionEntity;
                    IEntity section;
                    IEntity barreEntity;
                    IEntity sectionQuality;

                    string section_name = String.Format("Rond*{0}", rond.Diametre);
                    string description = String.Format("Rond Diamete={0} mm", rond.Diametre);
                    //recuperation de la section
                    //TubeExclusion_WithId
                    //creation du tube


                    //creation conditionnelle de la section
                    bool created = false;//true si nouvelle section
                    sectionEntity = Create_Section_If_Not_Exists(section_name, Section_Key, out created);
                    //cas du tube nouvelement cree
                    if (created)
                    {
                        //string key = Guid.NewGuid().ToString();
                        //sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_KEY", key);
                        sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", section_name);// +"x"+ep.ToString());;
                        sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_STANDARD", true);
                        sectionEntity.SetFieldValue("_P_D", rond.Diametre);

                        //descpription //
                        sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_DESCRIPTION", description);

                        //sectionEntity.SetFieldValue("_P_T", this.ep);           
                        sectionEntity.Complete = true;

                        logFileRond.Write("Tube : " + rond.COARTI + " sauvegarde");
                        sectionEntity.Save();

                        //creation de la qualité
                        rond.Material = rond.GetGrade(contextlocal, rond.Nuance, rond.Etat);
                        //section = sectionEntity.GetImplementEntity("_SECTION");
                        sectionQuality = Create_Section_Quality_If_Not_Exists(sectionEntity, rond.GetGrade(contextlocal, rond.Nuance, rond.Etat));
                        sectionQuality.SetFieldValue("_QUALITY", rond.Material.Id);
                        sectionQuality.Save();
                    }

                    else
                    //si elle existe deja on recupere la section //
                    {

                        sectionEntity = getExistingSectionEntity(section_name, Section_Key);
                        sectionQuality = Create_Section_Quality_If_Not_Exists(sectionEntity, rond.GetGrade(contextlocal, rond.Nuance, rond.Etat));


                    }

                    created = false;
                    string barrename;
                    if (rond.IsMultiDim == true)
                    {
                        barrename = rond.COARTI + "*" + rond.Longueur;
                    }
                    else { barrename = rond.COARTI; }
                    //creation de la barre
                    IEntity quality = rond.GetGrade(contextlocal, rond.Nuance, rond.Etat);
                    barreEntity = Create_Barre_If_Not_Exists(barrename, Section_Key, quality, sectionEntity.GetImplementEntity("_SECTION"), out created);
                    if (created)
                    {
                        if (barreEntity != null)
                        {
                            logFileRond.Write("section " + description + " sauvegardée");
                            barreEntity.SetFieldValue("_QUALITY", rond.Material.Id32);
                            barreEntity.SetFieldValue("_LENGTH", rond.Longueur);
                            barreEntity.Save();

                        }
                    }

                    if (barreEntity != null)
                    {   //on set les valeurs
                        if (rond.IsMultiDim == false)
                        {  //monodim on declare les 
                            barreEntity.SetFieldValue("_AS_SPECIFIC_COST", true);
                            barreEntity.SetFieldValue("_BUY_COST", rond.PRIXART);
                        }
                        else
                        {
                            //recuperation de lid de section dans la table des sections
                            section = sectionEntity.GetImplementEntity("_SECTION");

                            if (rond.IsMultiDim)
                            {   //update prix article// prix au mettre /1000 pour obtenir le prix au mm
                                sectionQuality.SetFieldValue("_SECTION", section.Id);
                                sectionQuality.SetFieldValue("_BUY_COST", rond.PRIXART / 1000);
                                sectionQuality.Save();
                            }
                            //ATTENTION recuperation des infos de l'implemented sectio

                        }

                        barreEntity.Save();

                    }


                    logFileRond.Write("section " + description + "lng " + rond.Longueur + " sauvegardée");
                }

            }
            else
            {
                logFileRond.Write("fail to import tube rond : no tube found");

            }


        }


       




    }


    /// <summary>
    /// recuperation des rectangles dans une liste de d'objet rectangles
    /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
    /// les methodes virtuelles read : lise le contenu de lcipper avec le lien odbc 
    /// les methodes virtuelles write : ecrive les tubes selon la typologie
    /// </summary>

    public class Clipper_Import_Tube_Rectangle : Clipper_Article_Tube, IDisposable
    {
        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Rectangle;
        public double Diametre;
        public double Longueur;
        private List<TubeRec> List_Recs;

        TextWriterTraceListener logFileRec = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_REC_" + Properties.Resources.ImportTubeLog);


        public Clipper_Import_Tube_Rectangle(IContext context)
        {
            //verification de la connexion obdc
      
                this.contextlocal = context;
                Section_Key = "_SECTION_SHARP_RECTANGLE";
                SectionExclusion = GetSectionExclusionList(Section_Key);
                TubeExclusion = GetExclusionList("_BARTUBE", "_REFERENCE");
                AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT TUBE", "Import des sections " + Section_Key);
            
        }



        /// <summary>
        /// verification de l'integrité du rec
        /// </summary>
        /// <param name="tr"></param>
        /// <returns></returns>
        public bool CheckspecificTubeIntegrity(TubeRec tr)
        {
            bool rst;
            //verification matiere
            logFileRec.Write(tr.COARTI);
            rst = Check_Tube_Integrity(tr.Longueur, tr.Material);
            //cas du diamtre null 
            if (tr.Largeur * tr.Hauteur == 0)
            {
                logFileRec.Write("longeur ou largeur null ");
                rst = rst & false;
            }
            //cas de dimenssion inferieure a l'epaisseur

            //if (tr.Largeur < tr.Largeur )
            var numbers = new List<double> { tr.Largeur, tr.Largeur };
            double min = numbers.Min();
            if (min - tr.epaisseur < 0) { rst = rst & false; logFileRec.Write("epaisseur trop importante "); }


            return rst;
        }
                     
        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void ReadTubes()
        {
            //creation de la liste des tube ronds
            //List<TubeRond> listeTubeRond;//= new List<TubeRond>();

            try
            {
                //creation de la liste des tube ronds
                //listeTubeRond = new List<TubeRond>();
                this.List_Recs = new List<TubeRec>();
                //recuperation des tube rectangluaires
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.TubeRectangle");
                logFileRec.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                long ii = 0;
                long tubecaptured = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    var  tuberec = new TubeRec();
                    tuberec.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    tuberec.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim(); ;
                    tuberec.Type = Convert.ToInt32(GetSqlNumericValue("TYPE"));
                    tuberec.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    tuberec.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    
                   
                    tuberec.IsMultiDim = Convert.ToBoolean(TABLE_ARTICLEM["MULTIDIM"]);
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM["EPAISSEUR"]);
                    tuberec.Densite = Convert.ToDouble(GetSqlNumericValue("DENSITE"));
                    tuberec.Material = tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat);
                    if (tuberec.IsMultiDim)
                    {

                        //cas du carre 14
                        if (tuberec.Type == 14)
                        {
                            tuberec.Largeur = Convert.ToDouble(GetSqlNumericValue("DIM1"));
                            tuberec.Hauteur = tuberec.Largeur; // Convert.ToDouble(GetSqlNumericValue("HAUTEUR"));
                            tuberec.epaisseur = Convert.ToDouble(GetSqlNumericValue("DIM2"));
                            tuberec.Longueur = Convert.ToDouble(GetSqlNumericValue("DIMENSIO_DIM1"));
                        }//cas du rectangle 17
                        else if (tuberec.Type == 17)
                        {
                            tuberec.Hauteur = Convert.ToDouble(GetSqlNumericValue("DIM2"));
                            tuberec.Largeur = Convert.ToDouble(GetSqlNumericValue("DIM1"));
                            tuberec.epaisseur = Convert.ToDouble(GetSqlNumericValue("DIM3"));
                            tuberec.Longueur = Convert.ToDouble(GetSqlNumericValue("DIMENSIO_DIM1"));
                        }

                    }
                    else
                    {

                        //cas du carre
                        if (tuberec.Type == 14)
                        {
                            tuberec.Largeur = Convert.ToDouble(GetSqlNumericValue("DIM2"));
                            tuberec.Hauteur = tuberec.Largeur; // Convert.ToDouble(GetSqlNumericValue("HAUTEUR"));
                            tuberec.epaisseur = Convert.ToDouble(GetSqlNumericValue("DIM3"));
                            tuberec.Longueur = Convert.ToDouble(GetSqlNumericValue("DIM1"));

                        }//cas du rectangle une dim en plus
                        else if (tuberec.Type == 17)
                        {
                            tuberec.Hauteur = Convert.ToDouble(GetSqlNumericValue("DIM3"));
                            tuberec.Largeur = Convert.ToDouble(GetSqlNumericValue("DIM2"));
                            tuberec.epaisseur = Convert.ToDouble(GetSqlNumericValue("DIM4"));
                            tuberec.Longueur = Convert.ToDouble(GetSqlNumericValue("DIM1"));

                        }

                    }

                    //definition du prix
                    double price = Convert.ToDouble(GetSqlNumericValue("PRIXART"));
                    string unit = TABLE_ARTICLEM["UnitePrix"].ToString().Trim();
                    tuberec.PRIXART = GetPrice(unit, price, tuberec.Longueur, 0);

                    tuberec.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + tuberec.Longueur;

                    if (CheckspecificTubeIntegrity(tuberec))
                    {
                        tubecaptured++;
                        this.List_Recs.Add(tuberec);
                    }

                    // logFileRec.Write("Tube : " + tuberec.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();
                logFileRec.Write("Tube : " + tubecaptured + " capturés");
                //logFileRond.Flush();
                //logFileRond.Close();
                //return listeTubeRond;
            }

            catch (Exception ie)
            {
                List_Recs = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFileRec.Close();

                /*return listeTubeRond;*/
            }

        }


        //public void create(IContext contextlocal,string section_name, double diam, double ep, double lng, double cost)
        /// <summary>
        /// ecriture/mise a jour dans la base almacam des tubes 
        /// </summary>
        public override void WriteTubes()
        {
            try
            {

                if (List_Recs.Any())
                {

                    foreach (TubeRec tuberec in List_Recs)
                    {

                        //creation de la section (type...)//
                        IEntity sectionEntity;
                        IEntity section;
                        IEntity barreEntity;
                        IEntity sectionQuality;

                        string section_name = String.Format("Rec*{0}*{1}*{2}", tuberec.Largeur, tuberec.Hauteur, tuberec.epaisseur);
                        string description = String.Format("Rec hauteur={1} mm Larg={0} mm  Ep={2} mm", tuberec.Hauteur, tuberec.Largeur, tuberec.epaisseur);
                        //recuperation de la section
                        //TubeExclusion_WithId
                        //creation du tube


                        //creation conditionnelle de la section
                        bool created = false;//true si nouvelle section
                        sectionEntity = Create_Section_If_Not_Exists(section_name, Section_Key, out created);
                        //cas du tube nouvelement cree
                        if (created)
                        {
                            //string key = Guid.NewGuid().ToString();
                            //sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_KEY", key);
                            sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", section_name);// +"x"+ep.ToString());;
                            sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_STANDARD", true);
                            sectionEntity.SetFieldValue("_P_B", tuberec.Largeur);
                            sectionEntity.SetFieldValue("_P_H", tuberec.Hauteur);
                            sectionEntity.SetFieldValue("_P_T", tuberec.epaisseur);

                            //descpription //
                            sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_DESCRIPTION", description);

                            //sectionEntity.SetFieldValue("_P_T", this.ep);           
                            sectionEntity.Complete = true;

                            logFileRec.Write("Tube : " + tuberec.COARTI + " sauvegarde");
                            sectionEntity.Save();

                            //creation de la qualité
                            tuberec.Material = tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat);
                            //section = sectionEntity.GetImplementEntity("_SECTION");
                            sectionQuality = Create_Section_Quality_If_Not_Exists(sectionEntity, tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat));
                            sectionQuality.SetFieldValue("_QUALITY", tuberec.Material.Id);
                            sectionQuality.Save();
                        }

                        else
                        //si elle existe deja on recupere la section //
                        {

                            sectionEntity = getExistingSectionEntity(section_name, Section_Key);
                            sectionQuality = Create_Section_Quality_If_Not_Exists(sectionEntity, tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat));


                        }

                        created = false;
                        string barrename;
                        if (tuberec.IsMultiDim == true)
                        {
                            barrename = tuberec.COARTI + "*" + tuberec.Longueur;
                        }
                        else { barrename = tuberec.COARTI; }
                        //creation de la barre
                        IEntity quality = tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat);
                        barreEntity = Create_Barre_If_Not_Exists(barrename, Section_Key, quality, sectionEntity.GetImplementEntity("_SECTION"), out created);
                        if (created)
                        {
                            if (barreEntity != null)
                            {
                                logFileRec.Write("section " + description + " sauvegardée");
                                barreEntity.SetFieldValue("_QUALITY", tuberec.Material.Id32);
                                barreEntity.SetFieldValue("_LENGTH", tuberec.Longueur);
                                barreEntity.Save();

                            }
                        }

                        if (barreEntity != null)
                        {   //on set les valeurs
                            if (tuberec.IsMultiDim == false)
                            {  //monodim on declare les 
                                barreEntity.SetFieldValue("_AS_SPECIFIC_COST", true);
                                barreEntity.SetFieldValue("_BUY_COST", tuberec.PRIXART);
                            }
                            else
                            {
                                //recuperation de lid de section dans la table des sections
                                section = sectionEntity.GetImplementEntity("_SECTION");

                                if (tuberec.IsMultiDim)
                                {   //update prix article// prix au mettre /1000 pour obtenir le prix au mm
                                    sectionQuality.SetFieldValue("_SECTION", section.Id);
                                    sectionQuality.SetFieldValue("_BUY_COST", tuberec.PRIXART );
                                    sectionQuality.Save();
                                }
                                //ATTENTION recuperation des infos de l'implemented sectio

                            }

                            barreEntity.Save();

                        }


                        logFileRec.Write("section " + description + "lng " + tuberec.Longueur + " sauvegardée");
                    }
                    logFileRec.Write(" import tube rec  done and found");
                }
                else
                {
                    logFileRec.Write("fail to import tube rec : no tube found");

                }
            }

            catch (Exception ie) { MessageBox.Show(ie.Message); }

        }
       





    }


    /// <summary>
    /// recuperation des carres dans une liste de d'objet carres
    /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
    /// les methodes virtuelles read : lise le contenu de lcipper avec le lien odbc 
    /// les methodes virtuelles write : ecrive les tubes selon la typologie
    /// </summary>
    [Obsolete]
    public class Clipper_Import_Tube_Carre : Clipper_Article_Tube, IDisposable
    {
        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Rectangle;
        public double Diametre;
        public double Longueur;
        private List<TubeRec> List_Recs;

        TextWriterTraceListener logFileCar = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_REC_" + Properties.Resources.ImportTubeLog);



        public Clipper_Import_Tube_Carre(IContext context)
        {
            this.contextlocal = context;
            Section_Key = "_SECTION_SHARP_RECTANGLE";
            SectionExclusion = GetSectionExclusionList(Section_Key);
            TubeExclusion = GetExclusionList("_BARTUBE", "_REFERENCE");
            AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT TUBE", "Import des sections "+ Section_Key);
        }

        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void ReadTubes()
        {
            //creation de la liste des tube ronds
            //List<TubeRond> listeTubeRond;//= new List<TubeRond>();

            try
            {
                //creation de la liste des tube ronds
                //listeTubeRond = new List<TubeRond>();
                this.List_Recs = new List<TubeRec>();
                //recuperation des tube rectangluaires
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.TubeCarre");
                logFileCar.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                long ii = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    TubeRec tuberec = new TubeRec();
                  
                };

                TABLE_ARTICLEM.Close();
                logFileCar.WriteLine(ii + " Flat Tube found");

            }

            catch (Exception ie)
            {
                List_Recs = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFileCar.Close();

                /*return listeTubeRond;*/
            }

        }


        //public void create(IContext contextlocal,string section_name, double diam, double ep, double lng, double cost)
        public override void WriteTubes()
        {
           

        }


    }


    /// <summary>
    /// recuperation des flats dans une liste de d'objet flat
    /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
    /// les methodes virtuelles read : lise le contenu de lcipper avec le lien odbc 
    /// les methodes virtuelles write : ecrive les tubes selon la typologie
    /// </summary>


    public class Clipper_Import_Tube_Flat : Clipper_Article_Tube, IDisposable
    {
        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Flat;
        public double Diametre;
        public double Longueur;
        public double seuil = 300; //regalge valeur seuil pour detection des toles longeur + largeur =50max + 300max abac arcelor
        private List<TubeRec> List_Recs;

        public Clipper_Import_Tube_Flat(IContext context)
        {
            this.contextlocal = context;
            Section_Key = "_SECTION_FLAT";
            SectionExclusion = GetSectionExclusionList(Section_Key);
            TubeExclusion = GetExclusionList("_BARTUBE", "_REFERENCE");
            AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT TUBE", "Import des sections " + Section_Key);
        }

        TextWriterTraceListener logFileFlat = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_FLAT_" + Properties.Resources.ImportTubeLog);

        /// <summary>
        /// verification de l'integrité du rec
        /// </summary>
        /// <param name="tr"></param>
        /// <returns></returns>
        public bool CheckspecificTubeIntegrity(TubeRec tr)
        {
            bool rst;
            //verification matiere
            logFileFlat.Write(tr.COARTI);
            rst = Check_Tube_Integrity(tr.Longueur, tr.Material);
            //cas du diamtre null 
            if (tr.Largeur * tr.Hauteur == 0)
            {
                logFileFlat.Write("longeur ou largeur null ");
                rst = rst & false;
            }
            //cas de dimenssion inferieure a limite

            //if (tr.Largeur < tr.Largeur )
            var numbers = new List<double> { tr.Largeur, tr.Hauteur };
            double max = numbers.Max();
            if (max > seuil) { rst = rst & false; logFileFlat.Write(" tole detectée car dim > seuil=" + seuil); }


            return rst;
        }


        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond
        /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
        /// lecture des "tole" dont toutes les dimenssion de la section sont iinferieur a la valuer seuil de 300
        /// </summary>
        /// <returns>List<TubeRond></returns>
        /// 
        public override void ReadTubes()
        {
            //creation de la liste des tube ronds
            //List<TubeRond> listeTubeRond;//= new List<TubeRond>();

            try
            {
                this.List_Recs = new List<TubeRec>();
                //recuperation des tube rectangluaires
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.Flat");
                logFileFlat.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                long ii = 0;
                long tubecaptured = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();



                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    //TubeRec tuberec = new TubeRec();
                    var tuberec = new TubeRec();
                    tuberec.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    tuberec.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim(); ;
                    tuberec.Type = Convert.ToInt32(GetSqlNumericValue("TYPE"));
                    tuberec.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    tuberec.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    tuberec.IsMultiDim = Convert.ToBoolean(TABLE_ARTICLEM["MULTIDIM"]);
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM["EPAISSEUR"]);
                    tuberec.Densite = Convert.ToDouble(GetSqlNumericValue("DENSITE"));
                    tuberec.Material = tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat);
                    if (tuberec.IsMultiDim)
                    {


                        tuberec.Hauteur = Convert.ToDouble(GetSqlNumericValue("DIM1"));
                        tuberec.Largeur = Convert.ToDouble(GetSqlNumericValue("DIMENSIO_DIM2"));
                        //tuberec.epaisseur = Convert.ToDouble(GetSqlNumericValue("DIM3"));
                        tuberec.Longueur = Convert.ToDouble(GetSqlNumericValue("DIMENSIO_DIM1"));


                    }
                    else
                    {



                        tuberec.Hauteur = Convert.ToDouble(GetSqlNumericValue("DIM3"));
                        tuberec.Largeur = Convert.ToDouble(GetSqlNumericValue("DIM2"));
                        //tuberec.epaisseur = Convert.ToDouble(GetSqlNumericValue("DIM4"));
                        tuberec.Longueur = Convert.ToDouble(GetSqlNumericValue("DIM1"));



                    }

                    //definition du prix
                    //definition du prix
                    double price = Convert.ToDouble(GetSqlNumericValue("PRIXART"));
                    string unit = TABLE_ARTICLEM["UnitePrix"].ToString().Trim();
                    tuberec.PRIXART = GetPrice(unit, price, tuberec.Longueur, 0);
                    
                  

                    tuberec.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + tuberec.Longueur;

                    if (CheckspecificTubeIntegrity(tuberec))
                    {
                        tubecaptured++;
                        this.List_Recs.Add(tuberec);
                    }

                    logFileFlat.Write("Tube : " + tuberec.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();
                logFileFlat.Write("Tube : " + tubecaptured + " capturés");

            }

            catch (Exception ie)
            {
                List_Recs = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFileFlat.Close();

                /*return listeTubeRond;*/
            }

        }


        //public void create(IContext contextlocal,string section_name, double diam, double ep, double lng, double cost)
        /// <summary>
        /// ecriture/mise a jour dans la base almacam des tubes 
        /// </summary>
        public override void WriteTubes()
        {
            try
            {

                if (List_Recs.Any())
                {

                    foreach (TubeRec tuberec in List_Recs)
                    {

                        //creation de la section (type...)//
                        IEntity sectionEntity;
                        IEntity section;
                        IEntity barreEntity;
                        IEntity sectionQuality;

                        string section_name = String.Format("flat*{0}*{1}", tuberec.Largeur, tuberec.Hauteur);
                        string description = String.Format("flat hauteur={1} mm Larg={0} mm ", tuberec.Hauteur, tuberec.Largeur);
                        //recuperation de la section
                        //TubeExclusion_WithId
                        //creation du tube


                        //creation conditionnelle de la section
                        bool created = false;//true si nouvelle section
                        sectionEntity = Create_Section_If_Not_Exists(section_name, Section_Key, out created);
                        //cas du tube nouvelement cree
                        if (created)
                        {
                            //string key = Guid.NewGuid().ToString();
                            //sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_KEY", key);
                            sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", section_name);// +"x"+ep.ToString());;
                            sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_STANDARD", true);
                            sectionEntity.SetFieldValue("_P_H", tuberec.Largeur);
                            sectionEntity.SetFieldValue("_P_B", tuberec.Hauteur);


                            //descpription //
                            sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_DESCRIPTION", description);

                            //sectionEntity.SetFieldValue("_P_T", this.ep);           
                            sectionEntity.Complete = true;

                            logFileFlat.Write("Tube : " + tuberec.COARTI + " sauvegarde");
                            sectionEntity.Save();

                            //creation de la qualité
                            tuberec.Material = tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat);
                            //section = sectionEntity.GetImplementEntity("_SECTION");
                            sectionQuality = Create_Section_Quality_If_Not_Exists(sectionEntity, tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat));
                            sectionQuality.SetFieldValue("_QUALITY", tuberec.Material.Id);
                            sectionQuality.Save();
                        }

                        else
                        //si elle existe deja on recupere la section //
                        {

                            sectionEntity = getExistingSectionEntity(section_name, Section_Key);
                            sectionQuality = Create_Section_Quality_If_Not_Exists(sectionEntity, tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat));


                        }

                        created = false;
                        string barrename;
                        if (tuberec.IsMultiDim == true)
                        {
                            barrename = tuberec.COARTI + "*" + tuberec.Longueur;
                        }
                        else { barrename = tuberec.COARTI; }
                        //creation de la barre
                        IEntity quality = tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat);
                        barreEntity = Create_Barre_If_Not_Exists(barrename, Section_Key, quality, sectionEntity.GetImplementEntity("_SECTION"), out created);
                        if (created)
                        {
                            if (barreEntity != null)
                            {
                                logFileFlat.Write("section " + description + " sauvegardée");
                                barreEntity.SetFieldValue("_QUALITY", tuberec.Material.Id32);
                                barreEntity.SetFieldValue("_LENGTH", tuberec.Longueur);
                                barreEntity.Save();

                            }
                        }

                        if (barreEntity != null)
                        {   //on set les valeurs
                            if (tuberec.IsMultiDim == false)
                            {  //monodim on declare les 
                                barreEntity.SetFieldValue("_AS_SPECIFIC_COST", true);
                                barreEntity.SetFieldValue("_BUY_COST", tuberec.PRIXART);
                            }
                            else
                            {
                                //recuperation de lid de section dans la table des sections
                                section = sectionEntity.GetImplementEntity("_SECTION");

                                if (tuberec.IsMultiDim)
                                {   //update prix article// prix au mettre /1000 pour obtenir le prix au mm
                                    sectionQuality.SetFieldValue("_SECTION", section.Id);
                                    sectionQuality.SetFieldValue("_BUY_COST", tuberec.PRIXART);
                                    sectionQuality.Save();
                                }
                                //ATTENTION recuperation des infos de l'implemented sectio

                            }

                            barreEntity.Save();

                        }


                        logFileFlat.Write("section " + description + "lng " + tuberec.Longueur + " sauvegardée");
                    }
                    logFileFlat.Write(" import tube rec  done and found");
                }
                else
                {
                    logFileFlat.Write("fail to import tube rec : no tube found");

                }
            }

            catch (Exception ie) { MessageBox.Show(ie.Message); }

        }




    }

    /// <summary>
    /// importe les tubes spéciaux : la racine du nom du tube doit contenir le nom de la section du tube spécial
    /// par exemple IPN80*S235*BRUT*3000 : la strucuture doit etre au minimum comme ceci: [SECTION ALMACAM].....[LONGUEUR] 
    /// ->[IPN80].....[3000] la section choisie est IPN80, la longeur est 3000
    /// </summary>

    public class Clipper_Import_Tube_Speciaux : Clipper_Article_Tube, IDisposable
    {
        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Speciaux;
        private List<TubeRec> List_Recs;


        public Clipper_Import_Tube_Speciaux(IContext context)
        {
            this.contextlocal = context;
            SpecificTubeList_WithId = getSpecificTubeList();
            AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT TUBE", "Import des sections spéciales..." );
           


        }

        /// <summary>
        /// utilise pour les tubes spes uniquement ou le section_key  (SECTION_UPE,SECTION_UPN..) peut varier
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="uid_field">chzamps unique de recherche</param>
        /// <param name="exclusionlist">list cumule d'exclusion.. </param>
        /// 
        /// <returns></returns>
        public void GetExclusionList(ref List<string> exclusion, string entitype, string uid_field)
        {
            try
            {

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                //exclusionlist[0].GetFieldValueAsEntity("_SECTION").ImplementedEntityType.Key
                exclusionlist.Fill(false);
                //List<string> exclusion = new List<string>();
                //exclusionlist.ToList();
                // Find material           
                //Entity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();

                if (exclusionlist.Count > 0)
                {
                    foreach (IEntity ex in exclusionlist)
                    {
                        if (ex.GetFieldValueAsString(uid_field) != null)
                        {
                            string uid = ex.GetFieldValueAsString(uid_field).Trim();
                            exclusion.Add(uid);

                        }




                    }

                }
                //return exclusion;
            }

            catch (Exception ie)
            {

                MessageBox.Show(ie.Message);
                // return null;
            }


        }
        /// <summary>
        /// utilise pour les tubes spes uniquement ou le section_key  (SECTION_UPE,SECTION_UPN..) peut varier
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="uid_field">chzamps unique de recherche</param>
        /// <param name="section_key">Entité section SECTION_UPE,SECTION_UPN.. </param>
        /// <param name="exclusionlist">list cumule d'exclusion.. </param>
        /// 
        /// <returns></returns>
        public void getRestrictedExclusionList(ref List<string> exclusion, string entitype, string uid_field, string section_key)
        {
            try
            {

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                //exclusionlist[0].GetFieldValueAsEntity("_SECTION").ImplementedEntityType.Key
                exclusionlist.Fill(false);
                //List<string> exclusion = new List<string>();
                //exclusionlist.ToList();
                // Find material           
                //Entity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();

                if (exclusionlist.Count > 0)
                {
                    foreach (IEntity ex in exclusionlist)
                    {
                        if (ex.GetFieldValueAsString(uid_field) != null & ex.GetFieldValueAsEntity("_SECTION").ImplementedEntityType.Key == section_key)
                        {
                            string uid = ex.GetFieldValueAsString(uid_field).Trim();
                            exclusion.Add(uid);

                        }




                    }

                }
                //return exclusion;
            }

            catch (Exception ie)
            {

                MessageBox.Show(ie.Message);
                // return null;
            }


        }


        TextWriterTraceListener logFilSpe = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_SPE_" + Properties.Resources.ImportTubeLog);


        /// <summary>
        /// verification de l'integrité du rec
        /// </summary>
        /// <param name="tr"></param>
        /// <returns></returns>
        public bool CheckspecificTubeIntegrity(TubeRec tr)
        {
            bool rst;
            //verification matiere
            logFilSpe.Write(tr.COARTI);
            rst = Check_Tube_Integrity(tr.Longueur, tr.Material);
            //cas du diamtre null 
            /*
            if (tr.Longueur * tr.PRIXART == 0)
            {
                logFilSpe.Write("longeur ou prix null ");
                rst = rst & false;
            }
           
            */
            return rst;
        }


        /// <summary>
        public IEntity Create_Spe_Section_Quality_If_Not_Exists(IEntity section, IEntity quality)
        {


            IEntity s = null;
            IEntityList sections = contextlocal.EntityManager.GetEntityList("_SECTION_QUALITY", LogicOperator.And, "_SECTION", ConditionOperator.Equal, section.Id, "_QUALITY", ConditionOperator.Equal, quality.Id);
            sections.Fill(false);

            if (sections.Count > 0)
            {
                s = sections.FirstOrDefault();
            }//creation
            else
            {
                //string key = Guid.NewGuid().ToString();
                //update prix article// prix au mettre /1000 pour obtenir le prix au mm

                s = contextlocal.EntityManager.CreateEntity("_SECTION_QUALITY");
                s.SetFieldValue("_SECTION", section);
                s.SetFieldValue("_QUALITY", quality);
                s.Save();
            }

            return s;
        }

        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond
        /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
        /// lecture des "tole" dont toutes les dimenssion de la section sont iinferieur a la valuer seuil de 300
        /// </summary>
        /// <returns>List<TubeRond></returns>
        /// 
        public override void ReadTubes()
        {
            //creation de la liste des tube ronds
            //List<TubeRond> listeTubeRond;//= new List<TubeRond>();

            try
            {
                this.List_Recs = new List<TubeRec>();
                //recuperation des tube rectangluaires
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.Profilspeciaux");
                logFilSpe.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                long ii = 0;
                long tubecaptured = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();



                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    var tubespe = new TubeRec();
                    tubespe.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    tubespe.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim(); ;
                    tubespe.Type = Convert.ToInt32(GetSqlNumericValue("TYPE"));
                    tubespe.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    tubespe.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();                    
                    tubespe.IsMultiDim = Convert.ToBoolean(TABLE_ARTICLEM["MULTIDIM"]);
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM["EPAISSEUR"]);
                    tubespe.Densite = Convert.ToDouble(GetSqlNumericValue("DENSITE"));
                    tubespe.Material = tubespe.GetGrade(contextlocal, tubespe.Nuance, tubespe.Etat);
                    if (tubespe.IsMultiDim)
                    {

                        tubespe.Longueur = Convert.ToDouble(GetSqlNumericValue("DIMENSIO_DIM1"));
                        tubespe.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + tubespe.Longueur;

                    }
                    else
                    {


                        tubespe.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                        tubespe.Longueur = Convert.ToDouble(GetSqlNumericValue("DIM1"));



                    }


                    //modification prix
                    double price = Convert.ToDouble(GetSqlNumericValue("PRIXART"));
                    string unit = TABLE_ARTICLEM["UnitePrix"].ToString().Trim();
                    tubespe.PRIXART = GetPrice(unit, price, 0, 0);


                    if (CheckspecificTubeIntegrity(tubespe))
                    {
                        tubecaptured++;
                        this.List_Recs.Add(tubespe);
                    }



                };

                logFilSpe.Write("Tube : " + tubecaptured + " capturés");
                TABLE_ARTICLEM.Close();
                //logFileRond.Flush();
                //logFileRond.Close();
                //return listeTubeRond;
            }

            catch (Exception ie)
            {
                List_Recs = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFilSpe.Close();

                /*return listeTubeRond;*/
            }

        }


        public override void WriteTubes()
        {
            try
            {
                //creation de la section (type...)//
                IEntity sectionEntity;
                //IEntity section;
                IEntity barreEntity;
                //IEntityList barreEntitys;
                IEntity sectionQuality;

                if (List_Recs.Any())
                {
                    foreach (TubeRec tubespe in List_Recs)
                    {
                        //recuperation de l'article
                        //barreEntitys = contextlocal.EntityManager.GetEntityList("_BARTUBE", "_REFERENCE", ConditionOperator.Equal, tubespe.Name);//.CreateEntity("_BARTUBE");
                        //barreEntitys.Fill(false);
                        //if (barreEntitys.Count>0) {
                        //barreEntitys.FirstOrDefault();
                        //doublons non supportés
                        barreEntity = getTubeFromSpecifiTubeList(tubespe.Name);
                        if (barreEntity != null)
                        {
                            if (tubespe.IsMultiDim)
                            {
                                //mettre a jour le prix section

                               


                                    sectionEntity = barreEntity.GetFieldValueAsEntity("_SECTION");

                                    IEntity quality = tubespe.GetGrade(contextlocal, tubespe.Nuance, tubespe.Etat);
                                    //section = sectionEntity.GetImplementEntity("_SECTION");
                                    sectionQuality = Create_Spe_Section_Quality_If_Not_Exists(sectionEntity, tubespe.GetGrade(contextlocal, tubespe.Nuance, tubespe.Etat));
                                    sectionQuality.SetFieldValue("_QUALITY", tubespe.GetGrade(contextlocal, tubespe.Nuance, tubespe.Etat));
                                    sectionQuality.SetFieldValue("_SECTION", sectionEntity.Id);
                                    sectionQuality.SetFieldValue("_BUY_COST", tubespe.PRIXART / 1000);
                                    sectionQuality.Save();
                               

                            }
                            else
                            {
                                //mettre a jour le prix article
                                //monodim on declare les 

                               
                                    barreEntity.SetFieldValue("_AS_SPECIFIC_COST", true);
                                    barreEntity.SetFieldValue("_BUY_COST", tubespe.PRIXART);
                                    barreEntity.Save();
                              
                            }
                        }
                        /*}
                        else
                        {

                            logFilSpe.Write(tubespe.Name + " not found, check for spaces or add length for multidim articles [coda*lenght]");
                            
                        }
                        */






                    }




                }
                //purge mémoire
                CloseImport();
            }
            catch (Exception ie) { MessageBox.Show(ie.Message); }
        }


    }


    #endregion


    #region Founitures


    /// <summary>
    /// import specifique des vis
    /// </summary>
    public class Clipper_Import_Fournitures_Divers : Clipper_Article, IDisposable
    {
        //
        // public double Diametre;
        //public double Longueur;

        private List<Founiture_Divers> List_Fourniture;

        public string Key; //"_SECTION_CIRCLE" or....
                           //public List<string> SectionExclusion = new List<string>();
        public Dictionary<long, string> Founiture_Divers_Exclusion = new Dictionary<long, string>();

        TextWriterTraceListener logFourniture = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_SIMPLE_SUPPLY_" + Properties.Resources.ImportFournitureLog);

        /// <summary>
        /// constructeur obligatoire pour ajuster les elements a controler
        /// </summary>
        /// <param name="context"></param>
        public Clipper_Import_Fournitures_Divers(IContext context)
        {



            this.contextlocal = context;
            contextlocal.TraceLogger.TraceInformation("Test de Connection ODBC");
            Key = "_SIMPLE_SUPPLY";
            Founiture_Divers_Exclusion = GetExclusionList(Key);
            Odbc_Connexion();
            contextlocal.TraceLogger.TraceInformation("Connection ODBC ok");
            AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT FOURINTURE DIVERS", "Import des fournitures vis.. " );
        }



        /// <summary>
        /// retourn une liste d'exclusion des sections ou des tubes
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="entity_uniquestring_field"></param>
        /// <returns></returns>
        public Dictionary<long, string> GetExclusionList(string entitypekey)
        {

            try
            {
                contextlocal.TraceLogger.TraceInformation("creation de la liste d'exclusion");
                IEntityList exclusionlist;

                exclusionlist = contextlocal.EntityManager.GetEntityList(entitypekey);
                exclusionlist.Fill(false);
                Dictionary<long, string> exclusion = new Dictionary<long, string>();
                // Find material           
                //Entity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();
                foreach (IEntity ex in exclusionlist)
                {

                    string coarti = ex.GetImplementEntity("_SUPPLY").GetFieldValueAsString("_REFERENCE").Trim();
                    long uid = ex.GetImplementEntity("_SUPPLY").Id;
                    //ex.GetImplementEntity(Key);

                    if (exclusion.ContainsKey(uid) == false && string.IsNullOrEmpty(coarti) == false)
                    {
                        exclusion.Add(uid, coarti);
                    }

                }

                contextlocal.TraceLogger.TraceInformation("creation de la liste d'exclusion terminée");
                return exclusion;
            }

            catch
            {

                return null;
            }
        }


        /// <summary>
        /// recuperation des ronds dans une liste de d'objet vis de la base clipper
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void Read()
        {
            contextlocal.TraceLogger.TraceInformation("lecture des fournitures clipper en cours");

            try
            {
                //creation de la liste des fournitures

                this.List_Fourniture = new List<Founiture_Divers>();
                //recuperation des tube rectangluaires
                string sql_vis = this.JSTOOLS.getJsonStringParametres("sql.Divers");
                logFourniture.Write("requete utilisée pour l'import des fournitures \r\n " + sql_vis);
                int ii = 0;
                this.DbCommand.CommandText = sql_vis;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    //Founiture_Divers Fourniture = new Founiture_Divers();
                   var Fourniture = new Founiture_Divers();


                    Fourniture.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    Fourniture.AMCLEUNIK = Convert.ToInt64(TABLE_ARTICLEM["AMCLEUNIK"]);
                    Fourniture.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim();
                    Fourniture.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    Fourniture.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    //Fourniture.PRIXART = Convert.ToDouble(GetSqlNumericValue("PRIXART"));
                    //Fourniture.PRIXART = Math.Round(GetPrice(Fourniture.UnitePrix, Convert.ToDouble(GetSqlNumericValue("PRIXART"))), 5);
                    //modification prix
                    double price = Convert.ToDouble(GetSqlNumericValue("PRIXART"));
                    string unit = TABLE_ARTICLEM["UnitePrix"].ToString().Trim();
                    Fourniture.PRIXART = GetPrice(unit, price, 0, 0);

                    Fourniture.DESA1 = TABLE_ARTICLEM["DESA1"].ToString().Trim();
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM_TUBE["EPAISSEUR"]);
                    //rond.Densite = Convert.ToDouble(GetSqlNumericValue("DENSITE"));
                    //rond.IsMultiDim = Convert.ToBoolean(TABLE_ARTICLEM["MULTIDIM"]);



                    /*
                    ///cle unique clipper
                    Fourniture.AMCLEUNIK = Convert.ToInt64(TABLE_ARTICLEM["AMCLEUNIK"]);
                    Fourniture.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    Fourniture.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    Fourniture.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    Fourniture.Densite = Convert.ToDouble(GetSqlNumericValue("DENSITE"));
                    Fourniture.UnitePrix = TABLE_ARTICLEM["UnitePrix"].ToString().Trim();
                    Fourniture.PRIXART = Math.Round(GetPrice(Fourniture.UnitePrix, Convert.ToDouble(GetSqlNumericValue("PRIXART"))), 5);
                    Fourniture.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim(); ;
                    Fourniture.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + GetSqlNumericValue("LNG");
                    Fourniture.DESA1 = TABLE_ARTICLEM["DESA1"].ToString().Trim();
                    */


                    this.List_Fourniture.Add(Fourniture);
                    logFourniture.Write("Fourniture : " + Fourniture.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();
                contextlocal.TraceLogger.TraceInformation("lecture des fournitures terminée");
                //return listeTubeRond;
                AF_ImportTools.SimplifiedMethods.NotifyMessage("IMPORT FOURINTURE DIVERS", "Import des fournitures terminé.. ");
            }

            catch (Exception ie)
            {
                this.List_Fourniture = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFourniture.Close();


            }

        }

        /// <summary>
        /// ecriture dans la base almacam
        /// </summary>
        /// 
        //ecriture des fournitures

        public override void Write()
        {
            try
            {
                contextlocal.TraceLogger.TraceInformation("Import des fournitures en cours");

                if (this.List_Fourniture.Any())
                {


                    foreach (Founiture_Divers fourniture in this.List_Fourniture)
                    {

                        if (!Founiture_Divers_Exclusion.ContainsValue(fourniture.COARTI))
                        {
                            IEntity Fourniture_Entity;
                            //
                            Fourniture_Entity = contextlocal.EntityManager.CreateEntity(Key);
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_REFERENCE", fourniture.COARTI);
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_DESIGNATION", fourniture.DESA1);
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_COMMENTS", fourniture.COFA);
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_BUY_COST", fourniture.PRIXART);
                            Fourniture_Entity.Save();
                        }
                        //update
                        else
                        {
                            IEntity Fourniture_Entity;
                            //recupere l'id et mets a jour les champs
                            //FirstOrDefault(x => x.Value == "one").Key; 
                            long uid = Founiture_Divers_Exclusion.FirstOrDefault(x => x.Value == fourniture.COARTI).Key;
                            Fourniture_Entity = contextlocal.EntityManager.GetEntity(uid, "_SUPPLY");
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_REFERENCE", fourniture.COARTI);
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_DESIGNATION", fourniture.DESA1);
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_COMMENTS", fourniture.COFA);
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_BUY_COST", fourniture.PRIXART);
                            Fourniture_Entity.Save();

                        }
                    }


                }



                contextlocal.TraceLogger.TraceInformation("Import des fornitures terminé avec succes");



            }


            catch
            {

            }

        }


        /// <summary>
        /// ecriture dans la base almacam
        /// </summary>
        /// 
        //ecriturze des vis

        public override void Update()
        {
            try
            {
                if (this.List_Fourniture.Any())
                {

                    foreach (Founiture_Divers fourniture in this.List_Fourniture)
                    {
                        IEntity Fourniture_Entity;


                        Fourniture_Entity = contextlocal.EntityManager.CreateEntity("_SUPPLY");
                        Fourniture_Entity.SetFieldValue("_REFERENCE", fourniture.Name);
                        Fourniture_Entity.SetFieldValue("_COMMENTS", "");

                        Fourniture_Entity.Save();
                    }


                }

            }


            catch
            {

            }

        }




    }

    #endregion



}

#endregion

#region Exception
public class MissingJsonFile : Exception
{

    public void MissingJsonFileException(string parametername)
    {
        {

            MessageBox.Show("Il manque le fichier Json " + parametername + " dans le repertoire almacam");


        }

    }

}

[Serializable]
internal class Missing_Obdc_Exception : Exception

{



    public Missing_Obdc_Exception(string nomodbc)

    {

        MessageBox.Show("la connexion odbc " + nomodbc + " n'est pas configurée. Contactez Clipper pour ce point ou bien referez vous a la documentation d'installaitons");
        Environment.Exit(0);

    }




}



#endregion

#region Tools
public static class AlmaCamTool
{

    /// <summary>
    /// recupere un context almacam
    /// </summary>
    /// <param name="context"></param>
    /// <returns></returns>
    public static IContext GetContext(IContext context)
    {
        try
        {

            if (context == null)
            {



                string DbName = Alma_RegitryInfos.GetLastDataBase();
                IModelsRepository modelsRepository = new ModelsRepository();
                IContext contextlocal = modelsRepository.GetModelContext(DbName);
                return contextlocal;

            }
            else
            {
                return context;
            }
        }
        catch (Exception ie)
        {

            string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
            MessageBox.Show(ie.Message);

            return null;
        }
    }

    /// <summary>
    /// verifie si la connexion odbc est possible
    /// </summary>
    /// <param name="DSN"></param>
    /// <returns></returns>
    public static bool Is_Odbc_Exists(string DSN)
    {
        bool rst = true;
        var mDbConnection = new OdbcConnection("DSN=" + DSN);

        try
        {




            using (mDbConnection)
            {
                mDbConnection.Open();

                rst = true;
            }


            return rst;


        }
        catch (Exception ie)
        {
            MessageBox.Show("Erreur connectin odbc :" + ie.Message);
            return false;
        }
        finally { mDbConnection.Close(); }
    }
}
#endregion