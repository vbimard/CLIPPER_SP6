using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AF_Import_ODBC_Clipper_AlmaCam;
using AF_ImportTools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Wpm.Implement.Manager;



namespace ConsoleApp2
{

   

    class Program
    {
        static void Main(string[] args)
        {
            bool import_matiere = true;
            bool tube_rond = true;
            bool rond = false;
            bool tube_rectangle = false;
            bool tube_carre = false;
            bool tube_flat = false;
            bool tube_speciaux =false;
            //bool Vis = false;
            bool divers = false;

            
            IContext contextlocal = null;
            contextlocal=AlmaCamTool.GetContext(contextlocal);
            Clipper8_ImportTubes_Processor tubes  = new Clipper8_ImportTubes_Processor(import_matiere, tube_rond, rond, tube_rectangle, tube_carre, tube_flat, tube_speciaux, divers);
            tubes.Import(contextlocal);
            

            //public Clipper_ImportTubes_Processor(bool import_matiere, bool tube_rond, bool rond, bool tube_rectangle, bool tube_carre, bool tube_flat, bool tube_speciaux, bool fourniture)


            /* 
            Clipper_Import_Matiere matiere = new Clipper_Import_Matiere(contextlocal);
            matiere.Import();
             */



            /*
            Clipper_Import_Toles_Processor Toles = new Clipper_Import_Toles_Processor();
            Toles.Import(contextlocal);
            */

            //Toles.Import(contextlocal);
            //test import matiere

            /*
            Clipper_Import_Matiere matiere = new Clipper_Import_Matiere();
            matiere.Import(contextlocal);
            */
            //
            /*
            Clipper_Import_Fournitures_Divers_Processor FournitureImporter = new Clipper_Import_Fournitures_Divers_Processor();
            FournitureImporter.Import(contextlocal);
            */


        }
    }
}
