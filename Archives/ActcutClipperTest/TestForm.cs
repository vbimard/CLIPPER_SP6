using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

using AF_Actcut.ActcutClipperApi;
using Wpm.Implement.Manager;
//using AF_Actcut.ActcutClipperApi;

namespace ActcutClipperTest
{
    public partial class TestForm : Form
    {
        private ClipperApi ClipperApi = null;
        private Process QuoteModelApiProcess = null;

        private string ParamFolder = null;
        private string ParamFile = null;
        private string ResultFile = null;

        private string AlmaCamDB = null;
        private string User = null;

        public TestForm()
        {
            InitializeComponent();

            ParamFolder = Program.AlmaCamBinFolder + @"\ActcutClipperExeParam"; // ActcutClipperExeParam : nom de sous répertoire imposé
            ParamFile = ParamFolder + @"\Param.txt"; // Nom du fichier contenant les commandes : nom imposé
            ResultFile = ParamFolder + @"\Result.txt"; // Nom du fichier contenant les resultats : nom imposé

        }

        #region Api Test
        /*
        private void BtnInitApi_Click(object sender, EventArgs e)
        {
            AlmaCamDB = comboDataBaseList.Text;///txt_Database.Text;;
            User = "SUPER";
            ClipperApi = new ClipperApi();
            ClipperApi.InitAlmaCam(AlmaCamDB, User);
        }
        private void BtnGetQuoteApi_Click(object sender, EventArgs e)
        {
            long quoteNumberReference = -1;
            ClipperApi.SelectQuoteUI(out quoteNumberReference);
        }
        private void BtnExportQuoteApi_Click(object sender, EventArgs e)
        {
            ClipperApi.ExportQuote(307, "987", @"C:\Temp\toto.txt");
        }
*/
        #endregion

        #region Exe Test
        /*
        private void BtnInit_Click(object sender, EventArgs e)
        {
            Process[] processList = Process.GetProcesses();
            foreach (Process process in processList)
            {
                if (process.ProcessName == "Actcut.ActcutClipperExe")
                {
                    QuoteModelApiProcess = process;
                    MessageBox.Show("Actcut.ActcutClipperExe deja lancé");
                    return;
                }
            }

            if (File.Exists(ResultFile)) File.Delete(ResultFile);

            QuoteModelApiProcess = new System.Diagnostics.Process();
            QuoteModelApiProcess.StartInfo.FileName = Path.Combine(Program.AlmaCamBinFolder, "Actcut.ActcutClipperExe.exe");
            QuoteModelApiProcess.StartInfo.Arguments = "Action=Init User=" + User + " Db=" + AlmaCamDB;
            QuoteModelApiProcess.Start();

            WaitResult(ResultFile);
        }
        private void BtnGetQuote_Click(object sender, EventArgs e)
        {
            if (File.Exists(ParamFile)) File.Delete(ParamFile);
            if (File.Exists(ResultFile)) File.Delete(ResultFile);

            File.WriteAllText(ParamFile, @"Action=GetQuote");
            WaitResult(ResultFile);
        }
        private void BtnExportQuote_Click(object sender, EventArgs e)
        {
            if (File.Exists(ParamFile)) File.Delete(ParamFile);
            if (File.Exists(ResultFile)) File.Delete(ResultFile);

            File.WriteAllText(ParamFile, @"Action=ExportQuote QuoteNumber=307 OrderNumber=987 ExportFile=C:\Temp\toto.txt");
            WaitResult(ResultFile);
        }
        private void BtnExit_Click(object sender, EventArgs e)
        {
            if (File.Exists(ParamFile)) File.Delete(ParamFile);

            File.WriteAllText(ParamFile, @"Action=Exit");
        }

        private void WaitResult(string resultFile)
        {
            while (File.Exists(resultFile) == false && QuoteModelApiProcess != null && QuoteModelApiProcess.HasExited == false)
                System.Threading.Thread.Sleep(10);

            if (File.Exists(resultFile))
            {
                string result = File.ReadAllText(resultFile);
                MessageBox.Show(result);
            }
            else
            {
                MessageBox.Show("Error");
            }
        }

    */
        #endregion

        private void TestForm_Load(object sender, EventArgs e)
        {
            this.Export.Enabled = false;
            this.ExportUI.Enabled = false;
            //IEntity quote = null;
            List<IModel> listModel = new List<IModel>();


            listModel = AF_ImportTools.DataBase.GetDataBaseList();

            foreach (IModel almacam_model in listModel)
            {
                comboDataBaseList.Items.Add(almacam_model.DatabaseName);
            }



        }


       

        private void linkquote_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe ", linkquote.Text);
        }

        private void linkemf_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe ", linkemf.Text);
        }

        private void Init_Click(object sender, EventArgs e)
        {
            AlmaCamDB = comboDataBaseList.Text;///txt_Database.Text;;
            User = "SUPER";
            ClipperApi = new ClipperApi();
            ClipperApi.ConnectAlmaCamDatabase(AlmaCamDB, User);

            this.Init.Enabled = false;
            this.Export.Enabled = true;
            this.ExportUI.Enabled = true;
            this.Select_Quote.Enabled = true;
        }

        private void Export_Click(object sender, EventArgs e)
        {
            long quotenumber;
            if(txtQuoteid.Text != null){
                bool parsed = Int64.TryParse(txtQuoteid.Text, out quotenumber);
                ClipperApi.ExportQuote(quotenumber, "CMD" + txtQuoteid.Text, @"c:\temp\toto.txt");
                if (!parsed)
                MessageBox.Show("Int32.TryParse could not parse '{0}' to an int.");
            }
          
               
        }
        
        private void Reset_Click(object sender, EventArgs e)
        {
            ClipperApi = null;
            ClipperApi = new ClipperApi();
            ClipperApi.ExitAlmaCam();
            this.Init.Enabled = true;
            this.Export.Enabled = false;
            this.ExportUI.Enabled = false;
            this.Select_Quote.Enabled = true;
        }

        private void ExportUI_Click(object sender, EventArgs e)
        {
            long quoteNumberReference = -1;
            ClipperApi.SelectQuoteUI(out quoteNumberReference);
            txtQuoteid.Text = quoteNumberReference.ToString();
        }

        private void Select_Quote_Click(object sender, EventArgs e)
        {
            long quotenumber;
            
                //bool parsed = Int64.TryParse(txtQuoteid.Text, out quotenumber);
                ClipperApi.SelectQuoteUI(out quotenumber);
                txtQuoteid.Text = quotenumber.ToString();



        }
    }
}
