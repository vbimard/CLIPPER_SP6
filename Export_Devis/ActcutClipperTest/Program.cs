using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;

namespace ActcutClipperTest
{
    static class Program
    {
        public static string AlmaCamBinFolder = null;

        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            AlmaCamBinFolder = @"C:\Developpement\Branches\AlmaCam\trunk\build\Bin\Debug";

            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TestForm());
        }

        static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            string assemblyName = new AssemblyName(args.Name).Name;

            string candidateAssemblyFile = Path.Combine(AlmaCamBinFolder, assemblyName + ".dll");
            if (File.Exists(candidateAssemblyFile))
            {
                Assembly assembly = Assembly.LoadFrom(candidateAssemblyFile);
                return assembly;
            }

            return null;
        }
    }
}
