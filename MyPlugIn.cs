#nullable enable
using System;
using System.AddIn;
using System.Reflection;
using System.Windows.Interop;
using System.Windows.Threading;
using Aucotec.EngineeringBase.Client.Runtime;
// Alias EB Application to avoid clash with System.Windows.Application
using EbApp = Aucotec.EngineeringBase.Client.Runtime.Application;

namespace JJ_Lurgi_Piping_EB
{
    [AddIn("JJ Lurgi Piping EB", Description = "", Publisher = "jjlem.dev")]
    public class MyPlugIn : PlugInWizard
    {
        public override void Run(EbApp myApplication)
        {
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                string path = asm.Location ?? string.Empty;
                string ver = asm.GetName().Version != null ? asm.GetName().Version.ToString() : "n/a";
                System.Windows.Forms.MessageBox.Show(
                    System.IO.Path.GetFileName(path) + "  |  v" + ver + "\nPath: " + path,
                    "JJ Lurgi Piping EB – Loaded",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch { /* optional */ }

            StartWindow start = new StartWindow(myApplication)
            {
                Title = "Piping Parts DS & Pipe Class Generator — Start"
            };

            try
            {
                var wih = new WindowInteropHelper(start);
                if (myApplication != null && myApplication.ActiveWindow != null)
                {
                    wih.Owner = myApplication.ActiveWindow.Handle;
                }
            }
            catch { }

            start.ShowDialog();

            if (!AppDomain.CurrentDomain.IsDefaultAppDomain())
                Dispatcher.CurrentDispatcher.InvokeShutdown();
        }
    }
}
#nullable disable
