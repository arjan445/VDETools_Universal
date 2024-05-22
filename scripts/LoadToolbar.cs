using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace VDETools.scripts
{
    /// <summary>
    /// This function is only for V2.9 and below!
    /// </summary>
    public class LoadToolbar
    {
        [DeclareAction("LoadToolbar")]
        public void Load(string toolbar)
        {
            string[] versions = { "2.9", "2.8", "2.7", "2.6" };
            
            string version = PathMap.SubstitutePath("$(EPLAN_VERSION_SHORT)");

            if (!versions.Contains(version))
            {
                MessageBox.Show("This version of EPLAN is not compactible with this script!");
                return;
            }

            string temp = PathMap.SubstitutePath("$(MD_SCRIPTS)") + @"\VDE_SYNC\#VDE\VDETools\statics\Instellingen\Algemeen\Menu's\" + toolbar + ".xml";
            //string temp = Functions.GetScriptLocation() + @"\Statics\Instellingen\V2.9\Algemeen\Menu's\" + toolbar + ".xml";
            CommandLineInterpreter aEx = new CommandLineInterpreter();
            ActionCallingContext aToolbar = new ActionCallingContext();
            aToolbar.AddParameter("File", temp);
            aToolbar.AddParameter("Replace", "Yes");
            aEx.Execute("MfImportToolbarAction", aToolbar);
        }
    }
}
