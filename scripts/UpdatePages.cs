using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Serialization;

namespace VDETools
{
    public class UpdatePages
    {
        [DeclareAction("UpdatePages")]
        public void Updatepage()
        {
            string projectpath = PathMap.SubstitutePath("$(PROJECTPATH)");
            if (projectpath == "")
            {
                MessageBox.Show("Geen project geselecteerd!");
                return;
            }

            Progress progress = new Progress("SimpleProgress");
            progress.SetTitle("Pagina's actualiseren");
            progress.ShowImmediately();


            if (!progress.Canceled())
            {
                CommandLineInterpreter aEx = new CommandLineInterpreter();
                aEx.Execute("XFgUpdateEvaluationAction");
            }

            progress.EndPart(true);
        }
    }
}
