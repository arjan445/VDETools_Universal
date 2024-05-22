using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VDETools
{
    public class ExportDocumentation
    {
        [DeclareAction("ExportDocumentation")]
        public void export()
        {
            DialogResult result = MessageBox.Show("Het volgende wordt geexporteerd:\nODC's\nGraveerplaatjes\nKabellabels\nKabellijst\nWarmteverliesvermogen\n\nDoorgaan?", "Paneelbouw exporteren", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                Settings settings = new Settings();
                CommandLineInterpreter aEx = new CommandLineInterpreter();

                string projectpath = PathMap.SubstitutePath("$(PROJECTPATH)");
                if (projectpath == "")
                {
                    MessageBox.Show("Geen project geselecteerd!");
                    return;
                }

                Progress progress = new Progress("SimpleProgress");
                progress.SetTitle("Documentatie aan het exporteren");
                progress.ShowImmediately();
                progress.BeginPart(5, "");

                string projectnaam = PathMap.SubstitutePath("$(PROJECTNAME)");
                string exportlocatie = settings.GetStringSetting("USER.SCRIPTS.VDE",3) + @"\" + projectnaam;

                //ODC exporteren
                ActionCallingContext aODC = new ActionCallingContext();
                aODC.AddParameter("PROJECTNAME", projectpath);
                aODC.AddParameter("CONFIGSCHEME", "Onderdeelcoderingen_VDE");
                aODC.AddParameter("LANGUAGE", "??_??");
                aODC.AddParameter("DESTINATIONFILE", exportlocatie + @"\" + projectnaam + "_ODC.xlsx");
                aODC.AddParameter("SHOWOUTPUT", "0");
                aODC.AddParameter("USESELECTION", "0");
                aEx.Execute("label", aODC);

                //Graveerplaatjes exporteren
                ActionCallingContext aGraveer = new ActionCallingContext();
                aGraveer.AddParameter("PROJECTNAME", projectpath);
                aGraveer.AddParameter("CONFIGSCHEME", "Graveerplaatjes VDE");
                aGraveer.AddParameter("LANGUAGE", "??_??");
                aGraveer.AddParameter("DESTINATIONFILE", exportlocatie + @"\" + projectnaam + "_Graveer.xlsx");
                aGraveer.AddParameter("SHOWOUTPUT", "0");
                aGraveer.AddParameter("USESELECTION", "0");
                aEx.Execute("label", aGraveer);

                //Graveerplaatjes exporteren
                ActionCallingContext aKabel = new ActionCallingContext();
                aKabel.AddParameter("PROJECTNAME", projectpath);
                aKabel.AddParameter("CONFIGSCHEME", "Kabellabels_VDE");
                aKabel.AddParameter("LANGUAGE", "??_??");
                aKabel.AddParameter("DESTINATIONFILE", exportlocatie + @"\" + projectnaam + "_Kabellabels.xlsx");
                aKabel.AddParameter("SHOWOUTPUT", "0");
                aKabel.AddParameter("USESELECTION", "0");
                aEx.Execute("label", aKabel);

                //Kabellijst exporteren
                ActionCallingContext aLijst = new ActionCallingContext();
                aLijst.AddParameter("PROJECTNAME", projectpath);
                aLijst.AddParameter("CONFIGSCHEME", "Kabellijst VDE");
                aLijst.AddParameter("LANGUAGE", "??_??");
                aLijst.AddParameter("DESTINATIONFILE", exportlocatie + @"\" + projectnaam + "_Kabellijst.xlsx");
                aLijst.AddParameter("SHOWOUTPUT", "0");
                aLijst.AddParameter("USESELECTION", "0");
                aEx.Execute("label", aLijst);

                //Vermogensverlies exporteren
                ActionCallingContext aWarmte = new ActionCallingContext();
                aWarmte.AddParameter("PROJECTNAME", projectpath);
                aWarmte.AddParameter("CONFIGSCHEME", "Vermogenverlies VDE");
                aWarmte.AddParameter("LANGUAGE", "??_??");
                aWarmte.AddParameter("DESTINATIONFILE", exportlocatie + @"\" + projectnaam + "_Verliesvermogen.xlsx");
                aWarmte.AddParameter("SHOWOUTPUT", "0");
                aWarmte.AddParameter("USESELECTION", "0");
                aEx.Execute("label", aWarmte);

                progress.EndPart(true);

                result = MessageBox.Show("Export naar gemaakt naar \n" + exportlocatie + "\nDe map openen?", "Export succesvol", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    aEx.Execute("OpenExportFolder");
                }
            }
        }
    }
}
