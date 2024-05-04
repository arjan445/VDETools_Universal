using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using System.Xml;

namespace VDETools
{
    public class ExportPdf
    {
        [DeclareAction("ExportPdf")]
        public void ExportPDF(string filter)
        { 
            CommandLineInterpreter aEx = new CommandLineInterpreter();
            string projectpad = PathMap.SubstitutePath("$(PROJECTPATH)");

            if (projectpad == "")
            {
                MessageBox.Show("Geen project geselecteerd!");
                return;
            }

            DialogResult result = MessageBox.Show("Alle verwerkingen actualiseren?", "Actualiseren", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                aEx.Execute("UpdatePages");
            }

            Progress progress = new Progress("SimpleProgress");
            progress.SetTitle("PDF aanmaken");
            progress.ShowImmediately();
            progress.BeginPart(100, "%");

            string shortname = "";
            string pagefilter = "";

            switch (filter)
            {
                case "Klantversie":
                    pagefilter = "Tekeningen tbv klant";
                    shortname = "KV";
                    break;
                case "Buitenmontage":
                    pagefilter = "Tekeningen tbv buitenmontage";
                    shortname = "BM";
                    break;
                case "CNC":
                    pagefilter = "Tekeningen tbv CNC";
                    shortname = "CNC";
                    break;
                case "Paneelbouw":
                    pagefilter = "Tekeningen tbv paneelbouw";
                    shortname = "PB";
                    break;
            }

            string projectname = PathMap.SubstitutePath("$(PROJECTNAME)");
            string location = @"C:\Temp_eplan"; //TODO retrieve this value from ENVVAR or User settings
            string exportlocatie = location + @"\" + projectname + "_" + shortname + ".pdf";

            //Layer519 op niet printen zetten tbv commentaren
            ActionCallingContext aLayer = new ActionCallingContext();
            aLayer.AddParameter("PROJECTNAME", projectpad);
            aLayer.AddParameter("LAYER", "EPLAN519");
            aLayer.AddParameter("PRINTED", "0");
            aEx.Execute("changelayer", aLayer);

            //Daadwerkelijke export aanmaken
            ActionCallingContext aPDF = new ActionCallingContext();
            aPDF.AddParameter("TYPE", "PDFPROJECTSCHEME");
            aPDF.AddParameter("PROJECTNAME", projectpad);
            aPDF.AddParameter("EXPORTFILE", exportlocatie);
            aPDF.AddParameter("EXPORTSCHEME", pagefilter);
            aPDF.AddParameter("BLACKWHITE", "1"); // PDF Print as Black & White
            bool sRet = aEx.Execute("export", aPDF);

            progress.EndPart(true);
            if (!sRet)
            {
                MessageBox.Show("Export niet gemaakt!\nZijn de instellingen geladen?");
                return;
            }

            result = MessageBox.Show("Export naar gemaakt naar \nC:/Temp_EPLAN/\nDe map openen?", "Export succesvol", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                new CommandLineInterpreter().Execute("OpenExportFolder", new ActionCallingContext());
            }
        }
    }
}