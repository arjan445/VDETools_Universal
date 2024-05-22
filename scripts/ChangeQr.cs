using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VDETools.scripts
{
    public class ChangeQr
    {
        [DeclareAction("SetQr")]
        public void SetQr()
        {
            DialogResult result = MessageBox.Show("Selecteer het PNG van bestand van de nieuwe QR in de volgende dialoog. \n\nDoorgaan?", "QR wisselen", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                CommandLineInterpreter aEx = new CommandLineInterpreter();

                string project = PathMap.SubstitutePath("$(PROJECTPATH)");
                string projectdirectory = PathMap.SubstitutePath("$(P)");


                string destination = projectdirectory + @"\Images\QR_Item.png";

                ActionCallingContext aClose = new ActionCallingContext();
                aClose.AddParameter("NOCLOSE", "0");
                aEx.Execute("XPrjActionProjectClose", aClose);

                OpenFileDialog OfdQR = new OpenFileDialog();
                OfdQR.DefaultExt = "png";
                OfdQR.Filter = "QR Bestand|*.png";
                OfdQR.Title = "Selecteer QR bestand";
                OfdQR.ValidateNames = true;

                if (OfdQR.ShowDialog() == DialogResult.OK)
                {
                    var sourceQR = OfdQR.FileName;

                    try
                    {
                        File.Delete(destination);
                        File.Copy(sourceQR, destination, true);
                        MessageBox.Show("QR gewisseld");
                        ActionCallingContext aOpen = new ActionCallingContext();
                        aOpen.AddParameter("Project", project);
                        aEx.Execute("ProjectOpen", aOpen);
                    }
                    catch (IOException Ex)
                    {
                        MessageBox.Show(Ex.InnerException.ToString());
                        MessageBox.Show(Ex.Message);
                    }
                }
            }
        }
    }
}
