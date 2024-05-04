using Eplan.EplApi.Scripting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace VDETools.scripts
{
    public class OpenExportFolder
    {
        [DeclareAction("OpenExportFolder")]
        public void Open()
        {
            string folderpath = @"C:\Temp";
            if (Directory.Exists(folderpath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo("explorer.exe")
                { Arguments = folderpath };

                Process.Start(startInfo);
            }
            else
            {
                MessageBox.Show(string.Format("{0} Folder bestaat niet!", folderpath));
            }
        }
    }
}
