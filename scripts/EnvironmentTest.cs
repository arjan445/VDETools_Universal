using Eplan.EplApi.Scripting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace VDETools.scripts
{
    public class EnvironmentTest
    {
        [DeclareAction("SetVariable")]
        public void SetVariable()
        {
            Environment.SetEnvironmentVariable("VDE_EPLAN_ExportLocation", @"C:\Temp", EnvironmentVariableTarget.Machine);
        }

        [DeclareAction("ReadVariable")]
        public void ReadVariable()
        {
            string test =Environment.GetEnvironmentVariable("VDE_EPLAN_ExportLocation", EnvironmentVariableTarget.Machine);
            MessageBox.Show(test);
        }
    }
}
