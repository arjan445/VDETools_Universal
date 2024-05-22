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
    internal class Layers
    {
        [DeclareAction("AddLayers")]
        public void AddLayers()
        {
            string VDELayers = @"C:\Users\arjan02\Source\Repos\VDETools_Universal\statics\Instellingen\Algemeen\Layerbeheer\VDELayers.elc";
            CommandLineInterpreter aEx = new CommandLineInterpreter();

            string projectpad = PathMap.SubstitutePath("$(PROJECTPATH)");
            if (projectpad == "")
            {
                MessageBox.Show("Geen project geselecteerd!");
                return;
            }

            ActionCallingContext aLayer = new ActionCallingContext();
            aLayer.AddParameter("TYPE", "IMPORT");
            aLayer.AddParameter("PROJECTNAME", projectpad);
            aLayer.AddParameter("IMPORTFILE", VDELayers);
            aEx.Execute("GraphicalLayerTable", aLayer);
        }

        [DeclareAction("WerkplaatsLayer")]
        public void WerkplaatsLayer(string activate)
        {
            CommandLineInterpreter aEx = new CommandLineInterpreter();
            string projectpad = PathMap.SubstitutePath("$(PROJECTPATH)");

            string value = "0";
            if (activate == "1")
            {
                value = "1";
            }

            ActionCallingContext aLayer = new ActionCallingContext();
            aLayer.AddParameter("PROJECTNAME", projectpad);
            aLayer.AddParameter("LAYER", "VDE_Werkplaats");
            aLayer.AddParameter("PRINTED", value);
            aLayer.AddParameter("VISIBLE", value);
            aEx.Execute("changelayer", aLayer);
        }

        [DeclareAction("Engineering3DLayer")]
        public void Engineering3DLayer(string activate)
        {
            CommandLineInterpreter aEx = new CommandLineInterpreter();
            string projectpad = PathMap.SubstitutePath("$(PROJECTPATH)");

            string value = "1";
            if (activate == "0")
            {
                value = "0.7";
            }

            ActionCallingContext aLayer = new ActionCallingContext();
            aLayer.AddParameter("PROJECTNAME", projectpad);
            aLayer.AddParameter("LAYER", "VDE_Engineering_3D");
            aLayer.AddParameter("TRANSPARENCY", value);
            aEx.Execute("changelayer", aLayer);
        }
    }
}
