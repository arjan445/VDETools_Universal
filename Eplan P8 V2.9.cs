using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using System.Xml;
using System.Collections.Generic;
using System.Security.Cryptography;
namespace VDETools
{
    public class Eplan_P8_V2
    {
        uint menuId = new uint();

        #region action at load in 
        [DeclareRegister]
        public void Register()
        {
            ScriptHandler();
        }
        #endregion

        #region action at load out
        [DeclareUnregister]
        public void UnRegister()
        {
            Eplan.EplApi.Gui.Menu menu = new Eplan.EplApi.Gui.Menu();
            menu.RemoveMenuItem(menuId);
        }
        #endregion


        #region menu handler
        [DeclareMenu]
        public void MenuFunction()
        {

            //AdministratorHandling.DebugMessage("NIEUWE VERSIE");
            // Main menu aanmaken
            Eplan.EplApi.Gui.Menu menu = new Eplan.EplApi.Gui.Menu();
            menuId = menu.AddMainMenu("VDE Tools", Eplan.EplApi.Gui.Menu.MainMenuName.eMainMenuUtilities, "Project actualiseren", "UpdatePages", "Statustext", 1);
            menu.AddMenuItem("Open sharepoint", "OpenSharepoint", "OpenSharepoint", menuId, 2, false, false);
            menu.AddMenuItem("Documentatie genereren", "ExportDocumentation", "ExportDocumentation", menuId, 3, false, false);
            //menu.AddMenuItem("Bestellijst exporteren (syntess)", "Syntess", "Syntess", menuId, 4, false, false);
            menu.AddMenuItem("QR instellen", "SetQr", "SetQr", menuId, 5, true, false);
            //menu.AddMenuItem("Engineer instellen", "SetEngineer", "SetEngineer", menuId, 6, false, false);
            //menu.AddMenuItem("Versie instellen", "SetVersion", "SetVersion", menuId, 7, false, false);
            //menu.AddMenuItem("Pagina revisie bewerken", "Revision", "Revision", menuId, 8, false, false);
            menu.AddMenuItem("Basisinstellingen instellen", "Basisinstellingen", "Basisinstellingen", menuId, 11, true, false);


            // instellingen menu aanmaken
            uint settingsmenu = new uint();
            settingsmenu = menu.AddPopupMenuItem("Instellingen laden", "LoadSettings /location:Panningen", "Panningen", "Selecteer optie", menuId, 10, false, false);
            menu.AddMenuItem("Boekel", "LoadSettings /location:Boekel", "Boekel", settingsmenu, 2, false, false);
            menu.AddMenuItem("Breda", "LoadSettings /location:Breda", "Breda", settingsmenu, 3, false, false);
            menu.AddMenuItem("Noord", "LoadSettings /location:Noord", "Noord", settingsmenu, 4, false, false);
            menu.AddMenuItem("Slowakije", "LoadSettings /location:Slowakije", "Slowakije", settingsmenu, 5, false, false);

            // pdf menu aanmaken
            uint pdfmenu = new uint();
            pdfmenu = menu.AddPopupMenuItem("PDF Export", "Export map openen", "OpenExportMap", "OpenExportMap", menuId, 2, true, false);
            menu.AddMenuItem("Werkplaats versie", "ExportPdf /filter:Paneelbow", "Werkplaatsversie", pdfmenu, 1, true, false);
            menu.AddMenuItem("CNC versie", "ExportPdf /filter:CNC", "CNCversie", pdfmenu, 2, false, false);
            menu.AddMenuItem("Buitenmontage versie", "ExportPdf /filter:Buitenmontage", "Buitenmontageversie", pdfmenu, 3, false, false);
            menu.AddMenuItem("Klant versie", "ExportPdf /filter:Klantversie", "Selecteer optie", pdfmenu, 4, false, false);

            // layer menu aanmaken
            uint layermenu = new uint();
            layermenu = menu.AddPopupMenuItem("Layer beheer", "Layers actualiseren", "AddLayers", "AddLayers", menuId, 2, true, false);
            menu.AddMenuItem("Werkplaats layer activeren", "WerkplaatsLayer /activate:1", "LayerWerkplaatsEnable", layermenu, 1, true, false);
            menu.AddMenuItem("Werkplaats layer deactiveren", "WerkplaatsLayer /activatre:0", "LayerWerkplaatsDisable", layermenu, 2, false, false);
            menu.AddMenuItem("3D engineering layer activeren", "Engineering3DLayer /activate:1", "Layer3DEnable", layermenu, 3, true, false);
            menu.AddMenuItem("3D engineering layer deactiveren", "Engineering3DLayer /activate:0", "Layer3DDisable", layermenu, 4, false, false);



            // Werkbalken menu aanmaken
            uint werkmenu = new uint();
            werkmenu = menu.AddPopupMenuItem("Werkbalken", "Selecteer werkbalk:", "-", "", menuId, 9, true, false);
            //menu.AddMenuItem("Propanel - Rehau 75 hoog", "Rehau75", "Rehau75", werkmenu, 1, false, false);
            menu.AddMenuItem("Propanel - OBO 75 hoog", "LoadToolbar /toolbar:\"Propanel - OBO\"", "OBO75", werkmenu, 2, false, false);
            menu.AddMenuItem("Propanel - Wiska wartels", "LoadToolbar /toolbar:\"Propanel - Wartels\"", "Wartels", werkmenu, 3, false, false);
            menu.AddMenuItem("Propanel - Aardrail", "LoadToolbar /toolbar:\"Propanel - Aardrail\"", "Aardrail", werkmenu, 4, false, false);
            menu.AddMenuItem("Propanel - Algemeen", "LoadToolbar /toolbar:\"Propanel - Algemeen\"", "Propanel", werkmenu, 5, false, false);
            menu.AddMenuItem("Algemeen - Artikel invoegen", "LoadToolbar /toolbar:\"Algemeen - Artikel invoegen\"", "ArtikelMenu", werkmenu, 6, true, false);
        }
        #endregion

        #region EventHandler on main start
        [DeclareEventHandler("onMainStart")]
        public void onMainStart()
        {
            // TODO: This works, but actually needed to be OnRegister and not on start!
            ScriptHandler();
        }
        #endregion

        #region loading scripts from directory
        public static void ScriptHandler()
        {
            //TODO Use $(MD_SCRIPTS) subsitute path
            string scriptspath = @"C:\Users\arjan02\Source\Repos\VDETools_Universal\scripts";

            var files = Directory.EnumerateFiles(scriptspath);

            foreach(var file in  files)
            {
                LoadScript(file);
            }
        }

        public static void LoadScript(string path)
        {
            //MessageBox.Show(path);
            CommandLineInterpreter aEx = new CommandLineInterpreter();
            ActionCallingContext script = new ActionCallingContext();
            script.AddParameter("ScriptFile", path);
            bool status = aEx.Execute("RegisterScript", script);
            //MessageBox.Show(status.ToString()) ;
        }
        #endregion

        [DeclareAction("ReloadScripts")]
        public static void ReloadScripts()
        {
            ScriptHandler();
            MessageBox.Show("Reloaded scripts");
        }
    }
}