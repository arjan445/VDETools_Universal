using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Gui;
using Eplan.EplApi.Scripting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace VDETools
{
    public class Eplan_2024
    {
        [DeclareMenu]
        public void MenuFunction()
        {
            Toolbars toolbars = new Toolbars();
            RibbonBar RibbonBar = new RibbonBar();

            toolbars.VDEToolbar(RibbonBar);

            toolbars.ProPanelToolbar(RibbonBar);

            toolbars.DebugToolbar(RibbonBar);

            toolbars.KomaxToolbar(RibbonBar);
        }


        /// <summary>
        /// The code below is the same for every version of Eplan!
        /// </summary>
        #region action at load in
        [DeclareRegister]
        public void Register()
        {
            ScriptHandler();
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
        /// <summary>
        /// This script handler is the same for every version of EPLAN!
        /// </summary>
        public static void ScriptHandler()
        {
            string scriptspath = PathMap.SubstitutePath("$(MD_SCRIPTS)") + @"\VDE_SYNC\#VDE\VDETools\scripts";

            var files = Directory.EnumerateFiles(scriptspath, "*.cs");

            foreach (var file in files)
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

        #region Reload script function
        [DeclareAction("ReloadScripts")]
        public static void ReloadScripts()
        {
            LoadScript(@"C:\Users\arjan02\Source\Repos\VDETools_Universal\Eplan 2024.cs");
            ScriptHandler();
            MessageBox.Show("Reloaded scripts");
        }
        #endregion
    }

    public class Toolbars
    {
        private MultiLangString MultiLangVDETools
        {
            get
            {
                MultiLangString tabName = new MultiLangString();
                tabName.AddString(ISOCode.Language.L_de_DE, "VDETools");
                tabName.AddString(ISOCode.Language.L_en_US, "VDETools");
                tabName.AddString(ISOCode.Language.L_nl_NL, "VDETools");
                return tabName;
            }
        }
        private MultiLangString MultiLangPropanel
        {
            get
            {
                MultiLangString tabName = new MultiLangString();
                tabName.AddString(ISOCode.Language.L_de_DE, "Propanel");
                tabName.AddString(ISOCode.Language.L_en_US, "Propanel");
                tabName.AddString(ISOCode.Language.L_nl_NL, "Propanel");
                return tabName;
            }
        }
        private MultiLangString MultiLangDebug
        {
            get
            {
                MultiLangString tabName = new MultiLangString();
                tabName.AddString(ISOCode.Language.L_de_DE, "Debug");
                tabName.AddString(ISOCode.Language.L_en_US, "Debug");
                tabName.AddString(ISOCode.Language.L_nl_NL, "Debug");
                return tabName;
            }
        }

        private MultiLangString MultiLangKomax
        {
            get
            {
                MultiLangString tabName = new MultiLangString();
                tabName.AddString(ISOCode.Language.L_de_DE, "Komax");
                tabName.AddString(ISOCode.Language.L_en_US, "Komax");
                tabName.AddString(ISOCode.Language.L_nl_NL, "Komax");
                return tabName;
            }
        }

        public void VDEToolbar(RibbonBar Ribbonbar)
        {
            // Create the tab VDETools if not exists.
            RibbonTab VDEToolsTab = Ribbonbar.GetTab(MultiLangVDETools, true);
            if (VDEToolsTab != null)
            {
                VDEToolsTab.Remove();
            }
            VDEToolsTab = Ribbonbar.AddTab(MultiLangVDETools);

            // Tools - VDETools menu
            RibbonCommandGroup GroupTools = VDEToolsTab.AddCommandGroup("Tools");
            GroupTools.AddCommand("Open Sharepoint", "OpenSharepoint");
            GroupTools.AddCommand("Open Export map", "OpenExportFolder");
            GroupTools.AddCommand("", " ");
            GroupTools.AddCommand("Project actualiseren", "UpdatePages");
            GroupTools.AddCommand("QR instellen", "SetQr");

            // Commentaar - VDETools menu
            RibbonCommandGroup Commentaar = VDEToolsTab.AddCommandGroup("Commentaar");
            Commentaar.AddCommand("Navigator", "XPdfCommentDlg");
            Commentaar.AddCommand("Invoegen", "InsertComment");

            // Commentaar - Snelkoppelingen
            RibbonCommandGroup Snelkoppelingen = VDEToolsTab.AddCommandGroup("Snelkoppelingen");
            Snelkoppelingen.AddCommand("Verbindingen actualiseren", "EsGenerateConnections");
            Snelkoppelingen.AddCommand("Ontwerpmodus", "XGedActionToggleConstructionMode");
            Snelkoppelingen.AddCommand("Apparaat invoegen (oud)", "XDLInsertDeviceAction");
            Snelkoppelingen.AddCommand("Insert symbool (old)", "XEGActionInsertSymRef");

            // PDF Export - VDETools menu
            RibbonCommandGroup GroupPDF = VDEToolsTab.AddCommandGroup("PDF");
            GroupPDF.AddCommand("Werkplaats", "ExportPDF /filter:\"Paneelbouw\"");
            GroupPDF.AddCommand("CNC", "ExportPDF /filter:\"CNC\"");
            GroupPDF.AddCommand("Buitenmontage", "ExportPDF /filter:\"Buitenmontage\"");
            GroupPDF.AddCommand("Klant", "ExportPDF /filter:\"Klantversie\"");
            GroupPDF.AddCommand("Volledig", "ExportPDF /filter:\"Volledig\"");

            // Genereren - VDETools menu
            RibbonCommandGroup GroupGenerate = VDEToolsTab.AddCommandGroup("Genereren");
            GroupGenerate.AddCommand("Documentatie", "ExportDocumentation");

            // Verwerkingsstappen - VDETools menu
            RibbonCommandGroup GroupVerwerking = VDEToolsTab.AddCommandGroup("Project verwerken");
            GroupVerwerking.AddCommand("1. Projecteigenschappen", "GfDlgMgrActionIGfWind /function:Copy");
            GroupVerwerking.AddCommand("2. Verwerkingen genereren", "XFgFormGenMain");
            GroupVerwerking.AddCommand("3. Modelaanzichten actualiseren", "XGedUpdateViewPlacementsAction");
            GroupVerwerking.AddCommand("4. 2D-Booraanzichten (verwerkingen)", "XGedUpdateDrillViewsAction");
            GroupVerwerking.AddCommand("5. Actualiseren", "Actualiseren");

            // Revisiebeheer - VDETools menu
            // RibbonCommandGroup GroupRevision = VDEToolsTab.AddCommandGroup("Revisiebeheer");
            // GroupRevision.AddCommand("Pagina revisie bewerken", "Revision");

            // Initiele instellingen
            RibbonCommandGroup GroupSetup = VDEToolsTab.AddCommandGroup("Instellingen");
            GroupSetup.AddCommand("Instellingen", "VdeSettingsGui");

        }

        public void ProPanelToolbar(RibbonBar Ribbonbar)
        {
            RibbonTab ProPanelTab = Ribbonbar.GetTab(MultiLangPropanel, true);
            if (ProPanelTab != null)
            {
                ProPanelTab.Remove();
            }
            ProPanelTab = Ribbonbar.AddTab(MultiLangPropanel);

            // Opties
            RibbonCommandGroup OptiesTools = ProPanelTab.AddCommandGroup("Opties");
            OptiesTools.AddCommand("Lengte wijzigen", "XCabChangeLengthAction /Cursor:CHANGELENGTH3D");
            OptiesTools.AddCommand("Draaien om as", "XCabRotateAroundAxisAction");

            // Beeld
            RibbonCommandGroup BeeldTools = ProPanelTab.AddCommandGroup("Beeld");
            BeeldTools.AddCommand("Inbouwafstanden", "XCabToggleShowMountingClearance");
            BeeldTools.AddCommand("Booraanzicht", "XCabToggleShowDrilling");
            BeeldTools.AddCommand("Vereenvoudigde weergave", "XCabToggleShowSimplifiedRepresentation");

            // recht Viewpoint
            RibbonCommandGroup RechtTools = ProPanelTab.AddCommandGroup("Viewpoint recht");
            RechtTools.AddCommand("Voor", "XGedViewpointAction /Direction:Front");
            RechtTools.AddCommand("Links", "XGedViewpointAction /Direction:Left");
            RechtTools.AddCommand("Rechts", "XGedViewpointAction /Direction:Right");
            RechtTools.AddCommand("Achter", "XGedViewpointAction /Direction:Back");
            RechtTools.AddCommand("Boven", "XGedViewpointAction /Direction:Top");
            RechtTools.AddCommand("Onder", "XGedViewpointAction /Direction:Bottom");

            // 3D Viewpoint
            RibbonCommandGroup IsoTools = ProPanelTab.AddCommandGroup("Viewpoint isometrisch");
            IsoTools.AddCommand("Zuidoost", "XGedViewpointAction /Direction:Southwest");
            IsoTools.AddCommand("Zuidwest", "XGedViewpointAction /Direction:Southeast");
            IsoTools.AddCommand("Noordoost", "XGedViewpointAction /Direction:Northeast");
            IsoTools.AddCommand("Noordwest", "XGedViewpointAction /Direction:Northwest");

            // Algemeen
            RibbonCommandGroup AlgemeenTools = ProPanelTab.AddCommandGroup("Algemeen");
            AlgemeenTools.AddCommand("DIN (7,5mm)", "XDLInsertDeviceAction /PartNr:PXC.1206421 /PartVariant:1");
            AlgemeenTools.AddCommand("DIN (15mm)", "XDLInsertDeviceAction /PartNr:PXC.1206599 /PartVariant:1");
            AlgemeenTools.AddCommand("DIN (45°)", "XDLInsertDeviceAction /PartNr:PRO.WEI.0178100000 /PartVariant:1");
            AlgemeenTools.AddCommand("EMC", "XDLInsertDeviceAction /PartNr:\"PRO.PXC.EMC Rail\" /PartVariant:1");
            AlgemeenTools.AddCommand("T", "XDLInsertDeviceAction /PartNr:PXC.3026476 /PartVariant:1");
            AlgemeenTools.AddCommand("45°", "XDLInsertDeviceAction /PartNr:WEI.0178100000 /PartVariant:1");
            AlgemeenTools.AddCommand("ODC-Strip", "XDLInsertDeviceAction /PartNr:PRO.PXC.1013737 /PartVariant:1");
            AlgemeenTools.AddCommand("Dummy", "XDLInsertDeviceAction /PartNr:PRO.Dummy /PartVariant:1");
            AlgemeenTools.AddCommand("Snel Artikel invoegen", "InsertArticle");
            AlgemeenTools.AddCommand("Verzamelrailsysteem", "XGedStartInteractionAction /Name:XCabIaInsertBusBarSystem");
            AlgemeenTools.AddCommand("Vrije montageplaat", " XCabIaCreateMountingPlate");
            AlgemeenTools.AddCommand("Toebehoren", "XCabSelectAdditionalParts");

            // Uitsparingen
            RibbonCommandGroup UitsparingTools = ProPanelTab.AddCommandGroup("Uitsparingen");
            UitsparingTools.AddCommand("Boring", "XCabIaInsertDrillRound");
            UitsparingTools.AddCommand("Schroefdraad", "XCabIaInsertDrillThread");
            UitsparingTools.AddCommand("Rechthoek", "XCabIaInsertDrillRect");
            UitsparingTools.AddCommand("Slobgat", "XCabIaInsertDrillSlot");
            UitsparingTools.AddCommand("Boorpatroon", "XCabIaInsertDrillConstruction");
            UitsparingTools.AddCommand("NIET BOREN", "XCabIaInsertDrillingKeepOutArea");

            // Aardrails
            RibbonCommandGroup AardrailTools = ProPanelTab.AddCommandGroup("Aardrails");
            AardrailTools.AddCommand("61,5MM", "XDLInsertDeviceAction /PartNr:WOHN.01926 /PartVariant:1");
            AardrailTools.AddCommand("187MM", "XDLInsertDeviceAction /PartNr:WOHN.01928 /PartVariant:1");
            AardrailTools.AddCommand("250MM", "XDLInsertDeviceAction /PartNr:WOHN.01929 /PartVariant:1");
            AardrailTools.AddCommand("310MM", "XDLInsertDeviceAction /PartNr:WOHN.01930 /PartVariant:1");
            AardrailTools.AddCommand("100A (M8)", "XDLInsertDeviceAction /PartNr:WOHN.01932 /PartVariant:1");
            AardrailTools.AddCommand("200A (M10)", "XDLInsertDeviceAction /PartNr:DIJ.01093 /PartVariant:1");

            // OBO Goten
            RibbonCommandGroup OBOTools = ProPanelTab.AddCommandGroup("OBO Goten");
            OBOTools.AddCommand("37,5x75MM", "XDLInsertDeviceAction /PartNr:\"OBO.LKV N 75037\" /PartVariant:1");
            OBOTools.AddCommand("50x75MM", "XDLInsertDeviceAction /PartNr:\"OBO.LKV N 75050\" /PartVariant:1");
            OBOTools.AddCommand("75x75MM", "XDLInsertDeviceAction /PartNr:\"OBO.LKV N 75075\" /PartVariant:1");
            OBOTools.AddCommand("100x75MM", "XDLInsertDeviceAction /PartNr:\"OBO.LKV N 75100\" /PartVariant:1");
            OBOTools.AddCommand("125x75MM", "XDLInsertDeviceAction /PartNr:\"OBO.LKV N 75125\" /PartVariant:1");

            // Wartels (WISKA)
            RibbonCommandGroup WiskaTools = ProPanelTab.AddCommandGroup("Wartels (WISKA)");
            WiskaTools.AddCommand("M12", "XDLInsertDeviceAction /PartNr:\"WISK.ESKV 12\" /PartVariant:1");
            WiskaTools.AddCommand("M16", "XDLInsertDeviceAction /PartNr:\"WISK.ESKV 16\" /PartVariant:1");
            WiskaTools.AddCommand("M20", "XDLInsertDeviceAction /PartNr:\"WISK.ESKV 20\" /PartVariant:1");
            WiskaTools.AddCommand("M25", "XDLInsertDeviceAction /PartNr:\"WISK.ESKV 25\" /PartVariant:1");
            WiskaTools.AddCommand("M32", "XDLInsertDeviceAction /PartNr:\"WISK.ESKV 32\" /PartVariant:1");
            WiskaTools.AddCommand("M40", "XDLInsertDeviceAction /PartNr:\"WISK.ESKV 40\" /PartVariant:1");
            WiskaTools.AddCommand("M50", "XDLInsertDeviceAction /PartNr:\"WISK.ESKV 50\" /PartVariant:1");
            WiskaTools.AddCommand("M63", "XDLInsertDeviceAction /PartNr:\"WISK.ESKV 63\" /PartVariant:1");
        }

        public void DebugToolbar(RibbonBar Ribbonbar)
        {
            RibbonTab DebugTab = Ribbonbar.GetTab(MultiLangDebug, true);
            if (DebugTab != null)
            {
                DebugTab.Remove();
            }
            DebugTab = Ribbonbar.AddTab(MultiLangDebug);

            RibbonCommandGroup Debug = DebugTab.AddCommandGroup("Debug");
            Debug.AddCommand("Herlaad script", "ReloadScripts");
            Debug.AddCommand("TestScript", "TestScript");
            
        }

        public void KomaxToolbar(RibbonBar Ribbonbar)
        {
            RibbonTab KomaxTab = Ribbonbar.GetTab(MultiLangDebug, true);
            if (KomaxTab != null)
            {
                KomaxTab.Remove();
            }
            KomaxTab = Ribbonbar.AddTab(MultiLangKomax);

            RibbonCommandGroup Orientation = KomaxTab.AddCommandGroup("Orientation");
            Orientation.AddCommand("Vertical", "XEsSetPropertyAction /PropertyId:19307 /PropertyValue:\"Verbindingen - verticaal\"");
            Orientation.AddCommand("Horizontal", "XEsSetPropertyAction /PropertyId:19307 /PropertyValue:\"Verbindingen - horizontaal\"");

            RibbonCommandGroup Size = KomaxTab.AddCommandGroup("Size");
            Size.AddCommand("0.50 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"0,5\"");
            Size.AddCommand("0.75 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"0,75\"");
            Size.AddCommand("1.00 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"1\"");
            Size.AddCommand("1,50 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"1,5\"");
            Size.AddCommand("2.50 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"2,5\"");
            Size.AddCommand("4.00 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"4\"");
            Size.AddCommand("6.00 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"6\"");
            Size.AddCommand("10.0 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"10\"");
            Size.AddCommand("16.0 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"16\"");
            Size.AddCommand("25.0 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"25\"");
            Size.AddCommand("35.0 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"35\"");
            Size.AddCommand("50.0 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"50\"");
            Size.AddCommand("70.0 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"70\"");
            Size.AddCommand("95.0 mm²", "XEsSetPropertyAction /PropertyId:31002 /PropertyIndex:0 /PropertyValue:\"95\"");

            RibbonCommandGroup Color = KomaxTab.AddCommandGroup("Color");
            Color.AddCommand("Black", "XEsSetPropertyAction /PropertyId:31004 /PropertyIndex:0 /PropertyValue:\"BK\"");
            Color.AddCommand("Lightblue", "XEsSetPropertyAction /PropertyId:31004 /PropertyIndex:0 /PropertyValue:\"LBU\"");
            Color.AddCommand("Green/Yellow", "XEsSetPropertyAction /PropertyId:31004 /PropertyIndex:0 /PropertyValue:\"GNYE\"");
            Color.AddCommand("Blue", "XEsSetPropertyAction /PropertyId:31004 /PropertyIndex:0 /PropertyValue:\"BU\"");
            Color.AddCommand("Red", "XEsSetPropertyAction /PropertyId:31004 /PropertyIndex:0 /PropertyValue:\"RD\"");
            Color.AddCommand("Orange", "XEsSetPropertyAction /PropertyId:31004 /PropertyIndex:0 /PropertyValue:\"OR\"");
            Color.AddCommand("Grey", "XEsSetPropertyAction /PropertyId:31004 /PropertyIndex:0 /PropertyValue:\"GY\"");
            Color.AddCommand("White", "XEsSetPropertyAction /PropertyId:31004 /PropertyIndex:0 /PropertyValue:\"WH\"");
            Color.AddCommand("Brown", "XEsSetPropertyAction /PropertyId:31004 /PropertyIndex:0 /PropertyValue:\"BN\"");
            Color.AddCommand("Violet", "XEsSetPropertyAction /PropertyId:31004 /PropertyIndex:0 /PropertyValue:\"VT\"");
            Color.AddCommand("White/Blue", "XEsSetPropertyAction /PropertyId:31004 /PropertyIndex:0 /PropertyValue:\"WHDBU\"");

        }
    }
}