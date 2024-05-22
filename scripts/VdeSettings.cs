using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace VDETools
{
    public class VdeSettings
    {
        [DeclareAction("LoadSettings")]
        public static void LoadSettings(string location)
        {
            Settings settings = new Settings();
            if (settings.ExistSetting("USER.SCRIPTS.VDE"))
            {
                Progress progress = new Progress("SimpleProgress");
                progress.SetTitle("Instellingen laden - " + location);
                progress.ShowImmediately();
                progress.BeginPart(5, "Instellingen laden");

                try
                {
                    // algemene instellingen + filters laden
                    string temp = PathMap.SubstitutePath("$(MD_SCRIPTS)") + @"\VDE_SYNC\#VDE\VDETools\statics\Instellingen\Algemeen";
                    DirectoryInfo tempd = new DirectoryInfo(temp);

                    foreach (var file in tempd.GetFiles("*.xml", SearchOption.AllDirectories))
                    {
                        settings.ReadSettings(file.FullName);
                    }

                    progress.EndPart();

                    // locatie specifieke instellingen laden:
                    temp = PathMap.SubstitutePath("$(MD_SCRIPTS)") + @"\VDE_SYNC\#VDE\VDETools\statics\Instellingen\" + location;
                    tempd = new DirectoryInfo(temp);

                    foreach (var file in tempd.GetFiles("*.xml", SearchOption.AllDirectories))
                    {
                        settings.ReadSettings(file.FullName);
                    }
                    progress.EndPart();

                    // Printmarges instellen
                    temp = @"C:\Users\arjan02\Source\Repos\VDETools_Universal\statics\Instellingen\Algemeen\Gebruikersinstellingen\Afdrukmargesinstellingen.xml";
                    ActionCallingContext aPrint = new ActionCallingContext();
                    CommandLineInterpreter aEx = new CommandLineInterpreter();
                    aPrint.AddParameter("XmlFile", temp);
                    aPrint.AddParameter("NODE", "STATION.Print");
                    aPrint.AddParameter("Option", "OVERWRITE");
                    aEx.Execute("XSettingsImport", aPrint);

                    // locatie specfieke artikeldatabase inladen
                    SchemeSetting oSchemeSetting = new SchemeSetting();
                    oSchemeSetting.Init("USER.PartSelectionGui.DataSourceScheme");
                    string strSchemeName = "EplanV29_VDE_WerkDB_" + location;
                    if (oSchemeSetting.CheckIfSchemeExists(strSchemeName))
                    {
                        oSchemeSetting.SetLastUsed(strSchemeName);
                    }
                    progress.EndPart();

                    // locatie specfieke directory's inladen
                    oSchemeSetting.Init("USER.ModalDialogs.PathsScheme");
                    strSchemeName = "EplanV29_VDE_WerkDB_" + location;
                    if (oSchemeSetting.CheckIfSchemeExists(strSchemeName))
                    {
                        oSchemeSetting.SetLastUsed(strSchemeName);
                    }

            
                    // locatie specfieke vertaaldatabase inladen
                    temp = @"C:\Users\arjan02\Source\Repos\VDETools_Universal\statics\Instellingen\" + location + @"\Gebruikersinstellingen\Woordenboek.xml";

                    ActionCallingContext aVertaal = new ActionCallingContext();
                    CommandLineInterpreter aExecute = new CommandLineInterpreter();
                    aVertaal.AddParameter("XmlFile", temp);
                    aVertaal.AddParameter("NODE", "USER.TRANSLATEGUI");
                    aVertaal.AddParameter("Option", "OVERWRITE");
                    bool test = aExecute.Execute("XSettingsImport", aVertaal);

                    MessageBox.Show("Gebruikersinstellingen geladen! \nHerstart EPLAN om alles definitief te maken!");
                }
                catch
                {
                    MessageBox.Show("Er ging iets fout!\nNetwerkschijven beschikbaar?");
                }
                progress.EndPart();

                progress.EndPart(true);

            }
            else
            {
                MessageBox.Show("Eerst instellingen instellen!");
            }
        }


        [DeclareAction("TestScript")]
        public static void LoadSchematic()
        {
            Settings settings = new Settings();
            MessageBox.Show("Test");
            // locatie specifieke instellingen laden:
            string temp = "C:\\Temp_EPLAN\\Test";
            var tempd = new DirectoryInfo(temp);

            foreach (var file in tempd.GetFiles("*.xml", SearchOption.AllDirectories))
            {
                MessageBox.Show(file.Name);
                settings.ReadSettings(file.FullName);
            }

            // Function for setting the overview schema to the right setting!
            // Clean it up
            SchemeSetting oSchemeSetting = new SchemeSetting();
            oSchemeSetting.Init("USER.EnfMVC.Property.GridDisplayOrder.117.ArticleOverview.-1.DisplayScheme");
            string strSchemeName = "VDE Overzicht";
            if (oSchemeSetting.CheckIfSchemeExists(strSchemeName))
            {
                oSchemeSetting.SetLastUsed(strSchemeName);
            }

            // Function for setting the article properties to the right one
            // Clean it up
            oSchemeSetting.Init("USER.EnfMVC.Property.GridDisplayOrder.117.ArticleProperties.-1.DisplayScheme");
            strSchemeName = "VDE Eigenschappen";
            if (oSchemeSetting.CheckIfSchemeExists(strSchemeName))
            {
                oSchemeSetting.SetLastUsed(strSchemeName);
            }




            MessageBox.Show(oSchemeSetting.Description.GetAsString());

            MessageBox.Show("Runned test");
        }
    }
}
