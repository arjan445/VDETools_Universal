using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
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
                    // TODO add right location of the files discuss this with Peter
                    string temp = @"C:\Users\arjan02\Source\Repos\VDETools_Universal\statics\Instellingen\Algemeen";
                    DirectoryInfo tempd = new DirectoryInfo(temp);

                    foreach (var file in tempd.GetFiles("*.xml", SearchOption.AllDirectories))
                    {
                        settings.ReadSettings(file.FullName);
                    }

                    progress.EndPart();

                    // locatie specifieke instellingen laden:
                    temp = @"C:\Users\arjan02\Source\Repos\VDETools_Universal\statics\Instellingen\" + location;
                    tempd = new DirectoryInfo(temp);

                    foreach (var file in tempd.GetFiles("*.xml", SearchOption.AllDirectories))
                    {
                        settings.ReadSettings(file.FullName);
                    }
                    progress.EndPart();

                    // VDEToolbar laden
                    new CommandLineInterpreter().Execute("LoadToolbar /toolbar:\"Propanel Instellingen en genereerstappen_V2\"");

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
                MessageBox.Show("Basisinstellingen niet geladen!");
            }
        }
    }
}
