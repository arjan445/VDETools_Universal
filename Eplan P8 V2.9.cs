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
        #region action at load in 
        [DeclareRegister]
        public void Register()
        {
            ScriptHandler();
        }
        #endregion

        #region action at load out
        [DeclareUnregister]
        public void UnRegister(){
        
        }
        #endregion


        #region menu handler
        [DeclareMenu]
        public void MenuFunction()
        {

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