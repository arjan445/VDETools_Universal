// InsertComment, Version 1.2.0, vom 16.04.2013
//
// Erweitert Eplan Electric P8 um die Möglichkeit Kommentare einzufügen,
// diese können dann mit dem Kommentare-Navigator verwaltet werden.
//
// Copyright by Frank Schöneck, 2013
// letzte Änderung: Frank Schöneck, 28.02.2013 V1.0.0, Projektbeginn
//					Frank Schöneck, 01.03.2013 V1.1.0, Ebene, Linientyp und Musterlänge als Variable eingesetzt
//					Frank Schöneck, 16.04.2013 V1.2.0, Neuer Reiter "Einstellungen" mit der Möglichkeit zum gruppieren,
//                                                     Name geändert von "InsertPDFComment" in "InsertComment"
//
// für Eplan Electric P8, ab V2.2
//
using System.Windows.Forms;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using System;
using System.Xml;
using System.Security.Cryptography.X509Certificates;
using System.Security.Authentication.ExtendedProtection;
using Settings = Eplan.EplApi.Base.Settings;
using System.Text;
using System.Collections.Generic;
using System.IO;

public partial class VdeSettingsGui : System.Windows.Forms.Form
{
    private Button btnOK;
    private Button btnAbbrechen;
    private TabControl tabControl1;
    private Label label1;
    private TextBox Scriptlocation;
    private TabPage tabKommentar;
    private ComboBox Branch;
    private Label label2;
    private TextBox Initialised;
    private TextBox Version;
    private TextBox ExportLocation;
    private TextBox EplanLocation;
    private Label label6;
    private Label label5;
    private Label label4;
    private Label label3;
    private TextBox EplanVersion;
    private Label label7;
    private Button Import;

    static string SettingName = "USER.SCRIPTS.VDE";
    private TabPage Scripts;
    private DataGridView ListOfScripts;
    #region Vom Windows Form-Designer generierter Code

    /// <summary>
    /// Erforderliche Designervariable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Verwendete Ressourcen bereinigen.
    /// </summary>
    /// <param name="disposing">True, wenn verwaltete Ressourcen
    /// gelöscht werden sollen; andernfalls False.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    /// <summary>
    /// Erforderliche Methode für die Designerunterstützung.
    /// Der Inhalt der Methode darf nicht mit dem Code-Editor
    /// geändert werden.
    /// </summary>
    private void InitializeComponent()
    {
            this.btnOK = new System.Windows.Forms.Button();
            this.btnAbbrechen = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabKommentar = new System.Windows.Forms.TabPage();
            this.EplanVersion = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.Initialised = new System.Windows.Forms.TextBox();
            this.Version = new System.Windows.Forms.TextBox();
            this.ExportLocation = new System.Windows.Forms.TextBox();
            this.EplanLocation = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Branch = new System.Windows.Forms.ComboBox();
            this.Scriptlocation = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.Scripts = new System.Windows.Forms.TabPage();
            this.ListOfScripts = new System.Windows.Forms.DataGridView();
            this.Import = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabKommentar.SuspendLayout();
            this.Scripts.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ListOfScripts)).BeginInit();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnOK.Location = new System.Drawing.Point(199, 382);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(110, 25);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnAbbrechen
            // 
            this.btnAbbrechen.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAbbrechen.Location = new System.Drawing.Point(328, 382);
            this.btnAbbrechen.Name = "btnAbbrechen";
            this.btnAbbrechen.Size = new System.Drawing.Size(110, 25);
            this.btnAbbrechen.TabIndex = 1;
            this.btnAbbrechen.Text = "Annuleren";
            this.btnAbbrechen.UseVisualStyleBackColor = true;
            this.btnAbbrechen.Click += new System.EventHandler(this.btnAbbrechen_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabKommentar);
            this.tabControl1.Controls.Add(this.Scripts);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(426, 357);
            this.tabControl1.TabIndex = 2;
            this.tabControl1.TabStop = false;
            // 
            // tabKommentar
            // 
            this.tabKommentar.BackColor = System.Drawing.Color.Transparent;
            this.tabKommentar.Controls.Add(this.EplanVersion);
            this.tabKommentar.Controls.Add(this.label7);
            this.tabKommentar.Controls.Add(this.Initialised);
            this.tabKommentar.Controls.Add(this.Version);
            this.tabKommentar.Controls.Add(this.ExportLocation);
            this.tabKommentar.Controls.Add(this.EplanLocation);
            this.tabKommentar.Controls.Add(this.label6);
            this.tabKommentar.Controls.Add(this.label5);
            this.tabKommentar.Controls.Add(this.label4);
            this.tabKommentar.Controls.Add(this.label3);
            this.tabKommentar.Controls.Add(this.label2);
            this.tabKommentar.Controls.Add(this.Branch);
            this.tabKommentar.Controls.Add(this.Scriptlocation);
            this.tabKommentar.Controls.Add(this.label1);
            this.tabKommentar.Location = new System.Drawing.Point(4, 22);
            this.tabKommentar.Name = "tabKommentar";
            this.tabKommentar.Padding = new System.Windows.Forms.Padding(3);
            this.tabKommentar.Size = new System.Drawing.Size(418, 331);
            this.tabKommentar.TabIndex = 0;
            this.tabKommentar.Text = "Instellingen";
            // 
            // EplanVersion
            // 
            this.EplanVersion.Location = new System.Drawing.Point(110, 199);
            this.EplanVersion.Name = "EplanVersion";
            this.EplanVersion.ReadOnly = true;
            this.EplanVersion.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.EplanVersion.Size = new System.Drawing.Size(281, 20);
            this.EplanVersion.TabIndex = 13;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(20, 203);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 13);
            this.label7.TabIndex = 12;
            this.label7.Text = "Eplan versie";
            // 
            // Initialised
            // 
            this.Initialised.Location = new System.Drawing.Point(110, 172);
            this.Initialised.Name = "Initialised";
            this.Initialised.ReadOnly = true;
            this.Initialised.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.Initialised.Size = new System.Drawing.Size(281, 20);
            this.Initialised.TabIndex = 11;
            // 
            // Version
            // 
            this.Version.Location = new System.Drawing.Point(110, 143);
            this.Version.Name = "Version";
            this.Version.ReadOnly = true;
            this.Version.Size = new System.Drawing.Size(281, 20);
            this.Version.TabIndex = 10;
            // 
            // ExportLocation
            // 
            this.ExportLocation.Location = new System.Drawing.Point(110, 114);
            this.ExportLocation.Name = "ExportLocation";
            this.ExportLocation.Size = new System.Drawing.Size(281, 20);
            this.ExportLocation.TabIndex = 9;
            // 
            // EplanLocation
            // 
            this.EplanLocation.Location = new System.Drawing.Point(110, 85);
            this.EplanLocation.Name = "EplanLocation";
            this.EplanLocation.Size = new System.Drawing.Size(281, 20);
            this.EplanLocation.TabIndex = 8;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(20, 176);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(70, 13);
            this.label6.TabIndex = 7;
            this.label6.Text = "Geinitaliseerd";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(20, 147);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(86, 13);
            this.label5.TabIndex = 6;
            this.label5.Text = "VDETools versie";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 118);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Export locatie";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 89);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(68, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Eplan locatie";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(68, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Script locatie";
            // 
            // Branch
            // 
            this.Branch.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Branch.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.Branch.Items.AddRange(new object[] {
            "Ongedefinieerd",
            "Boekel",
            "Breda",
            "Panningen",
            "Heteren",
            "Slowakije"});
            this.Branch.Location = new System.Drawing.Point(110, 27);
            this.Branch.MaxDropDownItems = 5;
            this.Branch.Name = "Branch";
            this.Branch.Size = new System.Drawing.Size(281, 21);
            this.Branch.TabIndex = 2;
            // 
            // Scriptlocation
            // 
            this.Scriptlocation.Location = new System.Drawing.Point(110, 56);
            this.Scriptlocation.Name = "Scriptlocation";
            this.Scriptlocation.Size = new System.Drawing.Size(281, 20);
            this.Scriptlocation.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(50, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Vestiging";
            // 
            // Scripts
            // 
            this.Scripts.Controls.Add(this.ListOfScripts);
            this.Scripts.Location = new System.Drawing.Point(4, 22);
            this.Scripts.Name = "Scripts";
            this.Scripts.Size = new System.Drawing.Size(418, 331);
            this.Scripts.TabIndex = 1;
            this.Scripts.Text = "Scripts";
            this.Scripts.UseVisualStyleBackColor = true;
            // 
            // ListOfScripts
            // 
            this.ListOfScripts.AllowUserToAddRows = false;
            this.ListOfScripts.AllowUserToDeleteRows = false;
            this.ListOfScripts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ListOfScripts.Location = new System.Drawing.Point(0, 0);
            this.ListOfScripts.Name = "ListOfScripts";
            this.ListOfScripts.ReadOnly = true;
            this.ListOfScripts.Size = new System.Drawing.Size(418, 331);
            this.ListOfScripts.TabIndex = 0;
            this.ListOfScripts.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // Import
            // 
            this.Import.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Import.Location = new System.Drawing.Point(12, 382);
            this.Import.Name = "Import";
            this.Import.Size = new System.Drawing.Size(145, 25);
            this.Import.TabIndex = 3;
            this.Import.Text = "Importeer instellingen";
            this.Import.UseVisualStyleBackColor = true;
            this.Import.Click += new System.EventHandler(this.Import_Click);
            // 
            // VdeSettingsGui
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnAbbrechen;
            this.ClientSize = new System.Drawing.Size(450, 419);
            this.Controls.Add(this.Import);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnAbbrechen);
            this.Controls.Add(this.btnOK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "VdeSettingsGui";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "VDETools instellingen";
            this.Load += new System.EventHandler(this.OnLoad);
            this.tabControl1.ResumeLayout(false);
            this.tabKommentar.ResumeLayout(false);
            this.tabKommentar.PerformLayout();
            this.Scripts.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ListOfScripts)).EndInit();
            this.ResumeLayout(false);

    }

    public VdeSettingsGui()
    {
        InitializeComponent();
    }

    #endregion


    //Action um die Form aufzurufen
    [DeclareAction("VdeSettingsGui")]
    public void VdeSettingsGui_Action()
    {
        VdeSettingsGui frm = new VdeSettingsGui();
        frm.ShowDialog();
        return;
    }

    //Form_Load
    private void OnLoad(object sender, System.EventArgs e)
    {
        Settings settings = new Settings();
        if (settings.ExistSetting(SettingName))
        {
            RetrieveSettings();
            EplanVersion.Text = PathMap.SubstitutePath("$(EPLAN_VERSION_SHORT)");
        }
        else
        {
            settings.AddStringSetting(SettingName, new string[] { }, new string[] { }, ISettings.CreationFlag.Insert);
            SaveSettings();
            ExportLocation.Text = @"C:\Temp_Eplan";
        }

        // Read all scripts:
        var settingsUrlScripts = "STATION.EplanEplApiScriptGui.Scripts";
        int countOfScripts = settings.GetCountOfValues(settingsUrlScripts);
        StringBuilder stringBuilder = new StringBuilder();
        List<Script> scripts = new List<Script>();
        for (int i = 0; i < countOfScripts; i++)
        {
            string path = settings.GetStringSetting(settingsUrlScripts, i);
            string filename = new FileInfo(path).Name;
            scripts.Add(new Script
            {
                Name = filename,
                Path = path,
            });
        }

        ListOfScripts.DataSource = scripts;



    }

    //Button Abbrechen Click
    private void btnAbbrechen_Click(object sender, System.EventArgs e)
    {
        Close();
    }

    //Button OK Click
    private void btnOK_Click(object sender, System.EventArgs e)
    {
        Settings settings = new Settings();

        if (Branch.Text == "Ongedefinieerd")
        {
            MessageBox.Show("Geen locatie geselecteerd, selecteer locatie");
            return;
        }
        // First check if the settings has changed!
        string oldbranch = settings.GetStringSetting(SettingName, 0);
        if (oldbranch != Branch.Text)
        {
            DialogResult result = MessageBox.Show("Locatie is veranderd, Schema's laden?", "Locatie gewijzigd", MessageBoxButtons.YesNo);

            if (result == DialogResult.OK)
            {

                SaveSettings();
                new CommandLineInterpreter().Execute("LoadSettings /location:" + Branch.Text);
                Close();
            }
        }

        SaveSettings();

        Close();

        return;

    }
    
    private void Import_Click(object sender, EventArgs e)
    {
        //Instellingen die specifiek voor locaties zijn;
        //ID 0 = Locatie (Panningen, Boekel, Breda, Noord of Slowakije)
        //ID 1 = Scriptlocatie
        //ID 2 = EPLANLocatie
        //ID 3 = Exportlocatie (PDF, Documentatie)

        XmlDocument xml = new XmlDocument();
        OpenFileDialog ofd = new OpenFileDialog();
        Settings settings = new Settings();

        ofd.DefaultExt = "xml";
        ofd.Filter = "XML Files|*.xml";
        ofd.Title = "Selecteer XML bestand";
        ofd.ValidateNames = true;

        if (ofd.ShowDialog() == DialogResult.OK)
        {
            xml.Load(ofd.FileName);


            if (settings.ExistSetting(SettingName))
            {
                //controle of de instellingen al aanwezig zijn.
            }
            else
            {
                settings.AddStringSetting(SettingName, new string[] { }, new string[] { }, ISettings.CreationFlag.Insert);
            }

            Branch.Text = xml.DocumentElement.SelectSingleNode("LOCATION").InnerText;
            Scriptlocation.Text = xml.DocumentElement.SelectSingleNode("SCRIPTLOCATION").InnerText;
            EplanLocation.Text = xml.DocumentElement.SelectSingleNode("EPLANLOCATION").InnerText;
            ExportLocation.Text = xml.DocumentElement.SelectSingleNode("EXPORTLOCATION").InnerText;
        }
    }


    #region Settingshandling
    private void RetrieveSettings()
    {
        Settings settings = new Settings();
        Branch.Text = settings.GetStringSetting(SettingName, 0);
        Scriptlocation.Text = settings.GetStringSetting(SettingName, 1);
        EplanLocation.Text = settings.GetStringSetting(SettingName, 2);
        //TODO add version and initialised!
        ExportLocation.Text = settings.GetStringSetting(SettingName, 3);
        //Version.Text = settings.GetStringSetting(SettingName, 4);
        //Initialised.Text = settings.GetStringSetting(SettingName, 5);
    }

    private void SaveSettings()
    {
        Settings settings = new Settings();
        settings.SetStringSetting(SettingName, Branch.Text, 0);
        settings.SetStringSetting(SettingName, Scriptlocation.Text, 1);
        settings.SetStringSetting(SettingName, EplanLocation.Text, 2);
        settings.SetStringSetting(SettingName, ExportLocation.Text, 3);
    }



    #endregion

    private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
    {

    }
}

public class Script
{
    public string Name { get; set; }
    public string Path { get; set; }
}