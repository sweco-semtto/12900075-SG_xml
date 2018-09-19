namespace SG_xml
{
    partial class XMLParser
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(XMLParser));
            this.XMLSträng = new System.Windows.Forms.RichTextBox();
            this.LäsInXML = new System.Windows.Forms.Button();
            this.Infotext = new System.Windows.Forms.Label();
            this.Rensa = new System.Windows.Forms.Button();
            this.DatabasText = new System.Windows.Forms.Label();
            this.Sökväg = new System.Windows.Forms.TextBox();
            this.Bläddra = new System.Windows.Forms.Button();
            this.tabControlSG = new System.Windows.Forms.TabControl();
            this.tabPageXMLReader = new System.Windows.Forms.TabPage();
            this.tabPageExcelWriter = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.labelNumerOfRows = new System.Windows.Forms.Label();
            this.labelCollectedRows = new System.Windows.Forms.Label();
            this.listViewSelected = new System.Windows.Forms.ListView();
            this.buttonWriteExcelFileToDir = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonChooseExcelPath = new System.Windows.Forms.Button();
            this.textBoxExcelDirectory = new System.Windows.Forms.TextBox();
            this.labelSaveDirectory = new System.Windows.Forms.Label();
            this.tabControlSG.SuspendLayout();
            this.tabPageXMLReader.SuspendLayout();
            this.tabPageExcelWriter.SuspendLayout();
            this.SuspendLayout();
            // 
            // XMLSträng
            // 
            this.XMLSträng.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.XMLSträng.Location = new System.Drawing.Point(9, 19);
            this.XMLSträng.Name = "XMLSträng";
            this.XMLSträng.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.XMLSträng.Size = new System.Drawing.Size(485, 383);
            this.XMLSträng.TabIndex = 1;
            this.XMLSträng.Text = "";
            this.XMLSträng.TextChanged += new System.EventHandler(this.XMLSträng_TextChanged);
            // 
            // LäsInXML
            // 
            this.LäsInXML.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.LäsInXML.Location = new System.Drawing.Point(419, 408);
            this.LäsInXML.Name = "LäsInXML";
            this.LäsInXML.Size = new System.Drawing.Size(75, 23);
            this.LäsInXML.TabIndex = 2;
            this.LäsInXML.Text = "Läs in xml";
            this.LäsInXML.UseVisualStyleBackColor = true;
            this.LäsInXML.Click += new System.EventHandler(this.LäsInXML_Click);
            // 
            // Infotext
            // 
            this.Infotext.AutoSize = true;
            this.Infotext.Location = new System.Drawing.Point(6, 3);
            this.Infotext.Name = "Infotext";
            this.Infotext.Size = new System.Drawing.Size(199, 13);
            this.Infotext.TabIndex = 4;
            this.Infotext.Text = "Klistra in xml-strängen ifrån e-post nedan:";
            // 
            // Rensa
            // 
            this.Rensa.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.Rensa.Location = new System.Drawing.Point(9, 408);
            this.Rensa.Name = "Rensa";
            this.Rensa.Size = new System.Drawing.Size(75, 23);
            this.Rensa.TabIndex = 5;
            this.Rensa.Text = "Rensa text";
            this.Rensa.UseVisualStyleBackColor = true;
            this.Rensa.Click += new System.EventHandler(this.Rensa_Click);
            // 
            // DatabasText
            // 
            this.DatabasText.AutoSize = true;
            this.DatabasText.Location = new System.Drawing.Point(1, 9);
            this.DatabasText.Name = "DatabasText";
            this.DatabasText.Size = new System.Drawing.Size(181, 13);
            this.DatabasText.TabIndex = 6;
            this.DatabasText.Text = "Ange accessdatabas att arbeta med:";
            // 
            // Sökväg
            // 
            this.Sökväg.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.Sökväg.Location = new System.Drawing.Point(4, 25);
            this.Sökväg.Name = "Sökväg";
            this.Sökväg.Size = new System.Drawing.Size(418, 20);
            this.Sökväg.TabIndex = 7;
            // 
            // Bläddra
            // 
            this.Bläddra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.Bläddra.Location = new System.Drawing.Point(432, 23);
            this.Bläddra.Name = "Bläddra";
            this.Bläddra.Size = new System.Drawing.Size(75, 23);
            this.Bläddra.TabIndex = 8;
            this.Bläddra.Text = "Bläddra...";
            this.Bläddra.UseVisualStyleBackColor = true;
            this.Bläddra.Click += new System.EventHandler(this.Bläddra_Click);
            // 
            // tabControlSG
            // 
            this.tabControlSG.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControlSG.Controls.Add(this.tabPageXMLReader);
            this.tabControlSG.Controls.Add(this.tabPageExcelWriter);
            this.tabControlSG.Location = new System.Drawing.Point(4, 51);
            this.tabControlSG.Name = "tabControlSG";
            this.tabControlSG.SelectedIndex = 0;
            this.tabControlSG.Size = new System.Drawing.Size(514, 463);
            this.tabControlSG.TabIndex = 9;
            // 
            // tabPageXMLReader
            // 
            this.tabPageXMLReader.Controls.Add(this.LäsInXML);
            this.tabPageXMLReader.Controls.Add(this.Rensa);
            this.tabPageXMLReader.Controls.Add(this.Infotext);
            this.tabPageXMLReader.Controls.Add(this.XMLSträng);
            this.tabPageXMLReader.Location = new System.Drawing.Point(4, 22);
            this.tabPageXMLReader.Name = "tabPageXMLReader";
            this.tabPageXMLReader.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageXMLReader.Size = new System.Drawing.Size(506, 437);
            this.tabPageXMLReader.TabIndex = 0;
            this.tabPageXMLReader.Text = "Läs in XML till databas";
            this.tabPageXMLReader.UseVisualStyleBackColor = true;
            // 
            // tabPageExcelWriter
            // 
            this.tabPageExcelWriter.Controls.Add(this.label2);
            this.tabPageExcelWriter.Controls.Add(this.labelNumerOfRows);
            this.tabPageExcelWriter.Controls.Add(this.labelCollectedRows);
            this.tabPageExcelWriter.Controls.Add(this.listViewSelected);
            this.tabPageExcelWriter.Controls.Add(this.buttonWriteExcelFileToDir);
            this.tabPageExcelWriter.Controls.Add(this.label1);
            this.tabPageExcelWriter.Controls.Add(this.buttonChooseExcelPath);
            this.tabPageExcelWriter.Controls.Add(this.textBoxExcelDirectory);
            this.tabPageExcelWriter.Controls.Add(this.labelSaveDirectory);
            this.tabPageExcelWriter.Location = new System.Drawing.Point(4, 22);
            this.tabPageExcelWriter.Name = "tabPageExcelWriter";
            this.tabPageExcelWriter.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageExcelWriter.Size = new System.Drawing.Size(506, 437);
            this.tabPageExcelWriter.TabIndex = 1;
            this.tabPageExcelWriter.Text = "Skriv Excel från databas";
            this.tabPageExcelWriter.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(215, 78);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(16, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "till";
            // 
            // labelNumerOfRows
            // 
            this.labelNumerOfRows.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelNumerOfRows.AutoSize = true;
            this.labelNumerOfRows.Location = new System.Drawing.Point(230, 377);
            this.labelNumerOfRows.Name = "labelNumerOfRows";
            this.labelNumerOfRows.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.labelNumerOfRows.Size = new System.Drawing.Size(13, 13);
            this.labelNumerOfRows.TabIndex = 9;
            this.labelNumerOfRows.Text = "0";
            // 
            // labelCollectedRows
            // 
            this.labelCollectedRows.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelCollectedRows.AutoSize = true;
            this.labelCollectedRows.Location = new System.Drawing.Point(8, 377);
            this.labelCollectedRows.Name = "labelCollectedRows";
            this.labelCollectedRows.Size = new System.Drawing.Size(216, 13);
            this.labelCollectedRows.TabIndex = 8;
            this.labelCollectedRows.Text = "Antal hämtade ordernummer från databasen:";
            // 
            // listViewSelected
            // 
            this.listViewSelected.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewSelected.CheckBoxes = true;
            this.listViewSelected.GridLines = true;
            this.listViewSelected.Location = new System.Drawing.Point(9, 102);
            this.listViewSelected.Name = "listViewSelected";
            this.listViewSelected.Size = new System.Drawing.Size(490, 272);
            this.listViewSelected.TabIndex = 7;
            this.listViewSelected.UseCompatibleStateImageBehavior = false;
            this.listViewSelected.View = System.Windows.Forms.View.Details;
            this.listViewSelected.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.listViewSelected_ColumnClick);
            this.listViewSelected.SelectedIndexChanged += new System.EventHandler(this.listViewSelected_SelectedIndexChanged);
            this.listViewSelected.BindingContextChanged += new System.EventHandler(this.listViewSelected_BindingContextChanged);
            this.listViewSelected.TabIndexChanged += new System.EventHandler(this.listViewSelected_TabIndexChanged);
            // 
            // buttonWriteExcelFileToDir
            // 
            this.buttonWriteExcelFileToDir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonWriteExcelFileToDir.Location = new System.Drawing.Point(275, 406);
            this.buttonWriteExcelFileToDir.Name = "buttonWriteExcelFileToDir";
            this.buttonWriteExcelFileToDir.Size = new System.Drawing.Size(224, 23);
            this.buttonWriteExcelFileToDir.TabIndex = 6;
            this.buttonWriteExcelFileToDir.Text = "Skriv Excel filer till katalog utifrån valt datum";
            this.buttonWriteExcelFileToDir.UseVisualStyleBackColor = true;
            this.buttonWriteExcelFileToDir.Click += new System.EventHandler(this.buttonWriteExcelFileToDir_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(191, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Välj datum att hämta ordernummer från:";
            this.label1.Click += new System.EventHandler(this.label1_Click_1);
            // 
            // buttonChooseExcelPath
            // 
            this.buttonChooseExcelPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonChooseExcelPath.Location = new System.Drawing.Point(424, 19);
            this.buttonChooseExcelPath.Name = "buttonChooseExcelPath";
            this.buttonChooseExcelPath.Size = new System.Drawing.Size(75, 23);
            this.buttonChooseExcelPath.TabIndex = 2;
            this.buttonChooseExcelPath.Text = "Bläddra...";
            this.buttonChooseExcelPath.UseVisualStyleBackColor = true;
            this.buttonChooseExcelPath.Click += new System.EventHandler(this.buttonChooseExcelPath_Click);
            // 
            // textBoxExcelDirectory
            // 
            this.textBoxExcelDirectory.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxExcelDirectory.Location = new System.Drawing.Point(9, 21);
            this.textBoxExcelDirectory.Name = "textBoxExcelDirectory";
            this.textBoxExcelDirectory.Size = new System.Drawing.Size(405, 20);
            this.textBoxExcelDirectory.TabIndex = 1;
            this.textBoxExcelDirectory.ModifiedChanged += new System.EventHandler(this.textBoxExcelDirectory_ModifiedChanged);
            this.textBoxExcelDirectory.TextChanged += new System.EventHandler(this.textBoxExcelDirectory_TextChanged);
            // 
            // labelSaveDirectory
            // 
            this.labelSaveDirectory.AutoSize = true;
            this.labelSaveDirectory.Location = new System.Drawing.Point(8, 5);
            this.labelSaveDirectory.Name = "labelSaveDirectory";
            this.labelSaveDirectory.Size = new System.Drawing.Size(170, 13);
            this.labelSaveDirectory.TabIndex = 0;
            this.labelSaveDirectory.Text = "Ange sökväg att spara Excel filer i:";
            // 
            // XMLParser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(519, 514);
            this.Controls.Add(this.Bläddra);
            this.Controls.Add(this.tabControlSG);
            this.Controls.Add(this.DatabasText);
            this.Controls.Add(this.Sökväg);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "XMLParser";
            this.Text = "XMLParser";
            this.tabControlSG.ResumeLayout(false);
            this.tabPageXMLReader.ResumeLayout(false);
            this.tabPageXMLReader.PerformLayout();
            this.tabPageExcelWriter.ResumeLayout(false);
            this.tabPageExcelWriter.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox XMLSträng;
        private System.Windows.Forms.Button LäsInXML;
        private System.Windows.Forms.Label Infotext;
        private System.Windows.Forms.Button Rensa;
        private System.Windows.Forms.Label DatabasText;
        private System.Windows.Forms.TextBox Sökväg;
        private System.Windows.Forms.Button Bläddra;
        private System.Windows.Forms.TabControl tabControlSG;
        private System.Windows.Forms.TabPage tabPageXMLReader;
        private System.Windows.Forms.TabPage tabPageExcelWriter;
        private System.Windows.Forms.Label labelSaveDirectory;
        private System.Windows.Forms.TextBox textBoxExcelDirectory;
        private System.Windows.Forms.Button buttonChooseExcelPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListView listViewSelected;
        private System.Windows.Forms.Button buttonWriteExcelFileToDir;
        private System.Windows.Forms.Label labelNumerOfRows;
        private System.Windows.Forms.Label labelCollectedRows;
        private System.Windows.Forms.Label label2;
    }
}

