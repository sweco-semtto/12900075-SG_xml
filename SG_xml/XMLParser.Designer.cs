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
            this.XMLStr�ng = new System.Windows.Forms.RichTextBox();
            this.L�sInXML = new System.Windows.Forms.Button();
            this.Infotext = new System.Windows.Forms.Label();
            this.Rensa = new System.Windows.Forms.Button();
            this.DatabasText = new System.Windows.Forms.Label();
            this.S�kv�g = new System.Windows.Forms.TextBox();
            this.Bl�ddra = new System.Windows.Forms.Button();
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
            // XMLStr�ng
            // 
            this.XMLStr�ng.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.XMLStr�ng.Location = new System.Drawing.Point(9, 19);
            this.XMLStr�ng.Name = "XMLStr�ng";
            this.XMLStr�ng.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.XMLStr�ng.Size = new System.Drawing.Size(485, 383);
            this.XMLStr�ng.TabIndex = 1;
            this.XMLStr�ng.Text = "";
            this.XMLStr�ng.TextChanged += new System.EventHandler(this.XMLStr�ng_TextChanged);
            // 
            // L�sInXML
            // 
            this.L�sInXML.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.L�sInXML.Location = new System.Drawing.Point(419, 408);
            this.L�sInXML.Name = "L�sInXML";
            this.L�sInXML.Size = new System.Drawing.Size(75, 23);
            this.L�sInXML.TabIndex = 2;
            this.L�sInXML.Text = "L�s in xml";
            this.L�sInXML.UseVisualStyleBackColor = true;
            this.L�sInXML.Click += new System.EventHandler(this.L�sInXML_Click);
            // 
            // Infotext
            // 
            this.Infotext.AutoSize = true;
            this.Infotext.Location = new System.Drawing.Point(6, 3);
            this.Infotext.Name = "Infotext";
            this.Infotext.Size = new System.Drawing.Size(199, 13);
            this.Infotext.TabIndex = 4;
            this.Infotext.Text = "Klistra in xml-str�ngen ifr�n e-post nedan:";
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
            // S�kv�g
            // 
            this.S�kv�g.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.S�kv�g.Location = new System.Drawing.Point(4, 25);
            this.S�kv�g.Name = "S�kv�g";
            this.S�kv�g.Size = new System.Drawing.Size(418, 20);
            this.S�kv�g.TabIndex = 7;
            // 
            // Bl�ddra
            // 
            this.Bl�ddra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.Bl�ddra.Location = new System.Drawing.Point(432, 23);
            this.Bl�ddra.Name = "Bl�ddra";
            this.Bl�ddra.Size = new System.Drawing.Size(75, 23);
            this.Bl�ddra.TabIndex = 8;
            this.Bl�ddra.Text = "Bl�ddra...";
            this.Bl�ddra.UseVisualStyleBackColor = true;
            this.Bl�ddra.Click += new System.EventHandler(this.Bl�ddra_Click);
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
            this.tabPageXMLReader.Controls.Add(this.L�sInXML);
            this.tabPageXMLReader.Controls.Add(this.Rensa);
            this.tabPageXMLReader.Controls.Add(this.Infotext);
            this.tabPageXMLReader.Controls.Add(this.XMLStr�ng);
            this.tabPageXMLReader.Location = new System.Drawing.Point(4, 22);
            this.tabPageXMLReader.Name = "tabPageXMLReader";
            this.tabPageXMLReader.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageXMLReader.Size = new System.Drawing.Size(506, 437);
            this.tabPageXMLReader.TabIndex = 0;
            this.tabPageXMLReader.Text = "L�s in XML till databas";
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
            this.tabPageExcelWriter.Text = "Skriv Excel fr�n databas";
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
            this.labelCollectedRows.Text = "Antal h�mtade ordernummer fr�n databasen:";
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
            this.buttonWriteExcelFileToDir.Text = "Skriv Excel filer till katalog utifr�n valt datum";
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
            this.label1.Text = "V�lj datum att h�mta ordernummer fr�n:";
            this.label1.Click += new System.EventHandler(this.label1_Click_1);
            // 
            // buttonChooseExcelPath
            // 
            this.buttonChooseExcelPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonChooseExcelPath.Location = new System.Drawing.Point(424, 19);
            this.buttonChooseExcelPath.Name = "buttonChooseExcelPath";
            this.buttonChooseExcelPath.Size = new System.Drawing.Size(75, 23);
            this.buttonChooseExcelPath.TabIndex = 2;
            this.buttonChooseExcelPath.Text = "Bl�ddra...";
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
            this.labelSaveDirectory.Text = "Ange s�kv�g att spara Excel filer i:";
            // 
            // XMLParser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(519, 514);
            this.Controls.Add(this.Bl�ddra);
            this.Controls.Add(this.tabControlSG);
            this.Controls.Add(this.DatabasText);
            this.Controls.Add(this.S�kv�g);
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

        private System.Windows.Forms.RichTextBox XMLStr�ng;
        private System.Windows.Forms.Button L�sInXML;
        private System.Windows.Forms.Label Infotext;
        private System.Windows.Forms.Button Rensa;
        private System.Windows.Forms.Label DatabasText;
        private System.Windows.Forms.TextBox S�kv�g;
        private System.Windows.Forms.Button Bl�ddra;
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

