using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.Odbc;


namespace SG_xml
{
    public partial class XMLParser : Form
    {

        #region instansvariabler

        /// <summary>
        /// S�kv�gen till nyckeln i registert. 
        /// </summary>
        private const string nyckelplats = "Software\\SWECO\\XMLReader";

        /// <summary>
        /// Det senaste felmeddelandet som denna klass genererat. 
        /// </summary>
        private string _Felmeddelande;

        #endregion


        /// <summary>
        /// Skapar en ny XMLParser. 
        /// </summary>
        public XMLParser()
        {
            InitializeComponent();

            // H�mtar s�kv�gen till databasen ifr�n registret. 
            S�kv�g.Text = H�mtaRegisternyckel("XMLParser");
            Accessdatabas.S�kv�gDatabas = S�kv�g.Text;

            //Initialize custom components for Excel writer
            InitializeCustomComponentForExcelWriter();

        }

        /// <summary>
        /// Tar hand om klick p� knappen "L�s in xml". 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void L�sInXML_Click(object sender, EventArgs e)
        {
            //L�ser av s�kv�gen igen i fall den har �ndrats direkt i textboxen "S�kv�g"
            Accessdatabas.S�kv�gDatabas = S�kv�g.Text;

            // Tar fram xml-str�ngen
            string xml = XMLStr�ng.Text;

            // Skapar en ny best�llning
            Best�llning best�llning = new Best�llning(xml);
        }

        /// <summary>
        /// H�mtar en nyckel ifr�n registret f�r denna anv�ndare. 
        /// </summary>
        /// <param name="nyckelnamn">Nyckelnamnet. </param>
        /// <returns>V�rdet kopplat till nyckelnamnet. </returns>
        public string H�mtaRegisternyckel(string nyckelnamn)
        {
            try
            {
                // L�ser nyckeln ifr�n registret. 
                RegistryKey nyckel = Registry.CurrentUser.CreateSubKey(nyckelplats);
                string v�rde = nyckel.GetValue(nyckelnamn).ToString();
                nyckel.Close();
                return v�rde;
            }
            catch (Exception ex)
            {
                _Felmeddelande = ex.Message;

                return "";
            }
        }

        /// <summary>
        /// Korrigerar alla html-tecken f�r �, � och � och byter ut dem till riktiga �, � och �. 
        /// </summary>
        /// <param name="HTML_xml">Xml:en med html-tecken i som skall korrigera till vanliga. </param>
        /// <returns>Returnerar xml-str�ngen utan html-tecken i sig. </returns>
        private string Korrigera_HMTL_XML(string HTML_xml)
        {
            // Ers�tter alla &aring med �
            HTML_xml = HTML_xml.Replace("&aring", "�");

            // Ers�tter alla &Aring med �
            HTML_xml = HTML_xml.Replace("&Aring", "�");

            // Ers�tter alla &auml med �
            HTML_xml = HTML_xml.Replace("&auml", "�");

            // Ers�tter alla &Auml med �
            HTML_xml = HTML_xml.Replace("&Auml", "�");

            // Ers�tter alla &ouml med �
            HTML_xml = HTML_xml.Replace("&ouml", "�");

            // Ers�tter alla &Ouml med �
            HTML_xml = HTML_xml.Replace("&Ouml", "�");

            //Ers�tter �vriga specialtecken. 
            HTML_xml = HTML_xml.Replace("&uuml", "�");
            HTML_xml = HTML_xml.Replace("&Uuml", "�");
            HTML_xml = HTML_xml.Replace("&ucirc", "�");
            HTML_xml = HTML_xml.Replace("&Ucirc", "�");
            HTML_xml = HTML_xml.Replace("&egrave", "�");
            HTML_xml = HTML_xml.Replace("&Egrave", "�");
            HTML_xml = HTML_xml.Replace("&eacute", "�");
            HTML_xml = HTML_xml.Replace("&Eacute", "�");
            HTML_xml = HTML_xml.Replace("&amp", "&");
            HTML_xml = HTML_xml.Replace("&lt", "<");
            HTML_xml = HTML_xml.Replace("&gt", ">");
            HTML_xml = HTML_xml.Replace("&quot", "\"");
            HTML_xml = HTML_xml.Replace("&#39", "'");
            
            return HTML_xml;
        }

        /// <summary>
        /// Skapar en nyckel i registret f�r denna anv�ndare. 
        /// </summary>
        /// <param name="v�rde">V�rdet p� nyckeln. </param>
        /// <param name="nyckel">Nyckelns namn. </param>
        private void SkapaRegisternyckel(string nyckelnamn, string v�rde)
        {
            try
            {
                // Skriver nyckeln till registret. 
                RegistryKey nyckel = Registry.CurrentUser.CreateSubKey(nyckelplats);
                nyckel.SetValue(nyckelnamn, v�rde);
                nyckel.Close();
            }
            catch (Exception ex)
            {
                _Felmeddelande = ex.Message;
            }
        }

        /// <summary>
        /// Tar bort text som finns innan xml-str�ngen. 
        /// </summary>
        /// <param name="Epost_text">Tecken ifr�n eposten. </param>
        /// <returns>En Html-xml som har speciella tecken f�r �, � och �. </returns>
        private string TaBortTextUtanf�rXML(string Epost_text)
        {
            try
            {
                // Tar bort alla tecken framf�r xml-delen, s�ker framifr�n efter ett <-tecken. 
                int b�rjan = Epost_text.IndexOf("<") == -1 ? 0 : Epost_text.IndexOf("<");

                // Tar bort alla tecken bakom xml-delen, s�ker bakifr�n efter ett >-tecken. 
                int slutet = Epost_text.LastIndexOf(">") == -1 ? Epost_text.Length - 1 : Epost_text.LastIndexOf(">");

                return Epost_text.Substring(b�rjan, slutet - b�rjan + 1);
            }
            // Om ingen text finns innan eller efter
            catch (ArgumentOutOfRangeException)
            {
                return Epost_text;
            }
        }

        #region Lyssnare i f�nstret

        /// <summary>
        /// Tar hand om klick p� bl�ddra-knappen. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Bl�ddra_Click(object sender, EventArgs e)
        {

            

            StopEventForDateTimePicker = false;

            // �ppnar ett f�nster f�r att peka ut accessdatabasen. 
            OpenFileDialog �ppnaF�nstret = new OpenFileDialog();
            �ppnaF�nstret.AddExtension = true;
            �ppnaF�nstret.CheckPathExists = true;
            �ppnaF�nstret.DefaultExt = "mdb";
            �ppnaF�nstret.Filter = "Accessdatabaser (*.mdb)|*.mdb|Alla filer (*.*)|*.*";
            �ppnaF�nstret.FilterIndex = 1;
            �ppnaF�nstret.Title = "Bl�ddra till accessdatabas";
            if (S�kv�g.Text.ToString().EndsWith("\\"))
                �ppnaF�nstret.FileName = S�kv�g.Text.ToString().TrimEnd('\\');
            else
                �ppnaF�nstret.FileName = S�kv�g.Text;

            �ppnaF�nstret.ShowDialog();
            S�kv�g.Text = �ppnaF�nstret.FileName;
            Accessdatabas.S�kv�gDatabas = S�kv�g.Text;
            SkapaRegisternyckel("XMLParser", S�kv�g.Text);

            //Run event for data load
            dateTimePickerFrom_ValueChanged(null, null);

            oldDatabaseString = S�kv�g.Text.ToString().TrimEnd('\\');
        }

        //Stop dateTimePickerEvent, fix a bugg
        protected bool StopEventForDateTimePicker = false;

       
        /// <summary>
        /// Tar hand om klick p� knappen "Rensa" och rensar all text i "XMLStr�ng". 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Rensa_Click(object sender, EventArgs e)
        {
            XMLStr�ng.Clear();
        }

        /// <summary>
        /// Anropas n�r texten i xml-str�ngen �ndras. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void XMLStr�ng_TextChanged(object sender, EventArgs e)
        {
            // Fixar till texten fr�n eposten till en korrekt xml-str�ng. 
            XMLStr�ng.Text = TaBortTextUtanf�rXML(XMLStr�ng.Text);
            XMLStr�ng.Text = Korrigera_HMTL_XML(XMLStr�ng.Text);
            XMLStr�ng.SelectionStart = XMLStr�ng.Text.Length;
        }

        #endregion


        #region ExcelWriterPage


        // Create a new DateTimePicker control and initialize it.
        protected DateTimePicker dateTimePickerFrom = new DateTimePicker();

        // Create a new DateTimePicker control and initialize it.
        protected DateTimePicker dateTimePickerTo = new DateTimePicker();

        //Column sorter for list views
        private ListViewColumnSorter lvwColumnSorter;

        /// <summary>
        /// InitializeCustomComponentForExcelWriter: Initialize custom components for Excel Writer
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        private void InitializeCustomComponentForExcelWriter()
        {
            // Set the MinDate and MaxDate for 'from' date time picker.
            dateTimePickerFrom.MinDate = new DateTime(1985, 6, 20);
            dateTimePickerFrom.MaxDate = DateTime.Today;

            // Set the CustomFormat string.
            //dateTimePickerFrom.CustomFormat = "MMMM dd, yyyy - dddd";
            dateTimePickerFrom.CustomFormat = "'den' ddMMMMyyy - dddd";
            dateTimePickerFrom.Format = DateTimePickerFormat.Custom; //.Short; //.Custom;

            // Show the CheckBox and display the control as an up-down control.
            dateTimePickerFrom.ShowCheckBox = false;

            //Add date time picker to controll
            this.tabPageExcelWriter.Controls.Add(dateTimePickerFrom);
            dateTimePickerFrom.Location = new System.Drawing.Point(8, 74);
            
            //create event for date time picker.
            dateTimePickerFrom.ValueChanged += new EventHandler(dateTimePickerFrom_ValueChanged);


            // Set the MinDate and MaxDate for 'to' date time picker
            dateTimePickerTo.MinDate = new DateTime(1985, 6, 20);
            dateTimePickerTo.MaxDate = DateTime.Today;

            // Set the CustomFormat string.
            //dateTimePickerTo.CustomFormat = "MMMM dd, yyyy - dddd";
            dateTimePickerTo.CustomFormat = "'den' ddMMMMyyy - dddd";
            dateTimePickerTo.Format = DateTimePickerFormat.Custom;

            // Show the CheckBox and display the control as an up-down control.
            dateTimePickerTo.ShowCheckBox = true;

            //Add date time picker to controll
            this.tabPageExcelWriter.Controls.Add(dateTimePickerTo);
            dateTimePickerTo.Location = new System.Drawing.Point(238, 74);

            //create event for date time picker.
            dateTimePickerTo.ValueChanged += new EventHandler(dateTimePickerTo_ValueChanged);

            //Collect data for excel diretory from register
            textBoxExcelDirectory.Text = H�mtaRegisternyckel("ExcelWriter");

            //Set sort order
            SetSortOrder();
           
            //Initialize list view with fields
            listViewSelected.Columns.Add("Ordernr", 55, HorizontalAlignment.Left);
            listViewSelected.Columns.Add("Datum", 80, HorizontalAlignment.Left);           
            listViewSelected.Columns.Add("F�retagsnamn", 120, HorizontalAlignment.Left);
            listViewSelected.Columns.Add("Region/F�rvaltning", 120, HorizontalAlignment.Left);
            listViewSelected.Columns.Add("Distrikt/Omr�de", 110, HorizontalAlignment.Left);

            // Create an instance of a ListView column sorter and assign it 
            // to the ListView control.
            lvwColumnSorter = new ListViewColumnSorter();
            this.listViewSelected.ListViewItemSorter = lvwColumnSorter;

            //If database directory is nott null, trigg event and load data
            if (S�kv�g.Text != "")
                dateTimePickerFrom_ValueChanged(null, null);

            //Set last and first date
            SetFromDate();
            SetToDate();
        }

        /// <summary>
        /// SetSortOrder: Set sort order if found from register
        /// </summary>
        private void SetSortOrder()
        {
            try
            {
                if (H�mtaRegisternyckel("SortOrder") != "")
                    mySortOrderColumn = int.Parse(H�mtaRegisternyckel("SortOrder"));
            }
            catch (SystemException ex)
            {
            }
        }


        /// <summary>
        /// dateTimePickerFrom_ValueChanged: Event that occur when date is changed
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <param name="o"></param>
        /// <param name="e"></param>
        private void dateTimePickerFrom_ValueChanged(object o, EventArgs e)
        {
            if (StopEventForDateTimePicker)
                return;

            dateTimePickerTo.ValueChanged -= new EventHandler(dateTimePickerFrom_ValueChanged);
         
            try
            {
                //Check if checkbox for dateTimePickerTo is checked, if checked runt check to corret dates
                if (!dateTimePickerTo.Checked)
                    SetFromDate();
                else
                    CheckIfFromOk(dateTimePickerFrom.Value.Year, dateTimePickerFrom.Value.Month, dateTimePickerFrom.Value.Day);

                //Extract orderNR to list view
                GetOrderNRFromDatabase();
            }
            catch (SystemException ex)
            {
                MessageBox.Show(ex.ToString());

                dateTimePickerTo.ValueChanged += new EventHandler(dateTimePickerFrom_ValueChanged);
            }

            dateTimePickerTo.ValueChanged += new EventHandler(dateTimePickerFrom_ValueChanged);
        }

        /// <summary>
        /// dateTimePickerToIsChecked: Remebers if the dateTimePicker is cheked to be used with dateTimePickerEvent
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        protected bool dateTimePickerToIsChecked = true;

        /// <summary>
        /// dateTimePickerFrom_ValueChanged: Event that occur when date is changed
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <param name="o"></param>
        /// <param name="e"></param>
        private void dateTimePickerTo_ValueChanged(object o, EventArgs e)
        {
            if (StopEventForDateTimePicker)
                return;
            try
            {                
                //Check if date time checkbox is checked
                if (dateTimePickerTo.Checked)
                {
                    //Check if checkbox was checked
                    if (!dateTimePickerToIsChecked)
                    {
                        CheckLegalDate();
                        dateTimePickerToIsChecked = true;
                    }
                }
                else
                    dateTimePickerToIsChecked = false;

                //Check and corret dates
                CheckIfToOk(dateTimePickerTo.Value.Year, dateTimePickerTo.Value.Month, dateTimePickerTo.Value.Day);
                
                //Extract orderNR to list view
                GetOrderNRFromDatabase();             
            }
            catch (SystemException ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        /// <summary>
        /// buttonChooseExcelPath_Click: Select folder to store Excel files in
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>     
        private void buttonChooseExcelPath_Click(object sender, EventArgs e)
        {
            //Select folder for Excel files
            FolderBrowserDialog newPathToExcelFiles = new FolderBrowserDialog();
            newPathToExcelFiles.SelectedPath = textBoxExcelDirectory.Text;

            newPathToExcelFiles.ShowDialog();

            newPathToExcelFiles.Description = "Bl�ddra till katalog f�r Excel filer";
            textBoxExcelDirectory.Text = newPathToExcelFiles.SelectedPath;

            //Write to register with new excel directory
            SkapaRegisternyckel("ExcelWriter", textBoxExcelDirectory.Text);
        }


        //Remeber last date in database query
        private string oldDateStringFrom = "";
        private string oldDateStringTo = "";
        private string oldDatabaseString = "";

        /// <summary>
        /// GetOrderNRFromDatabase: Extract data to list view
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void GetOrderNRFromDatabase()
        {
            if (S�kv�g.Text == "")
                MessageBox.Show("V�lj Microsoft Access databas att arbeta med f�rst");
            else
            {
                //Create new date based on selected date
                string myDateStringFrom = CreateOrderDateFrom();
                string myDateStringTo = "";
                
                if (dateTimePickerTo.Checked) 
                    myDateStringTo = CreateOrderDateTo();

                if (oldDateStringFrom == myDateStringFrom && myDateStringTo == oldDateStringTo && S�kv�g.Text.ToString().TrimEnd('\\') == oldDatabaseString)
                    return;

                oldDateStringFrom = myDateStringFrom;
                oldDateStringTo = myDateStringTo;

                try
                {
                    DataSet myData;
                    
                    //Extract data from Access database
                    if (dateTimePickerTo.Checked)
                        myData = Accessdatabas.L�sIfr�nDatabas("Select * from f�retag where Best�llningsdatum between #" + myDateStringFrom + "# and #" + myDateStringTo + "# order by Ordernr");

                    else
                        myData = Accessdatabas.L�sIfr�nDatabas("Select * from f�retag where Best�llningsdatum = #" + myDateStringFrom + "# order by Ordernr");

                    //Populate list view with selected dates
                    if (myData != null)
                        populateListViewWithSelectedDate(myData);
                    else
                    {
                        // Clear the SearchResultIs control
                        listViewSelected.Items.Clear();
                        labelNumerOfRows.Text = "0";
                        StopEventForDateTimePicker = true;
                    }

                }
                catch (SystemException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }



        private int mySortOrderColumn = 1;
        /// <summary>
        /// populateListViewWithSelectedDate: Populate the list view with preview data
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <param name="selectedDates"></param>
        private void populateListViewWithSelectedDate(DataSet selectedDates)
        {

            // Clear the SearchResultIs control
            listViewSelected.Items.Clear();
            labelNumerOfRows.Text = "0";

            try
            {
                // Get the table from the data set
                DataTable dtable = selectedDates.Tables[0];

                // Sort the items in the list in ascending order.
                listViewSelected.Sorting = SortOrder.Ascending;

                labelNumerOfRows.Text = dtable.Rows.Count.ToString();

                // Display items in the ListView control
                for (int i = 0; i < dtable.Rows.Count; i++)
                {
                    DataRow drow = dtable.Rows[i];

                    // Only row that have not been deleted
                    if (drow.RowState != DataRowState.Deleted)
                    {
                        // Define the list items
                        string dateOfOrder = drow["Best�llningsdatum"].ToString();
                        dateOfOrder = dateOfOrder.Remove(dateOfOrder.IndexOf(' '));

                        //Polulate list view with data from dataSet
                        ListViewItem lvi = new ListViewItem(drow["Ordernr"].ToString());
                        lvi.SubItems.Add(dateOfOrder);
                        lvi.SubItems.Add(drow["F�retagsnamn"].ToString());
                        lvi.SubItems.Add(drow["Region_F�rvaltning"].ToString());
                        lvi.SubItems.Add(drow["Distrikt_Omr�de"].ToString());
                        lvi.Checked = true;

                        listViewSelected.Items.Add(lvi);
                    }

                    //Sort list view in Ascending order 
                    lvwColumnSorter.SortColumn = mySortOrderColumn;

                    //Sort list view in Ascending order   
                    listViewSelected.Sort();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        /// <summary>
        /// CollectOrderNrFromListView: Generate a Query depending on OrderNr that is checked
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <returns></returns>
        private string CollectOrderNrFromListView()
        {
            string mySQLQuery = "";

            //Add all order number to query that is checked
            for (int i = 0; i < listViewSelected.Items.Count; i++)
            {
                if (listViewSelected.Items[i].Checked)
                    mySQLQuery = mySQLQuery + "Ordernr = '" + listViewSelected.Items[i].SubItems[0].Text.ToString() + "' OR ";

            }

            if (mySQLQuery.Length > 0)
                mySQLQuery = mySQLQuery.Substring(0, mySQLQuery.Length - 4);

            return mySQLQuery;
        }

        //Remeber last date
        protected int fromYear = 0;
        protected int fromMoth = 0;
        protected int fromDay = 0;

        //Remeber first date
        protected int toYear = 0;
        protected int toMoth = 0;
        protected int toDay = 0;

        /// <summary>
        /// SetFromDate: Sets from date variables
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        private void SetFromDate()
        {
            fromYear = dateTimePickerFrom.Value.Year;
            fromMoth = dateTimePickerFrom.Value.Month;
            fromDay = dateTimePickerFrom.Value.Day;
        }

        /// <summary>
        /// SetToDate: Sets to date variables
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        private void SetToDate()
        {
            toYear = dateTimePickerTo.Value.Year;
            toMoth = dateTimePickerTo.Value.Month;
            toDay = dateTimePickerTo.Value.Day;
        }

        /// <summary>
        /// CreateDate: Creates a DateTime objekt
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// <param name="year"></param>
        /// <param name="moth"></param>
        /// <param name="day"></param>
        /// <returns></returns>
        private DateTime CreateDate(int year, int moth, int day)
        {
            DateTime myDateTime = new DateTime(year, moth, day);

            return myDateTime;
        }

        /// <summary>
        /// dateTimeFix: Needed for DateTimePicker bugg
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        bool dateTimeFix = true;

        /// CheckIfToOk: Checks if to date is OK
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        private bool CheckIfToOk(int year, int month, int day)
        {
            DateTime fromDate = CreateDate(fromYear, fromMoth, fromDay);
            DateTime toDate = CreateDate(year, month, day);

            int IsOK = fromDate.CompareTo(toDate);

            //If date is not OK, change date for from date time picker
            if (IsOK > 0) 
            {
                dateTimePickerFrom.ValueChanged -= new EventHandler(dateTimePickerFrom_ValueChanged);
                    
                if (!dateTimeFix)
                {
                    dateTimeFix = true;                        
                    DateTime newDate = CreateDate(dateTimePickerTo.Value.Year, dateTimePickerTo.Value.Month, dateTimePickerTo.Value.Day);
                    dateTimePickerFrom.Value = newDate;                       
                }
                else
                {
                    DateTime newDate = CreateDate(dateTimePickerTo.Value.Year, dateTimePickerTo.Value.Month, dateTimePickerTo.Value.Day);
                    dateTimePickerFrom.Value = newDate;                  
                    dateTimeFix = false;

                }
                   
                dateTimePickerFrom.ValueChanged += new EventHandler(dateTimePickerFrom_ValueChanged);

                //Set new from date
                SetFromDate();
            }
         
            //Set new to date
            SetToDate();
            return true;
        }


        /// <summary>
        /// CheckIfFromOk: Checks if from date is OK
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <param name="day"></param>
        /// <returns></returns>
        private bool CheckIfFromOk(int year, int month, int day)
        {
            //If year == 0 to date hase not been initialized
            if (toYear == 0)
                return true;

            DateTime toDate = CreateDate(toYear, toMoth, toDay);
            DateTime fromDate = CreateDate(year, month, day);

            //Check if date is OK
            int IsOK = fromDate.CompareTo(toDate);

            //Date is not OK, change it
            if (IsOK > 0)
            {
                dateTimePickerTo.ValueChanged -= new EventHandler(dateTimePickerFrom_ValueChanged);
                
                if (!dateTimeFix)
                {
                    dateTimeFix = true;                      
                    DateTime newDate = CreateDate(dateTimePickerFrom.Value.Year, dateTimePickerFrom.Value.Month, dateTimePickerFrom.Value.Day);
                    dateTimePickerTo.Value = newDate;                       
                }
                else
                {                       
                    DateTime newDate = CreateDate(dateTimePickerFrom.Value.Year, dateTimePickerFrom.Value.Month, dateTimePickerFrom.Value.Day);
                    dateTimePickerTo.Value = newDate;
                    dateTimeFix = false;
                }

                dateTimePickerTo.ValueChanged += new EventHandler(dateTimePickerTo_ValueChanged);

                //Set to date  
                SetToDate();            
            }

            SetFromDate();
            return true;
        }

        /// <summary>
        /// CheckLegalDate: Checks if the date is OK
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        private void CheckLegalDate()
        {
            if (toYear == 0)
                return;

            DateTime toDate = CreateDate(fromYear, fromMoth, fromDay);
            DateTime fromDate = CreateDate(dateTimePickerTo.Value.Year, dateTimePickerTo.Value.Month, dateTimePickerTo.Value.Day);

            int IsOK = fromDate.CompareTo(toDate);

            //Date not OK, change it
            if (IsOK < 0)
            {
                dateTimePickerTo.ValueChanged -= new EventHandler(dateTimePickerFrom_ValueChanged);
                dateTimePickerTo.Value = new DateTime(dateTimePickerFrom.Value.Year, dateTimePickerFrom.Value.Month, dateTimePickerFrom.Value.Day);
                dateTimePickerTo.ValueChanged += new EventHandler(dateTimePickerTo_ValueChanged);
            }
        }

        /// <summary>
        /// CreateOrderDate: Create new dateString to be used when extracting data from Microsoft Access database
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <returns></returns>
        private string CreateOrderDateFrom()
        {

            string dateString = dateTimePickerFrom.Value.Year.ToString();

            if (dateTimePickerFrom.Value.Month < 10)
                dateString = dateString + "-0" + dateTimePickerFrom.Value.Month.ToString();
            else
                dateString = dateString + "-" + dateTimePickerFrom.Value.Month.ToString();

            if (dateTimePickerFrom.Value.Day < 10)
                dateString = dateString + "-0" + dateTimePickerFrom.Value.Day.ToString();
            else
                dateString = dateString + "-" + dateTimePickerFrom.Value.Day.ToString();

            return dateString;
        }

        /// <summary>
        /// CreateOrderDate: Create new dateString to be used when extracting data from Microsoft Access database
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <returns></returns>
        private string CreateOrderDateTo()
        {

            string dateString = dateTimePickerTo.Value.Year.ToString();

            if (dateTimePickerTo.Value.Month < 10)
                dateString = dateString + "-0" + dateTimePickerTo.Value.Month.ToString();
            else
                dateString = dateString + "-" + dateTimePickerTo.Value.Month.ToString();

            if (dateTimePickerTo.Value.Day < 10)
                dateString = dateString + "-0" + dateTimePickerTo.Value.Day.ToString();
            else
                dateString = dateString + "-" + dateTimePickerTo.Value.Day.ToString();

            return dateString;
        }

        /// <summary>
        /// buttonWriteExcelFileToDir_Click: Write Excel file to path on hardrive
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonWriteExcelFileToDir_Click(object sender, EventArgs e)
        {

            //Create new date based on selected date
            string myDateStringFrom = CreateOrderDateFrom();
            string myDateStringTo = CreateOrderDateTo();
            bool reportGenerated = false;


            ExcelWriter myExcelWriter = new ExcelWriter();

            try
            {
                if (textBoxExcelDirectory.Text == "")
                    MessageBox.Show("Ingen katalog f�r excel filer vald");

                else
                {
                    string myOrdNr = CollectOrderNrFromListView();

                    if (myOrdNr.Length == 0)
                    {
                        MessageBox.Show("Inga ordernummer att exportera");
                        return;
                    }
                    //Extract data from Access database     
                    DataSet myData = Accessdatabas.L�sIfr�nDatabas("Select * from f�retag where " + myOrdNr);

                    //Write Excel files
                    if (myData != null)
                        if (myData.Tables[0].Rows.Count > 0)
                        {
                            //Create FileIO objekt to check path
                            FileIO myFile = new FileIO();

                            //Check if path exist, or has bo be created
                            if (myFile.ChechPath(textBoxExcelDirectory.Text))
                            {
                                //Check if path was created successfully, if it was created
                                if (myFile.CheckPathAfterCreate(textBoxExcelDirectory.Text))
                                { 
                                    //Write the excel files
                                    reportGenerated = myExcelWriter.WriteExcelFileFromDataset(myData, textBoxExcelDirectory.Text);

                                    GC.Collect();

                                    


                                    if (reportGenerated)
                                    {
                                        DialogResult processDone = MessageBox.Show("Operationen klar, Excel filer �terfinns i katalog " + textBoxExcelDirectory.Text + " Vill du �ppna katalogen?", "Klart", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                        if (processDone == DialogResult.Yes)
                                        {
                                            System.Diagnostics.Process.Start(textBoxExcelDirectory.Text);
                                        }
                                    }
                                    else
                                    {
                                        DialogResult processDone = MessageBox.Show("Operationen klar, inga nya excel filer skapade. �ppna katalog " + textBoxExcelDirectory.Text + " end�?", "Klart", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                        if (processDone == DialogResult.Yes)
                                        {
                                            System.Diagnostics.Process.Start(textBoxExcelDirectory.Text);
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("S�kv�gen " + textBoxExcelDirectory.Text + " �r inte giltig");
                                }
                            }
                        }         
                }

            }
            catch (SystemException ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        /// <summary>
        /// textBoxExcelDirectory_TextChanged: Write new key when textbox changed
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBoxExcelDirectory_TextChanged(object sender, EventArgs e)
        {
            SkapaRegisternyckel("ExcelWriter", textBoxExcelDirectory.Text);
        }

        /// <summary>
        /// listViewSelected_ColumnClick: Sort list view depending on column user clicks
        /// Fredrik Bj�rklund FEBJ, SWECO Position, 2008
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listViewSelected_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            //Sort list view in Ascending order 
            lvwColumnSorter.SortColumn = e.Column;
            //lvwColumnSorter.Order = SortOrder.Ascending;

            //Set sort order variable
            mySortOrderColumn = e.Column;

            //Sort list view in Ascending order   
            listViewSelected.Sort();

            try
            {
                SkapaRegisternyckel("SortOrder", e.Column.ToString());
            }
            catch (SystemException ex)
            {

            }
        }
                 

        #endregion


        private void listViewSelected_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBoxExcelDirectory_ModifiedChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void listViewSelected_BindingContextChanged(object sender, EventArgs e)
        {

        }

        private void listViewSelected_TabIndexChanged(object sender, EventArgs e)
        {

        }

        



    }
}