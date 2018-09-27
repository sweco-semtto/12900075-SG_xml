/****************************************************
 * ImageRequest.cs l�ser ur information fr�n en 
 * inkommande xml-request.
 * 
 * 
 * LSAM, SWECO Position, 2005
 * 
 *****************************************************/

using System;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Windows.Forms;
using System.Xml;

namespace SG_xml
{
	/// <summary>
	/// L�ser information fr�n den inkommande xml-str�ngen f�r f�retagstaggen. F�r att skapa ett nytt F�rtagsobjekt
    /// skall man anv�nda sig av metoden ByggUppObjekt(string) eftersom klassen �r skriven enligt designm�nstret
    /// "Factory". 
    /// 
    /// Skapad av MTTO. 
	/// </summary>
	public class F�retag
    {
        #region instansvariabler

        private string _Tidsst�mpel;
        private string _Ordernummer;
        private string _Best�llningsdatum;
		private string _F�retagsnamn;
		private string _Faktureringsadress;
		private string _Postnummer;
		private string _Ort;
		private string _Region_F�rvaltning;
		private string _VAT;
		private string _Distrikt_Omr�de;
		private string _Best�llningsreferens;
		private string _Kontaktperson1;
        private string _TelefonArb1;
        private string _TelefonMob1;
        private string _TelefonHem1;
        private string _Epost1;
        private string _Kontaktperson2;
        private string _TelefonArb2;
        private string _TelefonMob2;
        private string _TelefonHem2;
        private string _Epost2;
        private string _Kommentar;

        private static string _Felmeddelande;
        private static bool _FelIXML = false;

        #endregion

        /// <summary>
        /// Skapar ett nytt F�retagsobjekt. 
        /// </summary>
        /// <param name="xml">Den xml-str�ng som bygger upp objektet. </param>
        protected F�retag(string xml)
		{
			L�sIn(xml);
		}

        public static F�retag ByggUppObjekt(string xml)
        {
            F�retag nyttObjekt = new F�retag(xml);

            return nyttObjekt;
        }

        /// <summary>
        /// Bygger upp ett SQL-kommando (OleDbCommand) utifr�n en best�llning. 
        /// </summary>
        /// <returns>Returnerar ett sql-kommando som l�gger in alla uppgifter i en databas. </returns>
        public OleDbCommand ByggUppSQL()
        {
            // Kommandot som skall byggas upp
            OleDbCommand kommando = new OleDbCommand();
         
            // Bygger upp grunden f�r sqkl.kommandot
            string SQLSats = "INSERT INTO F�retag (Ordernr, Best�llningsreferens, Best�llningsdatum, Tidsst�mpel, ";
            SQLSats += "F�retagsnamn, Faktureringsadress, Postnummer, Ort, Region_F�rvaltning, Distrikt_Omr�de, ";
            SQLSats += "VAT, Kontaktperson1, TelefonArb1, TelefonMobil1, TelefonHem1, Epostadress1, Kontaktperson2, ";
            SQLSats += "TelefonArb2, TelefonMobil2, TelefonHem2, Epostadress2, Kommentar) VALUES (";

            // L�gger till alla uppgifter ifr�n best�llningen. 
            SQLSats += "@_Ordernummer, @_Best�llningsreferens, @_Best�llningsdatum, @_Tidsst�mpel, @_F�retagsnamn, ";
            SQLSats += "@_Faktureringsadress, @_Postnummer, @_Ort, @_Region_F�rvaltning, @_Distrikt_Omr�de, ";
            SQLSats += "@_VAT, @_Kontaktperson1, @_TelefonArb1, @_TelefonMob1, ";
            SQLSats += "@_TelefonHem1, @_Epost1, @_Kontaktperson2, ";
            SQLSats += "@_TelefonArb2, @_TelefonMob2, @_TelefonHem2, @_Epost2, @_Kommentar)";

            // Anger kommandotexten. 
            kommando.CommandText = SQLSats;

            // Anger vilka typer som �r legala i kommandot. 
            kommando.Parameters.Add("@_Ordernummer", OleDbType.VarChar);
            kommando.Parameters.Add("@_Best�llningsreferens", OleDbType.VarChar);
            kommando.Parameters.Add("@_Best�llningsdatum", OleDbType.Date);
            kommando.Parameters.Add("@_Tidsst�mpel", OleDbType.VarChar);
            kommando.Parameters.Add("@_F�retagsnamn", OleDbType.VarChar);
            kommando.Parameters.Add("@_Faktureringsadress", OleDbType.VarChar);
            kommando.Parameters.Add("@_Postnummer", OleDbType.VarChar);
            kommando.Parameters.Add("@_Ort", OleDbType.VarChar);
            kommando.Parameters.Add("@_Region_F�rvaltning", OleDbType.VarChar);
            kommando.Parameters.Add("@_Distrikt_Omr�de", OleDbType.VarChar);
            kommando.Parameters.Add("@_VAT", OleDbType.VarChar);
            kommando.Parameters.Add("@_Kontaktperson1", OleDbType.VarChar);
            kommando.Parameters.Add("@_TelefonArb1", OleDbType.VarChar);
            kommando.Parameters.Add("@_TelefonMob1", OleDbType.VarChar);
            kommando.Parameters.Add("@_TelefonHem1", OleDbType.VarChar);
            kommando.Parameters.Add("@_Epost1", OleDbType.VarChar);
            kommando.Parameters.Add("@_Kontaktperson", OleDbType.VarChar);
            kommando.Parameters.Add("@_TelefonArb2", OleDbType.VarChar);
            kommando.Parameters.Add("@_TelefonMob2", OleDbType.VarChar);
            kommando.Parameters.Add("@_TelefonHem2", OleDbType.VarChar);
            kommando.Parameters.Add("@_Epost2", OleDbType.VarChar);
            kommando.Parameters.Add("@_Kommentar", OleDbType.VarChar);

            // L�gger till alla v�rden
            kommando.Parameters[0].Value = _Ordernummer;
            if (_Best�llningsreferens.Equals(String.Empty))
                kommando.Parameters[1].Value = _Ordernummer;
            else
                kommando.Parameters[1].Value = _Best�llningsreferens;
            kommando.Parameters[2].Value = _Best�llningsdatum;
            kommando.Parameters[3].Value = _Tidsst�mpel;
            kommando.Parameters[4].Value = _F�retagsnamn;
            kommando.Parameters[5].Value = _Faktureringsadress;
            kommando.Parameters[6].Value = _Postnummer;
            kommando.Parameters[7].Value = _Ort;
            kommando.Parameters[8].Value = _Region_F�rvaltning;
            kommando.Parameters[9].Value = _Distrikt_Omr�de;
            kommando.Parameters[10].Value = _VAT;
            kommando.Parameters[11].Value = _Kontaktperson1;
            kommando.Parameters[12].Value = _TelefonArb1;
            kommando.Parameters[13].Value = _TelefonMob1;
            kommando.Parameters[14].Value = _TelefonHem1;
            kommando.Parameters[15].Value = _Epost1;
            kommando.Parameters[16].Value = _Kontaktperson2;
            kommando.Parameters[17].Value = _TelefonArb2;
            kommando.Parameters[18].Value = _TelefonMob2;
            kommando.Parameters[19].Value = _TelefonHem2;
            kommando.Parameters[20].Value = _Epost2;
            kommando.Parameters[21].Value = _Kommentar;

            return kommando;
        }

        /// <summary>
        /// Bygger upp ett SQL-kommando (OleDbCommand) f�r att uppdatera ordernummret. Anledningen till detta �r att 
        /// alla kartprogram inte kan l�sa r�knaren f�r ordernummret. 
        /// </summary>
        /// <returns>Returnerar ett sql-kommando som uppdaterar ordernummret i databasen.  </returns>
        public OleDbCommand ByggUppSQLL�ggTillOrdernummer(string ordernummer)
        {
            // Kommandot som skall byggas upp
            OleDbCommand kommando = new OleDbCommand();

            // Bygger upp grunden f�r sqkl.kommandot
            string SQLSats = "UPDATE F�retag SET OrdernrText=@_Best�llningsreferens WHERE Ordernr=@_Best�llningsreferens";

            // Anger kommandotexten. 
            kommando.CommandText = SQLSats;

            // Anger vilka typer som �r legala i kommandot. 
            kommando.Parameters.Add("@_Best�llningsreferens", OleDbType.VarChar);

            // L�gger till alla v�rden
            kommando.Parameters[0].Value = ordernummer;

            return kommando;
        }

        /// <summary>
        /// Bygger upp en SQL-fr�ga som h�mtar ordernummer. 
        /// </summary>
        /// <param name="funktionAttL�saInMed">Funktion att l�sa med ifr�n databasen</param>
        /// <returns>Returnerar en str�ng som hittar ordernummret till denna order. </returns>
        public string H�mtaOrdernummer(L�sIfr�nDatabas funktionAttL�saInMed)
        {
            string SQLFr�ga = "SELECT * FROM F�retag where Tidsst�mpel = '" + _Tidsst�mpel + "'";

            // H�mtar ordernummret f�r att kunna skriva in startplatsen ifr�n databasen. 
            DataSet data = funktionAttL�saInMed(SQLFr�ga);
            
            // Tar fram ordernummret fr�n datasetet. 
            try
            {
                // Tar fram v�rdet i rad ett kolumn ett (d.v.s. ordernummer). 
                return data.Tables[0].Rows[0][1].ToString();
            }
            // Inget fel borde intr�ffa, om det g�r det �r n�got alvarligt fel. 
            catch (Exception)
            {
                MessageBox.Show("Felaktig data i databasen, v�nligen kontakta Sweco med detta meddelande. ", "Felaktigt v�rde i databasen", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

            return null;
        }
		
		/// <summary>
		/// Tar in en xml-str�ng och fyller p� samtliga instansvariabler i denna klass med information fr�n
        /// den xml:en. 
		/// </summary>
		/// <param name="xml">Inkommande xml-str�ng. </param>
        private void L�sIn(string xml)
		{
            try
            {
                // Anger var siffror har f�r kommaseparerare. 
                NumberFormatInfo nf = new NumberFormatInfo();
                nf.NumberDecimalSeparator = ".";

                // Laddar in xml-dokumentet. 
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("ogc", "http://www.opengis.net/ogc");

                // noden som anv�nds f�r att l�sa in de olika uppgifterna. 
                XmlNode xmlNode;

                // L�ser in ordernummer. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Ordernummer");
                _Ordernummer = xmlNode != null ? xmlNode.InnerText : MySqlCommunicator.GetNewOrdernumber();

                // L�ser in best�llningsdatum. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Best�llningsdatum");
                _Best�llningsdatum = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in f�retagsnman.
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::F�retagsnamn");
                _F�retagsnamn = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in faktureringsadress.
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Faktureringsadress");
                _Faktureringsadress = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�sser in postnummer.
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Postnummer");
                _Postnummer = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in ort. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Ort");
                _Ort = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in region/f�rvaltning
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Region_F�rvaltning");
                _Region_F�rvaltning = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in VAT
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::VAT");
                _VAT = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�sser in distrikt/omr�de
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Distrikt_Omr�de");
                _Distrikt_Omr�de = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in best�llningsreferens
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Best�llningsreferens");
                _Best�llningsreferens = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in kontaktperson ett. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Kontaktperson1");
                _Kontaktperson1 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in kontaktperson etts telefon till arbetet. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Telefon_Arb1");
                _TelefonArb1 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in kontaktperson etts mobiltelefonnummer. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Telefon_Mob1");
                _TelefonMob1 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in kontaktperson etts telefon till hemmet. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Telefon_Hem1");
                _TelefonHem1 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in kontaktperson etts e-postadress. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::E-postadress1");
                _Epost1 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in kontaktperson tv�. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Kontaktperson2");
                _Kontaktperson2 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in kontaktperson tv�s telefon till arbetet. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Telefon_Arb2");
                _TelefonArb2 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in kontaktperson tv�s mobiltelefonnummer. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Telefon_Mob2");
                _TelefonMob2 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in kontaktperson tv�s telefon till hemmet. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::Telefon_Hem2");
                _TelefonHem2 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in kontaktperson tv�s e-postadress. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::F�retag/child::E-postadress2");
                _Epost2 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�ser in kommentaren sist i xml:en. 
                xmlNode = xmlDoc.SelectSingleNode("Best�llning/child::Kommentar");
                _Kommentar = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // L�gger in en tidsst�mpel n�r allt inl�st ifr�n xml:en.
                _Tidsst�mpel = DateTime.Now.Ticks.ToString();
            }
            catch (XmlException xmlex)
            {
                _FelIXML = true;
                _Felmeddelande = xmlex.Message;

                // Meddelare anv�ndaren om detta fel. 
                MessageBox.Show("Xml-str�ngen inneh�ller fel inom F�retagstaggen och kan ej anv�ndas f�r att spara data med. \nLeta efter felet p� rad " + xmlex.LineNumber + " och teckennummer " + xmlex.LinePosition+ ". ", "Felaktig xml", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (Exception ex)
            {
                _Felmeddelande = ex.Message;
            }
        }

        #region get- och setegenskaper
	
        /// <summary>
        /// H�mtar best�llningsdatumet. 
        /// </summary>
		public string Best�llningsdatum
		{
			get
			{
                return _Best�llningsdatum;
            }
		}

        /// <summary>
        /// H�mtar kontaktpersons etts e-postadress. 
        /// </summary>
        public string EpostKontaktperson1
        {
            get
            {
                return _Epost1;
            }
        }

        /// <summary>
        /// H�mtar kontaktpersons tv�s e-postadress. 
        /// </summary>
        public string EpostKontaktperson2
        {
            get
            {
                return _Epost2;
            }
        }

        /// <summary>
        /// H�mtar faktureringsadressen. 
        /// </summary>
        public string Faktureringsadress
        {
            get
            {
                return _Faktureringsadress;
            }
        }

        /// <summary>
        /// H�mtar ett v�rde som talar om om inl�sningen av xml-str�ngen gick d�ligt. 
        /// </summary>
        public static bool FelIXML
        {
            get
            {
                return _FelIXML;
            }
        }

        /// <summary>
        /// H�mtar f�retagsnamnet. 
        /// </summary>
        public string F�retagsnamn
		{
			get
			{
                return _F�retagsnamn;
            }
		}

        /// <summary>
        /// H�mtar postnummret. 
        /// </summary>
        public string Postnummer
		{
			get
			{
                return _Postnummer;
            }
		}

        /// <summary>
        /// H�mtar orten. 
        /// </summary>
        public string Ort
		{
			get
			{
                return _Ort;
            }
		}

        /// <summary>
        /// H�mtar region/f�rvaltningen. 
        /// </summary>
        public string Region_F�rvaltning
		{
			get
			{
                return _Region_F�rvaltning;
            }
		}

        /// <summary>
        /// H�mtar VAT-numret. 
        /// </summary>
        public string VAT
		{
			get
			{
                return _VAT;
            } 
		}

        /// <summary>
        /// H�mtar distrikt/omr�de. 
        /// </summary>
        public string Distrikt_Omr�de
		{
			get
			{
                return _Distrikt_Omr�de;
            } 
		}

        /// <summary>
        /// H�mtar best�llningsreferensen. 
        /// </summary>
        public string Best�llningsreferens
		{
			get
			{
                return _Best�llningsreferens;
            } 
		}

        /// <summary>
        /// H�mtar kommentar
        /// </summary>
        public string Kommentar
        {
            get
            {
                return _Kommentar;
            }
        }

        /// <summary>
        /// H�mtar kontaktperson ett
        /// </summary>
        public string Kontaktperson1
		{
			get
			{
                return _Kontaktperson1;
            } 
		}

        /// <summary>
        /// H�mtar kontaktperson tv�
        /// </summary>
        public string Kontaktperson2
        {
            get
            {
                return _Kontaktperson2;
            }
        }

        /// <summary>
        /// H�mtar kontaktperson etts telefon till arbetet. 
        /// </summary>
        public string TelefonArbeteKontaktperson1
        {
            get
            {
                return _TelefonArb1;
            }
        }

        /// <summary>
        /// H�mtar kontaktperson etts mobiltelefonnummer. 
        /// </summary>
        public string TelefonMobilKontaktperson1
        {
            get
            {
                return _TelefonMob1;
            }
        }

        /// <summary>
        /// H�mtar kontaktperson tv�s mobiltelefonnummer. 
        /// </summary>
        public string TelefonMobilKontaktperson2
        {
            get
            {
                return _TelefonMob2;
            }
        }

        /// <summary>
        /// H�mtar kontaktperson etts telefon till hemmet. 
        /// </summary>
        public string TelefonHemKontaktperson1
        {
            get
            {
                return _TelefonHem1;
            }
        }

        /// <summary>
        /// H�mtar kontaktperson tv�s telefon till hemmet. 
        /// </summary>
        public string TelefonHemKontaktperson2
        {
            get
            {
                return _TelefonHem2;
            }
        }

        /// <summary>
        /// H�mtar tidsst�mpeln f�r denna best�llning. 
        /// </summary>
        public string Tidsst�mpel
        {
            get
            {
                return _Tidsst�mpel;
            }
        }

        #endregion
    }
}
