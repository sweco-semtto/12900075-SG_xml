/****************************************************
 * ImageRequest.cs läser ur information från en 
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
	/// Läser information från den inkommande xml-strängen för företagstaggen. För att skapa ett nytt Förtagsobjekt
    /// skall man använda sig av metoden ByggUppObjekt(string) eftersom klassen är skriven enligt designmönstret
    /// "Factory". 
    /// 
    /// Skapad av MTTO. 
	/// </summary>
	public class Företag
    {
        #region instansvariabler

        private string _Tidsstämpel;
        private string _Ordernummer;
        private string _Beställningsdatum;
		private string _Företagsnamn;
		private string _Faktureringsadress;
		private string _Postnummer;
		private string _Ort;
		private string _Region_Förvaltning;
		private string _VAT;
		private string _Distrikt_Område;
		private string _Beställningsreferens;
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
        /// Skapar ett nytt Företagsobjekt. 
        /// </summary>
        /// <param name="xml">Den xml-sträng som bygger upp objektet. </param>
        protected Företag(string xml)
		{
			LäsIn(xml);
		}

        public static Företag ByggUppObjekt(string xml)
        {
            Företag nyttObjekt = new Företag(xml);

            return nyttObjekt;
        }

        /// <summary>
        /// Bygger upp ett SQL-kommando (OleDbCommand) utifrån en beställning. 
        /// </summary>
        /// <returns>Returnerar ett sql-kommando som lägger in alla uppgifter i en databas. </returns>
        public OleDbCommand ByggUppSQL()
        {
            // Kommandot som skall byggas upp
            OleDbCommand kommando = new OleDbCommand();
         
            // Bygger upp grunden för sqkl.kommandot
            string SQLSats = "INSERT INTO Företag (Ordernr, Beställningsreferens, Beställningsdatum, Tidsstämpel, ";
            SQLSats += "Företagsnamn, Faktureringsadress, Postnummer, Ort, Region_Förvaltning, Distrikt_Område, ";
            SQLSats += "VAT, Kontaktperson1, TelefonArb1, TelefonMobil1, TelefonHem1, Epostadress1, Kontaktperson2, ";
            SQLSats += "TelefonArb2, TelefonMobil2, TelefonHem2, Epostadress2, Kommentar) VALUES (";

            // Lägger till alla uppgifter ifrån beställningen. 
            SQLSats += "@_Ordernummer, @_Beställningsreferens, @_Beställningsdatum, @_Tidsstämpel, @_Företagsnamn, ";
            SQLSats += "@_Faktureringsadress, @_Postnummer, @_Ort, @_Region_Förvaltning, @_Distrikt_Område, ";
            SQLSats += "@_VAT, @_Kontaktperson1, @_TelefonArb1, @_TelefonMob1, ";
            SQLSats += "@_TelefonHem1, @_Epost1, @_Kontaktperson2, ";
            SQLSats += "@_TelefonArb2, @_TelefonMob2, @_TelefonHem2, @_Epost2, @_Kommentar)";

            // Anger kommandotexten. 
            kommando.CommandText = SQLSats;

            // Anger vilka typer som är legala i kommandot. 
            kommando.Parameters.Add("@_Ordernummer", OleDbType.VarChar);
            kommando.Parameters.Add("@_Beställningsreferens", OleDbType.VarChar);
            kommando.Parameters.Add("@_Beställningsdatum", OleDbType.Date);
            kommando.Parameters.Add("@_Tidsstämpel", OleDbType.VarChar);
            kommando.Parameters.Add("@_Företagsnamn", OleDbType.VarChar);
            kommando.Parameters.Add("@_Faktureringsadress", OleDbType.VarChar);
            kommando.Parameters.Add("@_Postnummer", OleDbType.VarChar);
            kommando.Parameters.Add("@_Ort", OleDbType.VarChar);
            kommando.Parameters.Add("@_Region_Förvaltning", OleDbType.VarChar);
            kommando.Parameters.Add("@_Distrikt_Område", OleDbType.VarChar);
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

            // Lägger till alla värden
            kommando.Parameters[0].Value = _Ordernummer;
            if (_Beställningsreferens.Equals(String.Empty))
                kommando.Parameters[1].Value = _Ordernummer;
            else
                kommando.Parameters[1].Value = _Beställningsreferens;
            kommando.Parameters[2].Value = _Beställningsdatum;
            kommando.Parameters[3].Value = _Tidsstämpel;
            kommando.Parameters[4].Value = _Företagsnamn;
            kommando.Parameters[5].Value = _Faktureringsadress;
            kommando.Parameters[6].Value = _Postnummer;
            kommando.Parameters[7].Value = _Ort;
            kommando.Parameters[8].Value = _Region_Förvaltning;
            kommando.Parameters[9].Value = _Distrikt_Område;
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
        /// Bygger upp ett SQL-kommando (OleDbCommand) för att uppdatera ordernummret. Anledningen till detta är att 
        /// alla kartprogram inte kan läsa räknaren för ordernummret. 
        /// </summary>
        /// <returns>Returnerar ett sql-kommando som uppdaterar ordernummret i databasen.  </returns>
        public OleDbCommand ByggUppSQLLäggTillOrdernummer(string ordernummer)
        {
            // Kommandot som skall byggas upp
            OleDbCommand kommando = new OleDbCommand();

            // Bygger upp grunden för sqkl.kommandot
            string SQLSats = "UPDATE Företag SET OrdernrText=@_Beställningsreferens WHERE Ordernr=@_Beställningsreferens";

            // Anger kommandotexten. 
            kommando.CommandText = SQLSats;

            // Anger vilka typer som är legala i kommandot. 
            kommando.Parameters.Add("@_Beställningsreferens", OleDbType.VarChar);

            // Lägger till alla värden
            kommando.Parameters[0].Value = ordernummer;

            return kommando;
        }

        /// <summary>
        /// Bygger upp en SQL-fråga som hämtar ordernummer. 
        /// </summary>
        /// <param name="funktionAttLäsaInMed">Funktion att läsa med ifrån databasen</param>
        /// <returns>Returnerar en sträng som hittar ordernummret till denna order. </returns>
        public string HämtaOrdernummer(LäsIfrånDatabas funktionAttLäsaInMed)
        {
            string SQLFråga = "SELECT * FROM Företag where Tidsstämpel = '" + _Tidsstämpel + "'";

            // Hämtar ordernummret för att kunna skriva in startplatsen ifrån databasen. 
            DataSet data = funktionAttLäsaInMed(SQLFråga);
            
            // Tar fram ordernummret från datasetet. 
            try
            {
                // Tar fram värdet i rad ett kolumn ett (d.v.s. ordernummer). 
                return data.Tables[0].Rows[0][1].ToString();
            }
            // Inget fel borde inträffa, om det går det är något alvarligt fel. 
            catch (Exception)
            {
                MessageBox.Show("Felaktig data i databasen, vänligen kontakta Sweco med detta meddelande. ", "Felaktigt värde i databasen", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

            return null;
        }
		
		/// <summary>
		/// Tar in en xml-sträng och fyller på samtliga instansvariabler i denna klass med information från
        /// den xml:en. 
		/// </summary>
		/// <param name="xml">Inkommande xml-sträng. </param>
        private void LäsIn(string xml)
		{
            try
            {
                // Anger var siffror har för kommaseparerare. 
                NumberFormatInfo nf = new NumberFormatInfo();
                nf.NumberDecimalSeparator = ".";

                // Laddar in xml-dokumentet. 
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("ogc", "http://www.opengis.net/ogc");

                // noden som används för att läsa in de olika uppgifterna. 
                XmlNode xmlNode;

                // Läser in ordernummer. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Ordernummer");
                _Ordernummer = xmlNode != null ? xmlNode.InnerText : MySqlCommunicator.GetNewOrdernumber();

                // Läser in beställningsdatum. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Beställningsdatum");
                _Beställningsdatum = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in företagsnman.
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Företagsnamn");
                _Företagsnamn = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in faktureringsadress.
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Faktureringsadress");
                _Faktureringsadress = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Lässer in postnummer.
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Postnummer");
                _Postnummer = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in ort. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Ort");
                _Ort = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in region/förvaltning
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Region_Förvaltning");
                _Region_Förvaltning = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in VAT
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::VAT");
                _VAT = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Lässer in distrikt/område
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Distrikt_Område");
                _Distrikt_Område = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in beställningsreferens
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Beställningsreferens");
                _Beställningsreferens = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in kontaktperson ett. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Kontaktperson1");
                _Kontaktperson1 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in kontaktperson etts telefon till arbetet. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Telefon_Arb1");
                _TelefonArb1 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in kontaktperson etts mobiltelefonnummer. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Telefon_Mob1");
                _TelefonMob1 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in kontaktperson etts telefon till hemmet. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Telefon_Hem1");
                _TelefonHem1 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in kontaktperson etts e-postadress. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::E-postadress1");
                _Epost1 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in kontaktperson två. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Kontaktperson2");
                _Kontaktperson2 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in kontaktperson tvås telefon till arbetet. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Telefon_Arb2");
                _TelefonArb2 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in kontaktperson tvås mobiltelefonnummer. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Telefon_Mob2");
                _TelefonMob2 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in kontaktperson tvås telefon till hemmet. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::Telefon_Hem2");
                _TelefonHem2 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in kontaktperson tvås e-postadress. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Företag/child::E-postadress2");
                _Epost2 = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Läser in kommentaren sist i xml:en. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Kommentar");
                _Kommentar = xmlNode != null ? xmlNode.InnerText : string.Empty;

                // Lägger in en tidsstämpel när allt inläst ifrån xml:en.
                _Tidsstämpel = DateTime.Now.Ticks.ToString();
            }
            catch (XmlException xmlex)
            {
                _FelIXML = true;
                _Felmeddelande = xmlex.Message;

                // Meddelare användaren om detta fel. 
                MessageBox.Show("Xml-strängen innehåller fel inom Företagstaggen och kan ej användas för att spara data med. \nLeta efter felet på rad " + xmlex.LineNumber + " och teckennummer " + xmlex.LinePosition+ ". ", "Felaktig xml", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (Exception ex)
            {
                _Felmeddelande = ex.Message;
            }
        }

        #region get- och setegenskaper
	
        /// <summary>
        /// Hämtar beställningsdatumet. 
        /// </summary>
		public string Beställningsdatum
		{
			get
			{
                return _Beställningsdatum;
            }
		}

        /// <summary>
        /// Hämtar kontaktpersons etts e-postadress. 
        /// </summary>
        public string EpostKontaktperson1
        {
            get
            {
                return _Epost1;
            }
        }

        /// <summary>
        /// Hämtar kontaktpersons tvås e-postadress. 
        /// </summary>
        public string EpostKontaktperson2
        {
            get
            {
                return _Epost2;
            }
        }

        /// <summary>
        /// Hämtar faktureringsadressen. 
        /// </summary>
        public string Faktureringsadress
        {
            get
            {
                return _Faktureringsadress;
            }
        }

        /// <summary>
        /// Hämtar ett värde som talar om om inläsningen av xml-strängen gick dåligt. 
        /// </summary>
        public static bool FelIXML
        {
            get
            {
                return _FelIXML;
            }
        }

        /// <summary>
        /// Hämtar företagsnamnet. 
        /// </summary>
        public string Företagsnamn
		{
			get
			{
                return _Företagsnamn;
            }
		}

        /// <summary>
        /// Hämtar postnummret. 
        /// </summary>
        public string Postnummer
		{
			get
			{
                return _Postnummer;
            }
		}

        /// <summary>
        /// Hämtar orten. 
        /// </summary>
        public string Ort
		{
			get
			{
                return _Ort;
            }
		}

        /// <summary>
        /// Hämtar region/förvaltningen. 
        /// </summary>
        public string Region_Förvaltning
		{
			get
			{
                return _Region_Förvaltning;
            }
		}

        /// <summary>
        /// Hämtar VAT-numret. 
        /// </summary>
        public string VAT
		{
			get
			{
                return _VAT;
            } 
		}

        /// <summary>
        /// Hämtar distrikt/område. 
        /// </summary>
        public string Distrikt_Område
		{
			get
			{
                return _Distrikt_Område;
            } 
		}

        /// <summary>
        /// Hämtar beställningsreferensen. 
        /// </summary>
        public string Beställningsreferens
		{
			get
			{
                return _Beställningsreferens;
            } 
		}

        /// <summary>
        /// Hämtar kommentar
        /// </summary>
        public string Kommentar
        {
            get
            {
                return _Kommentar;
            }
        }

        /// <summary>
        /// Hämtar kontaktperson ett
        /// </summary>
        public string Kontaktperson1
		{
			get
			{
                return _Kontaktperson1;
            } 
		}

        /// <summary>
        /// Hämtar kontaktperson två
        /// </summary>
        public string Kontaktperson2
        {
            get
            {
                return _Kontaktperson2;
            }
        }

        /// <summary>
        /// Hämtar kontaktperson etts telefon till arbetet. 
        /// </summary>
        public string TelefonArbeteKontaktperson1
        {
            get
            {
                return _TelefonArb1;
            }
        }

        /// <summary>
        /// Hämtar kontaktperson etts mobiltelefonnummer. 
        /// </summary>
        public string TelefonMobilKontaktperson1
        {
            get
            {
                return _TelefonMob1;
            }
        }

        /// <summary>
        /// Hämtar kontaktperson tvås mobiltelefonnummer. 
        /// </summary>
        public string TelefonMobilKontaktperson2
        {
            get
            {
                return _TelefonMob2;
            }
        }

        /// <summary>
        /// Hämtar kontaktperson etts telefon till hemmet. 
        /// </summary>
        public string TelefonHemKontaktperson1
        {
            get
            {
                return _TelefonHem1;
            }
        }

        /// <summary>
        /// Hämtar kontaktperson tvås telefon till hemmet. 
        /// </summary>
        public string TelefonHemKontaktperson2
        {
            get
            {
                return _TelefonHem2;
            }
        }

        /// <summary>
        /// Hämtar tidsstämpeln för denna beställning. 
        /// </summary>
        public string Tidsstämpel
        {
            get
            {
                return _Tidsstämpel;
            }
        }

        #endregion
    }
}
