using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace SG_xml
{
    /// <summary>
    /// Läser information från den inkommande xml-strängen för startplatstaggen. För att skapa ett nytt
    /// Reservobjekt skall man använda sig av metoden ByggUppObjekt(string) eftersom klassen är skriven enligt
    /// designmönstret "Factory". 
    /// 
    /// Skapad av MTTO. 
    /// </summary>
    public class ReservObjekt
    {
        #region instansvariabler

        private string _Objektnummer;
        private string _Avdelningsnummer;
        private string _Avdelningsnamn;
        private double _Areal;
        private double _Giva;
        private string _Kommentar;

        private static string _Felmeddelande;
        private static bool _FelIXML = false;

        #endregion

        /// <summary>
        /// Skapar ett nytt reservobjekt. 
        /// </summary>
        protected ReservObjekt()
        {
        }

        /// <summary>
        /// Tar in en xml och tar ut alla reservobjekt som finns i xml:en. 
        /// </summary>
        /// <param name="xml">Xml-strängen med startplatser i. </param>
        /// <returns>Returnerar en lista med startplatser som finns i xml:en. </returns>
        public static List<ReservObjekt> ByggUppObjekt(string xml)
        {
            return LäsIn(xml);
        }

        /// <summary>
        /// Bygger upp ett SQL-kommando utifrån en reservstartplats.
        /// </summary>
        /// <param name="Ordernummer">Ordernummret kopplat till detta reservobjekt. </param>
        /// <returns>Returnerar ett sql-kommando som lägger in alla uppgifter i en databas. </returns>
        public OleDbCommand ByggUppSQL(string Ordernummer)
        {
            // Kommandot som skall byggas upp
            OleDbCommand kommando = new OleDbCommand();

            // Grunden i sql-satsen. 
            string SQLSats = "INSERT INTO Reservobjekt (Ordernr, Objektnr, Avdnr, Avdnamn, Areal_ha, Giva_KgN_ha, ";
            SQLSats += "Kommentar) VALUES (";

            // Lägger till alla uppgifter ifrån beställningen. 
            SQLSats += "@Ordernummer, @_Objektnummer, @_Avdelningsnummer, @_Avdelningsnamn, @_Areal, @_Giva, @_Kommentar)";

            // Anger kommandotexten. 
            kommando.CommandText = SQLSats;

            // Anger vilka typer som är legala i kommandot. 
            kommando.Parameters.Add("@Ordernummer", OleDbType.Integer);
            kommando.Parameters.Add("@_Objektnummer", OleDbType.VarChar);
            kommando.Parameters.Add("@_Avdelningsnummer", OleDbType.VarChar);
            kommando.Parameters.Add("@_Avdelningsnamn", OleDbType.VarChar);
            kommando.Parameters.Add("@_Areal", OleDbType.Double);
            kommando.Parameters.Add("@_Giva", OleDbType.Double);
            kommando.Parameters.Add("@_Kommentar", OleDbType.VarChar);

            // Lägger till alla värden
            kommando.Parameters[0].Value = Ordernummer;
            kommando.Parameters[1].Value = _Objektnummer;
            kommando.Parameters[2].Value = _Avdelningsnummer;
            kommando.Parameters[3].Value = _Avdelningsnamn;
            kommando.Parameters[4].Value = _Areal;
            kommando.Parameters[5].Value = _Giva;
            kommando.Parameters[6].Value = _Kommentar;

            return kommando;
        }


        /// <summary>
        /// Tar in en xml-sträng och fyller på samtliga instansvariabler i denna klass med information från
        /// den xml:en. 
        /// </summary>
        /// <param name="xml">Inkommande xml-sträng. </param>
        private static List<ReservObjekt> LäsIn(string xml)
        {
            List<ReservObjekt> reservobjektlista = new List<ReservObjekt>();

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
                XmlNodeList xmlNodeReservobjekt;
                xmlNodeReservobjekt = xmlDoc.SelectNodes("Beställning/child::Reservobjekt");

                // Loopar igenom alla reservobjekt och lägger in värden från dem. 
                for (int reservsobjektsIndex = 0; reservsobjektsIndex < xmlNodeReservobjekt.Count; reservsobjektsIndex++)
                {
                    ReservObjekt reservobjekt = new ReservObjekt();

                    // Lägger in objektnummret
                    reservobjekt._Objektnummer = xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[0].InnerText;

                    // Lägger in avdelningsnummret
                    reservobjekt._Avdelningsnummer = xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[1].InnerText;

                    // Lägger in avdelningsnamnet
                    reservobjekt._Avdelningsnamn = xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[2].InnerText;

                    // Lägger in arealen
                    reservobjekt._Areal = 
                        double.Parse(xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[3].InnerText, nf);

                    // Lägger in giva
                    reservobjekt._Giva = 
                        double.Parse(xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[4].InnerText, nf);

                    // Lägger in kommentaren
                    reservobjekt._Kommentar = xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[5].InnerText;

                    reservobjektlista.Add(reservobjekt);
                }
            }
            catch (XmlException xmlex)
            {
                _FelIXML = true;
                _Felmeddelande = xmlex.Message;

                // Meddelare användaren om detta fel. 
                MessageBox.Show("Xml-strängen innehåller fel inom en Reservobjektstagg och kan ej användas för att spara data med. \nLeta efter felet på rad " + xmlex.LineNumber + " och teckennummer " + xmlex.LinePosition+ ". ", "Felaktig xml", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (Exception ex)
            {
                _Felmeddelande = ex.Message;
            }

            return reservobjektlista;
        }

        #region get- och setegeneskaper

        /// <summary>
        /// Hämatar areal per hektar (ha). 
        /// </summary>
        public double Areal
        {
            get
            {
                return _Areal;
            }
        }

        /// <summary>
        /// Hämatar avdelningsnamn. 
        /// </summary>
        public string Avdelningsnamn
        {
            get
            {
                return _Avdelningsnamn;
            }
        }

        /// <summary>
        /// Hämtar avdelningsnummret. 
        /// </summary>
        public string Avdelningsnummer
        {
            get
            {
                return _Avdelningsnummer;
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
        /// Hämtar giva kilo kväve per hektar (kgN/ha). 
        /// </summary>
        public double Giva
        {
            get
            {
                return _Giva;
            }
        }

        /// <summary>
        /// Hämtar en kommentar. 
        /// </summary>
        public string Kommentar
        {
            get
            {
                return _Kommentar;
            }
        }

        /// <summary>
        /// Hämtar objektnummret. 
        /// </summary>
        public string Objektnummer
        {
            get
            {
                return _Objektnummer;
            }
        }

        #endregion
    }
}
