using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace SG_xml
{
    /// <summary>
    /// Läser information från den inkommande xml-strängen för startplatstaggen. För att skapa ett nytt 
    /// Startplatsobjekt skall man använda sig av metoden ByggUppObjekt(string) eftersom klassen är skriven enligt
    /// designmönstret "Factory". 
    /// 
    /// Skapad av MTTO. 
    /// </summary>
    public class Startplats
    {
        #region instansvariabler

        private string _Startplats;
        private double _Nordligkoordinat;
        private double _Ostligkoordinat;
        private double _Areal;
        private double _CAN;
        private List<ObjektIStartplats> _Objekt;

        private static bool _FelIXML = false; 
        private static string _Felmeddelande;
        private static string _FelFörVärde;

        #endregion

        /// <summary>
        /// Skapar en ny startplats utifrån en xml-sträng. 
        /// </summary>
        protected Startplats()
        {
        }

        /// <summary>
        /// Tar in en xml och tar ut alla startplatser som finns i xml:en. 
        /// </summary>
        /// <param name="xml">Xml-strängen med startplatser i. </param>
        /// <returns>Returnerar en lista med startplatser som finns i xml:en. </returns>
        public static List<Startplats> ByggUppObjekt(string xml)
        {
            return LäsIn(xml);
        }

        /// <summary>
        /// Bygger upp ett SQL-kommando utifrån en startplats.
        /// </summary>
        /// <param name="Ordernummer">Ordernummert som finns i beställningen. </param>
        /// <returns>Returnerar ett sql-kommando som lägger in alla uppgifter i en databas. </returns>
        public OleDbCommand ByggUppSQL(string Ordernummer)
        {
            // Kommandot som skall byggas upp
            OleDbCommand kommando = new OleDbCommand();

            // Bygger upp grunden för sqkl.kommandot
            string SQLSats = "INSERT INTO Startplats (Ordernr, Startplats, Nordligkoordinat_startplats, ";
            SQLSats += "Ostligkoordinat_startplats, Areal_ha_startplats, Skog_CAN_ton_startplats, Ingående_Objekt";
            SQLSats += ") VALUES (";

            // Lägger till alla uppgifter ifrån beställningen. 
            SQLSats += "@Ordernummer, @_Startplats, @_Nordligkoordinat, @_Ostligkoordinat, @_Areal, @_CAN, @_Ingående_Objekt)";

            // Anger kommandotexten. 
            kommando.CommandText = SQLSats;

            // Anger vilka typer som är legala i kommandot. 
            kommando.Parameters.Add("@Ordernummer", OleDbType.Integer);
            kommando.Parameters.Add("@_Startplats", OleDbType.VarChar);
            kommando.Parameters.Add("@_Nordligkoordinat", OleDbType.Integer);
            kommando.Parameters.Add("@_Ostligkoordinat", OleDbType.Integer);
            kommando.Parameters.Add("@_Areal", OleDbType.Double);
            kommando.Parameters.Add("@_CAN", OleDbType.Double);
            kommando.Parameters.Add("@_Ingående_Objekt", OleDbType.VarChar);

            // Lägger till alla värden
            kommando.Parameters[0].Value = Ordernummer;
            kommando.Parameters[1].Value = _Startplats;
            kommando.Parameters[2].Value = _Nordligkoordinat;
            kommando.Parameters[3].Value = _Ostligkoordinat;
            kommando.Parameters[4].Value = _Areal;
            kommando.Parameters[5].Value = _CAN;
            kommando.Parameters[6].Value = HämtaIngåendeObjekt();

            return kommando;
        }

        /// <summary>
        /// Hämtar alla ingående objekt för denna startplats, d.v.s. alla objekt som är kopplade till just denna 
        /// startplats och beställning och bygger ihop dem till en sträng enligt mönstret "1,2,3". 
        /// </summary>
        /// <returns>Returnerar tillbaka en sträng med alla ingående objekt för denna startplats. </returns>
        private string HämtaIngåendeObjekt()
        {
            string ingåendeObjekt = "";

            // Bygger upp raden med alla ingående objekt i enligt mönstret "1,2,3". 
            for (int objektElement = 0; objektElement < _Objekt.Count - 1; objektElement++)
                ingåendeObjekt += _Objekt[objektElement].Objektnummer + ",";
            ingåendeObjekt += _Objekt[_Objekt.Count - 1].Objektnummer;

            return ingåendeObjekt;
        }

        /// <summary>
        /// Tar in en xml-sträng och fyller på samtliga instansvariabler i denna klass med information från
        /// den xml:en. 
        /// </summary>
        /// <param name="xml">Inkommande xml-sträng. </param>
        private static List<Startplats> LäsIn(string xml)
        {
            List<Startplats> startplatser = new List<Startplats>();

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
                XmlNodeList xmlNodeStartplats;
                xmlNodeStartplats = xmlDoc.SelectNodes("Beställning/child::Startplats");

                // Loopar igenom alla startplatser och lägger in värden från dem. 
                for (int startplatsIndex = 0; startplatsIndex < xmlNodeStartplats.Count; startplatsIndex++)
                {
                    Startplats startplats = new Startplats();

                    // Lägger in startplatsen. 
                    startplats._Startplats = xmlNodeStartplats[startplatsIndex].ChildNodes[0].InnerText;

                    try
                    {
                        // Lägger in den nordliga koordinaten
                        startplats._Nordligkoordinat = 
                            double.Parse(xmlNodeStartplats[startplatsIndex].ChildNodes[1].InnerText, nf);
                    }
                    catch (Exception ex)
                    {
                        _FelIXML = true;
                        _Felmeddelande = ex.Message;
                        _FelFörVärde = "nordliga koordinaten";
                    }

                    try
                    {
                        // Lägger in den ostliga koordinaten
                        startplats._Ostligkoordinat = 
                            double.Parse(xmlNodeStartplats[startplatsIndex].ChildNodes[2].InnerText, nf);
                    }
                    catch (Exception ex)
                    {
                        _FelIXML = true;
                        _Felmeddelande = ex.Message;
                        _FelFörVärde = "ostliga koordinaten";
                    }

                    try
                    {
                        // Lägger in arealen
                        startplats._Areal = 
                            double.Parse (xmlNodeStartplats[startplatsIndex].ChildNodes[3].InnerText, nf);
                    }
                    catch (Exception ex)
                    {
                        _FelIXML = true;
                        _Felmeddelande = ex.Message;
                        _FelFörVärde = "arealen";
                    }

                    try
                    {
                        // Lägger in can
                        startplats._CAN = 
                            double.Parse(xmlNodeStartplats[startplatsIndex].ChildNodes[4].InnerText, nf);
                    }
                    catch (Exception ex)
                    {
                        _FelIXML = true;
                        _Felmeddelande = ex.Message;
                        _FelFörVärde = "CAN";
                    }

                    // Lägger in alla objekt som hittas i startplatsen 
                    startplats._Objekt = new List<ObjektIStartplats>();
                    for (int objektIndex = 5; objektIndex < xmlNodeStartplats[startplatsIndex].ChildNodes.Count; objektIndex++)
                        startplats._Objekt.Add(ObjektIStartplats.ByggUppObjekt(
                            xmlNodeStartplats[startplatsIndex].ChildNodes[objektIndex]));

                    // Lägger till den nyss skapade startplatsen
                    startplatser.Add(startplats);
                }
            }
            catch (XmlException xmlex)
            {
                _FelIXML = true;
                _Felmeddelande = xmlex.Message;

                // Meddelare användaren om detta fel. 
                MessageBox.Show("Xml-strängen innehåller fel inom en Startplatstagg och kan ej användas för att spara data med. \nLeta efter felet på rad " + xmlex.LineNumber + " och teckennummer " + xmlex.LinePosition + ". ", "Felaktig xml", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (Exception ex)
            {
                _Felmeddelande = ex.Message;
            }

            return startplatser;
        }

        #region get- och setegenskaper

        /// <summary>
        /// Hämtar areal för detta startplatsobjekt. 
        /// </summary>
        public double Areal
        {
            get
            {
                return _Areal;
            }
        }

        /// <summary>
        /// Hämtar can för detta startplatsobjekt. 
        /// </summary>
        public double CAN
        {
            get
            {
                return _CAN;
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
        /// Hämtar vad felmeddelandet gäller för något. 
        /// </summary>
        public static string Felmeddelande
        {
            get
            {
                return _Felmeddelande;
            }
        }

        /// <summary>
        /// Hämtar vad det felmeddelandet gäller för. 
        /// </summary>
        public static string FelFörVärde
        {
            get
            {
                return _FelFörVärde;
            }
        }

        /// <summary>
        /// Hämtar eller sätter nordligkoordinat för detta startplatsobjekt. 
        /// </summary>
        public double Nordligkoordinat
        {
            get
            {
                return _Nordligkoordinat;
            }
            set
            {
                _Nordligkoordinat = value;
            }
        }

        /// <summary>
        /// Hämtar all objekt för detta startplatsobjekt. 
        /// </summary>
        public List<ObjektIStartplats> Objekt
        {
            get
            {
                return _Objekt;
            }
        }

        /// <summary>
        /// Hämtar eller sätter ostligkoordinat för detta startplatsobjekt. 
        /// </summary>
        public double Ostligkoordinat
        {
            get
            {
                return _Ostligkoordinat;
            }
            set
            {
                _Ostligkoordinat = value;
            }
        }

        /// <summary>
        /// Hämtar startplatsen för detta startplatsobjekt. 
        /// </summary>
        public string StartPlats
        {
            get
            {
                return _Startplats;
            }
        }

        #endregion
    }
}
