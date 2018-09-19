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
    /// L�ser information fr�n den inkommande xml-str�ngen f�r startplatstaggen. F�r att skapa ett nytt 
    /// Startplatsobjekt skall man anv�nda sig av metoden ByggUppObjekt(string) eftersom klassen �r skriven enligt
    /// designm�nstret "Factory". 
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
        private static string _FelF�rV�rde;

        #endregion

        /// <summary>
        /// Skapar en ny startplats utifr�n en xml-str�ng. 
        /// </summary>
        protected Startplats()
        {
        }

        /// <summary>
        /// Tar in en xml och tar ut alla startplatser som finns i xml:en. 
        /// </summary>
        /// <param name="xml">Xml-str�ngen med startplatser i. </param>
        /// <returns>Returnerar en lista med startplatser som finns i xml:en. </returns>
        public static List<Startplats> ByggUppObjekt(string xml)
        {
            return L�sIn(xml);
        }

        /// <summary>
        /// Bygger upp ett SQL-kommando utifr�n en startplats.
        /// </summary>
        /// <param name="Ordernummer">Ordernummert som finns i best�llningen. </param>
        /// <returns>Returnerar ett sql-kommando som l�gger in alla uppgifter i en databas. </returns>
        public OleDbCommand ByggUppSQL(string Ordernummer)
        {
            // Kommandot som skall byggas upp
            OleDbCommand kommando = new OleDbCommand();

            // Bygger upp grunden f�r sqkl.kommandot
            string SQLSats = "INSERT INTO Startplats (Ordernr, Startplats, Nordligkoordinat_startplats, ";
            SQLSats += "Ostligkoordinat_startplats, Areal_ha_startplats, Skog_CAN_ton_startplats, Ing�ende_Objekt";
            SQLSats += ") VALUES (";

            // L�gger till alla uppgifter ifr�n best�llningen. 
            SQLSats += "@Ordernummer, @_Startplats, @_Nordligkoordinat, @_Ostligkoordinat, @_Areal, @_CAN, @_Ing�ende_Objekt)";

            // Anger kommandotexten. 
            kommando.CommandText = SQLSats;

            // Anger vilka typer som �r legala i kommandot. 
            kommando.Parameters.Add("@Ordernummer", OleDbType.Integer);
            kommando.Parameters.Add("@_Startplats", OleDbType.VarChar);
            kommando.Parameters.Add("@_Nordligkoordinat", OleDbType.Integer);
            kommando.Parameters.Add("@_Ostligkoordinat", OleDbType.Integer);
            kommando.Parameters.Add("@_Areal", OleDbType.Double);
            kommando.Parameters.Add("@_CAN", OleDbType.Double);
            kommando.Parameters.Add("@_Ing�ende_Objekt", OleDbType.VarChar);

            // L�gger till alla v�rden
            kommando.Parameters[0].Value = Ordernummer;
            kommando.Parameters[1].Value = _Startplats;
            kommando.Parameters[2].Value = _Nordligkoordinat;
            kommando.Parameters[3].Value = _Ostligkoordinat;
            kommando.Parameters[4].Value = _Areal;
            kommando.Parameters[5].Value = _CAN;
            kommando.Parameters[6].Value = H�mtaIng�endeObjekt();

            return kommando;
        }

        /// <summary>
        /// H�mtar alla ing�ende objekt f�r denna startplats, d.v.s. alla objekt som �r kopplade till just denna 
        /// startplats och best�llning och bygger ihop dem till en str�ng enligt m�nstret "1,2,3". 
        /// </summary>
        /// <returns>Returnerar tillbaka en str�ng med alla ing�ende objekt f�r denna startplats. </returns>
        private string H�mtaIng�endeObjekt()
        {
            string ing�endeObjekt = "";

            // Bygger upp raden med alla ing�ende objekt i enligt m�nstret "1,2,3". 
            for (int objektElement = 0; objektElement < _Objekt.Count - 1; objektElement++)
                ing�endeObjekt += _Objekt[objektElement].Objektnummer + ",";
            ing�endeObjekt += _Objekt[_Objekt.Count - 1].Objektnummer;

            return ing�endeObjekt;
        }

        /// <summary>
        /// Tar in en xml-str�ng och fyller p� samtliga instansvariabler i denna klass med information fr�n
        /// den xml:en. 
        /// </summary>
        /// <param name="xml">Inkommande xml-str�ng. </param>
        private static List<Startplats> L�sIn(string xml)
        {
            List<Startplats> startplatser = new List<Startplats>();

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
                XmlNodeList xmlNodeStartplats;
                xmlNodeStartplats = xmlDoc.SelectNodes("Best�llning/child::Startplats");

                // Loopar igenom alla startplatser och l�gger in v�rden fr�n dem. 
                for (int startplatsIndex = 0; startplatsIndex < xmlNodeStartplats.Count; startplatsIndex++)
                {
                    Startplats startplats = new Startplats();

                    // L�gger in startplatsen. 
                    startplats._Startplats = xmlNodeStartplats[startplatsIndex].ChildNodes[0].InnerText;

                    try
                    {
                        // L�gger in den nordliga koordinaten
                        startplats._Nordligkoordinat = 
                            double.Parse(xmlNodeStartplats[startplatsIndex].ChildNodes[1].InnerText, nf);
                    }
                    catch (Exception ex)
                    {
                        _FelIXML = true;
                        _Felmeddelande = ex.Message;
                        _FelF�rV�rde = "nordliga koordinaten";
                    }

                    try
                    {
                        // L�gger in den ostliga koordinaten
                        startplats._Ostligkoordinat = 
                            double.Parse(xmlNodeStartplats[startplatsIndex].ChildNodes[2].InnerText, nf);
                    }
                    catch (Exception ex)
                    {
                        _FelIXML = true;
                        _Felmeddelande = ex.Message;
                        _FelF�rV�rde = "ostliga koordinaten";
                    }

                    try
                    {
                        // L�gger in arealen
                        startplats._Areal = 
                            double.Parse (xmlNodeStartplats[startplatsIndex].ChildNodes[3].InnerText, nf);
                    }
                    catch (Exception ex)
                    {
                        _FelIXML = true;
                        _Felmeddelande = ex.Message;
                        _FelF�rV�rde = "arealen";
                    }

                    try
                    {
                        // L�gger in can
                        startplats._CAN = 
                            double.Parse(xmlNodeStartplats[startplatsIndex].ChildNodes[4].InnerText, nf);
                    }
                    catch (Exception ex)
                    {
                        _FelIXML = true;
                        _Felmeddelande = ex.Message;
                        _FelF�rV�rde = "CAN";
                    }

                    // L�gger in alla objekt som hittas i startplatsen 
                    startplats._Objekt = new List<ObjektIStartplats>();
                    for (int objektIndex = 5; objektIndex < xmlNodeStartplats[startplatsIndex].ChildNodes.Count; objektIndex++)
                        startplats._Objekt.Add(ObjektIStartplats.ByggUppObjekt(
                            xmlNodeStartplats[startplatsIndex].ChildNodes[objektIndex]));

                    // L�gger till den nyss skapade startplatsen
                    startplatser.Add(startplats);
                }
            }
            catch (XmlException xmlex)
            {
                _FelIXML = true;
                _Felmeddelande = xmlex.Message;

                // Meddelare anv�ndaren om detta fel. 
                MessageBox.Show("Xml-str�ngen inneh�ller fel inom en Startplatstagg och kan ej anv�ndas f�r att spara data med. \nLeta efter felet p� rad " + xmlex.LineNumber + " och teckennummer " + xmlex.LinePosition + ". ", "Felaktig xml", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (Exception ex)
            {
                _Felmeddelande = ex.Message;
            }

            return startplatser;
        }

        #region get- och setegenskaper

        /// <summary>
        /// H�mtar areal f�r detta startplatsobjekt. 
        /// </summary>
        public double Areal
        {
            get
            {
                return _Areal;
            }
        }

        /// <summary>
        /// H�mtar can f�r detta startplatsobjekt. 
        /// </summary>
        public double CAN
        {
            get
            {
                return _CAN;
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
        /// H�mtar vad felmeddelandet g�ller f�r n�got. 
        /// </summary>
        public static string Felmeddelande
        {
            get
            {
                return _Felmeddelande;
            }
        }

        /// <summary>
        /// H�mtar vad det felmeddelandet g�ller f�r. 
        /// </summary>
        public static string FelF�rV�rde
        {
            get
            {
                return _FelF�rV�rde;
            }
        }

        /// <summary>
        /// H�mtar eller s�tter nordligkoordinat f�r detta startplatsobjekt. 
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
        /// H�mtar all objekt f�r detta startplatsobjekt. 
        /// </summary>
        public List<ObjektIStartplats> Objekt
        {
            get
            {
                return _Objekt;
            }
        }

        /// <summary>
        /// H�mtar eller s�tter ostligkoordinat f�r detta startplatsobjekt. 
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
        /// H�mtar startplatsen f�r detta startplatsobjekt. 
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
