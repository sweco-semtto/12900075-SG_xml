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
    /// L�ser information fr�n den inkommande xml-str�ngen f�r startplatstaggen. F�r att skapa ett nytt
    /// Reservobjekt skall man anv�nda sig av metoden ByggUppObjekt(string) eftersom klassen �r skriven enligt
    /// designm�nstret "Factory". 
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
        /// <param name="xml">Xml-str�ngen med startplatser i. </param>
        /// <returns>Returnerar en lista med startplatser som finns i xml:en. </returns>
        public static List<ReservObjekt> ByggUppObjekt(string xml)
        {
            return L�sIn(xml);
        }

        /// <summary>
        /// Bygger upp ett SQL-kommando utifr�n en reservstartplats.
        /// </summary>
        /// <param name="Ordernummer">Ordernummret kopplat till detta reservobjekt. </param>
        /// <returns>Returnerar ett sql-kommando som l�gger in alla uppgifter i en databas. </returns>
        public OleDbCommand ByggUppSQL(string Ordernummer)
        {
            // Kommandot som skall byggas upp
            OleDbCommand kommando = new OleDbCommand();

            // Grunden i sql-satsen. 
            string SQLSats = "INSERT INTO Reservobjekt (Ordernr, Objektnr, Avdnr, Avdnamn, Areal_ha, Giva_KgN_ha, ";
            SQLSats += "Kommentar) VALUES (";

            // L�gger till alla uppgifter ifr�n best�llningen. 
            SQLSats += "@Ordernummer, @_Objektnummer, @_Avdelningsnummer, @_Avdelningsnamn, @_Areal, @_Giva, @_Kommentar)";

            // Anger kommandotexten. 
            kommando.CommandText = SQLSats;

            // Anger vilka typer som �r legala i kommandot. 
            kommando.Parameters.Add("@Ordernummer", OleDbType.Integer);
            kommando.Parameters.Add("@_Objektnummer", OleDbType.VarChar);
            kommando.Parameters.Add("@_Avdelningsnummer", OleDbType.VarChar);
            kommando.Parameters.Add("@_Avdelningsnamn", OleDbType.VarChar);
            kommando.Parameters.Add("@_Areal", OleDbType.Double);
            kommando.Parameters.Add("@_Giva", OleDbType.Double);
            kommando.Parameters.Add("@_Kommentar", OleDbType.VarChar);

            // L�gger till alla v�rden
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
        /// Tar in en xml-str�ng och fyller p� samtliga instansvariabler i denna klass med information fr�n
        /// den xml:en. 
        /// </summary>
        /// <param name="xml">Inkommande xml-str�ng. </param>
        private static List<ReservObjekt> L�sIn(string xml)
        {
            List<ReservObjekt> reservobjektlista = new List<ReservObjekt>();

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
                XmlNodeList xmlNodeReservobjekt;
                xmlNodeReservobjekt = xmlDoc.SelectNodes("Best�llning/child::Reservobjekt");

                // Loopar igenom alla reservobjekt och l�gger in v�rden fr�n dem. 
                for (int reservsobjektsIndex = 0; reservsobjektsIndex < xmlNodeReservobjekt.Count; reservsobjektsIndex++)
                {
                    ReservObjekt reservobjekt = new ReservObjekt();

                    // L�gger in objektnummret
                    reservobjekt._Objektnummer = xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[0].InnerText;

                    // L�gger in avdelningsnummret
                    reservobjekt._Avdelningsnummer = xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[1].InnerText;

                    // L�gger in avdelningsnamnet
                    reservobjekt._Avdelningsnamn = xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[2].InnerText;

                    // L�gger in arealen
                    reservobjekt._Areal = 
                        double.Parse(xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[3].InnerText, nf);

                    // L�gger in giva
                    reservobjekt._Giva = 
                        double.Parse(xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[4].InnerText, nf);

                    // L�gger in kommentaren
                    reservobjekt._Kommentar = xmlNodeReservobjekt[reservsobjektsIndex].ChildNodes[5].InnerText;

                    reservobjektlista.Add(reservobjekt);
                }
            }
            catch (XmlException xmlex)
            {
                _FelIXML = true;
                _Felmeddelande = xmlex.Message;

                // Meddelare anv�ndaren om detta fel. 
                MessageBox.Show("Xml-str�ngen inneh�ller fel inom en Reservobjektstagg och kan ej anv�ndas f�r att spara data med. \nLeta efter felet p� rad " + xmlex.LineNumber + " och teckennummer " + xmlex.LinePosition+ ". ", "Felaktig xml", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (Exception ex)
            {
                _Felmeddelande = ex.Message;
            }

            return reservobjektlista;
        }

        #region get- och setegeneskaper

        /// <summary>
        /// H�matar areal per hektar (ha). 
        /// </summary>
        public double Areal
        {
            get
            {
                return _Areal;
            }
        }

        /// <summary>
        /// H�matar avdelningsnamn. 
        /// </summary>
        public string Avdelningsnamn
        {
            get
            {
                return _Avdelningsnamn;
            }
        }

        /// <summary>
        /// H�mtar avdelningsnummret. 
        /// </summary>
        public string Avdelningsnummer
        {
            get
            {
                return _Avdelningsnummer;
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
        /// H�mtar giva kilo kv�ve per hektar (kgN/ha). 
        /// </summary>
        public double Giva
        {
            get
            {
                return _Giva;
            }
        }

        /// <summary>
        /// H�mtar en kommentar. 
        /// </summary>
        public string Kommentar
        {
            get
            {
                return _Kommentar;
            }
        }

        /// <summary>
        /// H�mtar objektnummret. 
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
