using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.Text;
using System.Xml;

namespace SG_xml
{
        /// <summary>
        /// L�ser information fr�n den inkommande xml-noden f�r startplatstaggen och tar ut objektdelen. F�r att
        /// skapa ett nytt ObejktIStartplats skall man anv�nda sig av metoden ByggUppObjekt(XmlNode) eftersom
        /// klassen anv�nder sig av designm�nstret "Factory". 
        /// 
        /// Skapad av MTTO. 
        /// </summary>
        public class ObjektIStartplats
        {
            #region instansvariabler

            private string _Objektnummer;
            private string _Avdelningsnummer;
            private string _Avdelningsnamn;
            private double _Areal;
            private double _Giva;
            private double _CAN;
            private string _Kommentar;

            private static string _Felmeddelande;
            private static bool _FelIXML = false;

            #endregion

            /// <summary>
            /// Skapar ett nytt objekt. 
            /// </summary>
            protected ObjektIStartplats()
            {
            }

            /// <summary>
            /// Bygger upp ett nytt objekt utifr�n en xml-node. 
            /// </summary>
            /// <param name="NodeObjekt">Xml-noden som har informationen om objektet. </param>
            /// <returns>Returnerar ett objekt i en startplats. </returns>
            public static ObjektIStartplats ByggUppObjekt(XmlNode NodeObjekt)
            {
                ObjektIStartplats nyttObjekt = new ObjektIStartplats();

                // Anger var siffror har f�r kommaseparerare. 
                NumberFormatInfo nf = new NumberFormatInfo();
                nf.NumberDecimalSeparator = ".";

                // L�gger in all uppgifter p� objektet. 
                nyttObjekt._Objektnummer = NodeObjekt.ChildNodes[0].InnerText;
                nyttObjekt._Avdelningsnummer = NodeObjekt.ChildNodes[1].InnerText;
                nyttObjekt._Avdelningsnamn = NodeObjekt.ChildNodes[2].InnerText;
                try
                {
                    nyttObjekt._Areal = double.Parse(NodeObjekt.ChildNodes[3].InnerText, nf);
                    nyttObjekt._Giva = double.Parse(NodeObjekt.ChildNodes[4].InnerText, nf);
                    nyttObjekt._CAN = double.Parse(NodeObjekt.ChildNodes[5].InnerText, nf);
                }
                catch (Exception ex)
                {
                    _FelIXML = true;
                    _Felmeddelande = ex.Message;
                }
                nyttObjekt._Kommentar = NodeObjekt.ChildNodes[6].InnerText;

                return nyttObjekt;
            }

            /// <summary>
            /// Bygger upp ett SQL-kommando utifr�n en startplats.
            /// </summary>
            /// <param name="Startplats">Startplatsen f�r detta objekt. </param>
            /// <param name="Ordernummer">Ordernummret kopplat till detta obejkt. </param>
            /// <returns>Returnerar ett sql-kommando som l�gger in alla uppgifter i en databas. </returns>
            public OleDbCommand ByggUppSQL(Startplats Startplats, string Ordernummer)
            {
                // Kommandot som skall byggas upp
                OleDbCommand kommando = new OleDbCommand();

                // Grunden i sql-satsen. 
                string SQLSats = "INSERT INTO Objekt (Ordernr, Startplats, Objektnr, Avdnr, Avdnamn, Areal_ha, ";
                SQLSats += "Giva_KgN_ha, Skog_CAN_ton, Kommentar) VALUES (";

                // L�gger till alla uppgifter ifr�n best�llningen. 
                SQLSats += "@Ordernummer, @StartPlats, @_Objektnummer, @_Avdelningsnummer, @_Avdelningsnamn, ";
                SQLSats += "@_Areal, @_Giva, @_CAN, @_Kommentar)";

                // Anger kommandotexten. 
                kommando.CommandText = SQLSats;

                // Anger vilka typer som �r legala i kommandot. 
                kommando.Parameters.Add("@Ordernummer", OleDbType.Integer);
                kommando.Parameters.Add("@StartPlats", OleDbType.VarChar);
                kommando.Parameters.Add("@_Objektnummer", OleDbType.VarChar);
                kommando.Parameters.Add("@_Avdelningsnummer", OleDbType.VarChar);
                kommando.Parameters.Add("@_Avdelningsnamn", OleDbType.VarChar);
                kommando.Parameters.Add("@_Areal", OleDbType.Double);
                kommando.Parameters.Add("@_Giva", OleDbType.Double);
                kommando.Parameters.Add("@_CAN", OleDbType.Double);
                kommando.Parameters.Add("@_Kommentar", OleDbType.VarChar);

                // L�gger till alla v�rden
                kommando.Parameters[0].Value = Ordernummer;
                kommando.Parameters[1].Value = Startplats.StartPlats;
                kommando.Parameters[2].Value = _Objektnummer;
                kommando.Parameters[3].Value = _Avdelningsnummer;
                kommando.Parameters[4].Value = _Avdelningsnamn;
                kommando.Parameters[5].Value = _Areal;
                kommando.Parameters[6].Value = _Giva;
                kommando.Parameters[7].Value = _CAN;
                kommando.Parameters[8].Value = _Kommentar;

                return kommando;
            }

            #region get- och setegenskaper

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
            /// H�mtar antal ton CAN. 
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
