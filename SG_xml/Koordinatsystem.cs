using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using ProjNet.CoordinateSystems;
using ProjNet.CoordinateSystems.Transformations;
using ProjNet.Converters.WellKnownText;

namespace SG_xml
{
    public class Koordinatsystem
    {
        #region instansvariabler

        private MöjligaKoordinatsystem _Koordinatsystem;

        private ICoordinateTransformation _Transformera;

        private static string _Felmeddelande;
        private static bool _FelIXML = false;

        #endregion

        /// <summary>
        /// Skapar ett nytt Koordinatsystemsobjekt. 
        /// </summary>
        /// <param name="xml">Den xml-sträng som bygger upp objektet. </param>
        protected Koordinatsystem(string xml)
		{
			LäsIn(xml);

            this.ProjektionsträngRT90_25gonV = "PROJCS[\"RT90_25_gon_W-approximativ\",GEOGCS[\"GCS_RT_1990\",DATUM[\"D_GRS_1980\",SPHEROID[\"GRS_1980\",6378137.0,298.257222101]],PRIMEM[\"Greenwich\",0.0],UNIT[\"Degree\",0.0174532925199433]],PROJECTION[\"Transverse_Mercator\"],PARAMETER[\"False_Easting\",1500064.274],PARAMETER[\"False_Northing\",-667.711],PARAMETER[\"Central_Meridian\",15.80628452944445],PARAMETER[\"Scale_Factor\",1.00000561024],PARAMETER[\"Latitude_Of_Origin\",0.0],UNIT[\"Meter\",1.0]]";
            this.ProjektionsträngSWEREF99_TM = "PROJCS[\"SWEREF99 TM\",GEOGCS[\"SWEREF99\",DATUM[\"D_SWEREF99\",SPHEROID[\"GRS_1980\",6378137,298.257222101]],PRIMEM[\"Greenwich\",0],UNIT[\"Degree\",0.017453292519943295]],PROJECTION[\"Transverse_Mercator\"],PARAMETER[\"latitude_of_origin\",0],PARAMETER[\"central_meridian\",15],PARAMETER[\"scale_factor\",0.9996],PARAMETER[\"false_easting\",500000],PARAMETER[\"false_northing\",0],UNIT[\"Meter\",1]]";
		}

        public static Koordinatsystem ByggUppObjekt(string xml)
        {
            Koordinatsystem nyttObjekt = new Koordinatsystem(xml);

            return nyttObjekt;
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

                // Läser in beställningsdatum. 
                xmlNode = xmlDoc.SelectSingleNode("Beställning/child::Koordinatsystem");
                string koordinatsystem = xmlNode != null ? xmlNode.InnerText : string.Empty;
                if (koordinatsystem.Equals("RT90"))
                    this._Koordinatsystem = MöjligaKoordinatsystem.RT90_25gonV;
                else if (koordinatsystem.Equals("SWEREF99"))
                    this._Koordinatsystem = MöjligaKoordinatsystem.SWEREF99_TM;
            }
            catch (XmlException xmlex)
            {
                _FelIXML = true;
                _Felmeddelande = xmlex.Message;

                // Meddelare användaren om detta fel. 
                MessageBox.Show("Xml-strängen innehåller fel inom Koordinatsystemstaggen och kan ej användas för att spara data med. \nLeta efter felet på rad " + xmlex.LineNumber + " och teckennummer " + xmlex.LinePosition+ ". ", "Felaktig xml", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (Exception ex)
            {
                _Felmeddelande = ex.Message;
            }
        }

        /// <summary>
        /// Transformerar koordinater om nuvarande koordinatsystem är RT90. 
        /// </summary>
        /// <param name="x">X-koordinaten i RT 90 2,5 gon V</param>
        /// <param name="y">Y-koordinaten i RT 90 2,5 gon V</param>
        public void TransformeraRT90KoordinaterTillSWEREF99(ref double x, ref double y)
        {
            // Anger från- och tillsystem
            IProjectedCoordinateSystem frånsystem =
                CoordinateSystemWktReader.Parse(this.ProjektionsträngRT90_25gonV) as IProjectedCoordinateSystem;
            IProjectedCoordinateSystem tillsystem =
                CoordinateSystemWktReader.Parse(this.ProjektionsträngSWEREF99_TM) as IProjectedCoordinateSystem;

            // Skapar en fabrik och anger från- och tillsystem. 
            CoordinateTransformationFactory transformationsfabriken = new CoordinateTransformationFactory();
            _Transformera = transformationsfabriken.CreateFromCoordinateSystems(frånsystem, tillsystem);

            // Transformerar korodinaterna
            double[] RT90_koordinater = new double[] { x, y };
            double[] SWEREF99_koordinater = _Transformera.MathTransform.Transform(RT90_koordinater);

            // Skriver tillbaka de nya värdena. 
            x = SWEREF99_koordinater[0];
            y = SWEREF99_koordinater[1];
        }

        /// <summary>
        /// Transformera startplatsernas koordinater vid behov, d.v.s. endast om koordinaterna är angivna i RT 90 2,5 gon V.  
        /// </summary>
        /// <param name="startplatser">Startplatserna med transformerade koordinater i,. </param>
        public void TransformeraStartplatser(List<Startplats> startplatser)
        {
            // Transformerar endast om vi inte har SWEREF 99 TM
            if (this.ValtKoordinatsystem == MöjligaKoordinatsystem.SWEREF99_TM)
                return;

            // Loopar igenom alla startplatser. 
            for (int startplatsId = 0; startplatsId < startplatser.Count; startplatsId++)
            {
                // Tar fram de gamla koordinaterna. 
                double x = startplatser[startplatsId].Ostligkoordinat;
                double y = startplatser[startplatsId].Nordligkoordinat;

                // TGransformerar koordinaterna. 
                TransformeraRT90KoordinaterTillSWEREF99(ref x, ref y);

                // Skriver tillbaka de nya koordinatvärderna. 
                startplatser[startplatsId].Ostligkoordinat = Math.Round(x, 0, MidpointRounding.AwayFromZero);
                startplatser[startplatsId].Nordligkoordinat = Math.Round(y, 0, MidpointRounding.AwayFromZero);
            }
        }

        #region get- och setegenskaper

        /// <summary>
        /// Hämtar vilket koordinatsystem som beställningen är gjord i. 
        /// </summary>
        public MöjligaKoordinatsystem ValtKoordinatsystem
        {
            get
            {
                return _Koordinatsystem;
            }
        }

        /// <summary>
        /// Hämtar eller anger projektionssträngen för RT 90 2,5 gon V
        /// </summary>
        public string ProjektionsträngRT90_25gonV
        {
            get;
            set;
        }

        /// <summary>
        /// Hämtar eller anger projektionssträngen för SWEREF99 TM
        /// </summary>
        public string ProjektionsträngSWEREF99_TM
        {
            get;
            set;
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

        #endregion
    }

    /// <summary>
    /// Anger vilka koordinatsystem som det finns stöd för. 
    /// </summary>
    public enum MöjligaKoordinatsystem
    {
        [System.ComponentModel.Description("RT 90 2,5 gonV")]
        RT90_25gonV = 0,

        [System.ComponentModel.Description("SWEREF99 TM")]
        SWEREF99_TM = 1
    }
}
