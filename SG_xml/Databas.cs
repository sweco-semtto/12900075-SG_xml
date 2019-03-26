using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Windows.Forms;

namespace SG_xml
{
    /// <summary>
    /// En delegerad funktion för att läsa ifrån databasen med. 
    /// </summary>
    /// <param name="SQLFråga">SQL-frågan som skall ställas till databasen. </param>
    /// <returns>Returnerar ett dataset med svaret ifrån databasen. </returns>
    public delegate DataSet LäsIfrånDatabas(string SQLFråga);

    /// <summary>
    /// Har hand om läsning och skrivning till databasen. 
    /// 
    /// Skapad av MTTO. 
    /// </summary>
    public class Accessdatabas
    {
        #region instansvariabler

        /// <summary>
        /// Sökvägen till databasen som finns angiven. 
        /// </summary>
        private static string _SökvägDatabas = "";

        /// <summary>
        /// Det senaste felmeddelandet som denna klass har genererat. 
        /// </summary>
        private static string _Felmeddelande;
        /// <summary>
        /// Anger om skrivningen till databasen gick dåligt. 
        /// </summary>
        private static bool _FelISQL = false;

        #endregion

        /// <summary>
        /// Läser ifrån databasen med en SQL-fråga. 
        /// </summary>
        /// <param name="SQLFråga">SQL-frågan som skall ställas till databasen. </param>
        /// <returns>Returnerar ett dataset med svaret från fråga. </returns>
        public static DataSet LäsIfrånDatabas(string SQLFråga)
        {
            // Skapar en uppkoppling mot databasen. 
            OleDbConnection uppkoppling = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _SökvägDatabas);

            try
            {
                // Öppnar databasuppkopplingen. 
                uppkoppling.Open();

                // Läser ifrån databasen. 
                DataSet data = new DataSet();
                OleDbDataAdapter adapter = new OleDbDataAdapter(SQLFråga, uppkoppling);
                adapter.Fill(data);

                // Stänger ned uppkopplingen. 
                uppkoppling.Close();

                return data;
            }
            //Some usual exception handling
            catch (OleDbException oledbe)
            {
               
                    //_FelISQL = true;
                    _Felmeddelande = oledbe.Message;

                    MessageBox.Show("Misslyckades läsa ifrån databas. (felnummer 1)\n\n" + oledbe.Message, "Misslyckades skriva till databas", MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }
            finally
            {
                // Om uppkopplingen är aktiv, skall den stängas ned efter fel. 
                if (uppkoppling != null)
                {
                    uppkoppling.Close();
                }
            }

            return null;
        }

		/// <summary>
		/// Avgör om det finns ett ordernummer redan är beställt, d.v.s. finns i Access-databasen. 
		/// </summary>
		/// <param name="Företag">Beställning från ett företg. </param>
		/// <returns></returns>
		public static bool FinnsOrdernummer(Företag Företag)
		{
			DataSet datasetOrdernummer = Accessdatabas.LäsIfrånDatabas(Företag.FinnsOrder());

			if (datasetOrdernummer.Tables.Count > 0 && datasetOrdernummer.Tables[0].Rows.Count > 0)
				return true;

			return false;
		}

        /// <summary>
        /// Skriver en beställning till databasen. 
        /// </summary>
        /// <param name="beställning">Beställningen som skall sparas i databasen. </param>
        /// <returns></returns>
        public static void SkrivTillDatabas(OleDbCommand kommando)
        {
            // Skapar en uppkoppling mot databasen. 
            OleDbConnection uppkoppling = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _SökvägDatabas);

            // Sätter anslutningen till kommandot. 
            kommando.Connection = uppkoppling;

            try
            {
                // Öppnar databasuppkopplingen. 
                uppkoppling.Open();

                // Skriver till databasen. 
                kommando.ExecuteNonQuery();

                // Stänger ned uppkopplingen. 
                uppkoppling.Close();
            }
            //Some usual exception handling
            catch (OleDbException oledbex)
            {
                _FelISQL = true;
                _Felmeddelande = oledbex.Message;

                MessageBox.Show("Misslyckades skriva till databas. (felnummer 2)\n\n" + oledbex.Message, "Misslyckades skriva till databas", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (InvalidOperationException ioex)
            {
                _FelISQL = true;
                _Felmeddelande = ioex.Message;

                MessageBox.Show("Misslyckades skriva till databas. (felnummer 3)\n\n" + ioex.Message, "Misslyckades skriva till databas", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (FormatException fex)
            {
                MessageBox.Show("Oväntat fel inträffat, vänligen rapportera detta meddelande till Sweco.\n" + fex, "Oväntat fel", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            finally
            {
                // Om uppkopplingen är aktiv, skall den stängas ned efter fel. 
                if (uppkoppling != null)
                {
                    uppkoppling.Close();
                }
            }
        }

        #region get- och setegenskaper

        /// <summary>
        /// Hämtar ett värde som talar om om skrivningen till databasen gick dåligt. 
        /// </summary>
        public static bool FelISQL
        {
            get
            {
                return _FelISQL;
            }
        }

        /// <summary>
        /// Hämtar eller skriver sökvägen till databasen. 
        /// </summary>
        public static string SökvägDatabas
        {
            get
            {
                return _SökvägDatabas;
            }
            set
            {
                _SökvägDatabas = value;
            }
        }

        #endregion
    }
}
