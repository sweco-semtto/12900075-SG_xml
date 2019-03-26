using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Windows.Forms;

namespace SG_xml
{
    /// <summary>
    /// En delegerad funktion f�r att l�sa ifr�n databasen med. 
    /// </summary>
    /// <param name="SQLFr�ga">SQL-fr�gan som skall st�llas till databasen. </param>
    /// <returns>Returnerar ett dataset med svaret ifr�n databasen. </returns>
    public delegate DataSet L�sIfr�nDatabas(string SQLFr�ga);

    /// <summary>
    /// Har hand om l�sning och skrivning till databasen. 
    /// 
    /// Skapad av MTTO. 
    /// </summary>
    public class Accessdatabas
    {
        #region instansvariabler

        /// <summary>
        /// S�kv�gen till databasen som finns angiven. 
        /// </summary>
        private static string _S�kv�gDatabas = "";

        /// <summary>
        /// Det senaste felmeddelandet som denna klass har genererat. 
        /// </summary>
        private static string _Felmeddelande;
        /// <summary>
        /// Anger om skrivningen till databasen gick d�ligt. 
        /// </summary>
        private static bool _FelISQL = false;

        #endregion

        /// <summary>
        /// L�ser ifr�n databasen med en SQL-fr�ga. 
        /// </summary>
        /// <param name="SQLFr�ga">SQL-fr�gan som skall st�llas till databasen. </param>
        /// <returns>Returnerar ett dataset med svaret fr�n fr�ga. </returns>
        public static DataSet L�sIfr�nDatabas(string SQLFr�ga)
        {
            // Skapar en uppkoppling mot databasen. 
            OleDbConnection uppkoppling = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _S�kv�gDatabas);

            try
            {
                // �ppnar databasuppkopplingen. 
                uppkoppling.Open();

                // L�ser ifr�n databasen. 
                DataSet data = new DataSet();
                OleDbDataAdapter adapter = new OleDbDataAdapter(SQLFr�ga, uppkoppling);
                adapter.Fill(data);

                // St�nger ned uppkopplingen. 
                uppkoppling.Close();

                return data;
            }
            //Some usual exception handling
            catch (OleDbException oledbe)
            {
               
                    //_FelISQL = true;
                    _Felmeddelande = oledbe.Message;

                    MessageBox.Show("Misslyckades l�sa ifr�n databas. (felnummer 1)\n\n" + oledbe.Message, "Misslyckades skriva till databas", MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }
            finally
            {
                // Om uppkopplingen �r aktiv, skall den st�ngas ned efter fel. 
                if (uppkoppling != null)
                {
                    uppkoppling.Close();
                }
            }

            return null;
        }

		/// <summary>
		/// Avg�r om det finns ett ordernummer redan �r best�llt, d.v.s. finns i Access-databasen. 
		/// </summary>
		/// <param name="F�retag">Best�llning fr�n ett f�retg. </param>
		/// <returns></returns>
		public static bool FinnsOrdernummer(F�retag F�retag)
		{
			DataSet datasetOrdernummer = Accessdatabas.L�sIfr�nDatabas(F�retag.FinnsOrder());

			if (datasetOrdernummer.Tables.Count > 0 && datasetOrdernummer.Tables[0].Rows.Count > 0)
				return true;

			return false;
		}

        /// <summary>
        /// Skriver en best�llning till databasen. 
        /// </summary>
        /// <param name="best�llning">Best�llningen som skall sparas i databasen. </param>
        /// <returns></returns>
        public static void SkrivTillDatabas(OleDbCommand kommando)
        {
            // Skapar en uppkoppling mot databasen. 
            OleDbConnection uppkoppling = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _S�kv�gDatabas);

            // S�tter anslutningen till kommandot. 
            kommando.Connection = uppkoppling;

            try
            {
                // �ppnar databasuppkopplingen. 
                uppkoppling.Open();

                // Skriver till databasen. 
                kommando.ExecuteNonQuery();

                // St�nger ned uppkopplingen. 
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
                MessageBox.Show("Ov�ntat fel intr�ffat, v�nligen rapportera detta meddelande till Sweco.\n" + fex, "Ov�ntat fel", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            finally
            {
                // Om uppkopplingen �r aktiv, skall den st�ngas ned efter fel. 
                if (uppkoppling != null)
                {
                    uppkoppling.Close();
                }
            }
        }

        #region get- och setegenskaper

        /// <summary>
        /// H�mtar ett v�rde som talar om om skrivningen till databasen gick d�ligt. 
        /// </summary>
        public static bool FelISQL
        {
            get
            {
                return _FelISQL;
            }
        }

        /// <summary>
        /// H�mtar eller skriver s�kv�gen till databasen. 
        /// </summary>
        public static string S�kv�gDatabas
        {
            get
            {
                return _S�kv�gDatabas;
            }
            set
            {
                _S�kv�gDatabas = value;
            }
        }

        #endregion
    }
}
