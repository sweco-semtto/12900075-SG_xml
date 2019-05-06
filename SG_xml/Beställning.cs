using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Text;
using System.Windows.Forms;

namespace SG_xml
{
    /// <summary>
    /// Skapar en ny beställning som skall sparas ned i en access-databas. 
    /// 
    /// Skapad av MTTO. 
    /// </summary>
    public class Beställning
    {
        #region instansvariabler

        private Företag _Företag;
        private Koordinatsystem _Koordinatsystem;
        private List<Startplats> _Startplatser;
        private List<ReservObjekt> _Reservobjekt;

        #endregion

        /// <summary>
        /// Skapar en ny beställning och skriver den till databasen.  
        /// </summary>
        /// <param name="xml">XML-strängen som skall läsas in. </param>
        public Beställning(string xml)
        {
            // Om en tom stäng är angiven skall detta meddelas. 
            if (xml == "")
            {
                MessageBox.Show("Den angvinga xml-stängen, texten i den stora rutan, är tom. Vänligen klista in text ifrån en epostbeställning. ", "Beställningstext saknas", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Bygger upp företagsobjektet och kontrollera om allt gick bra.
            _Företag = Företag.ByggUppObjekt(xml);
            if (Företag.FelIXML)
                return;

            _Koordinatsystem = Koordinatsystem.ByggUppObjekt(xml);
            if (Koordinatsystem.FelIXML)
                return;

            // Bygger upp startplatsobjekten och kontrollera om allt gick bra.
            _Startplatser = Startplats.ByggUppObjekt(xml);
            if (Startplats.FelIXML)
            {
                MessageBox.Show("Problem med att läsa in startplats och värdet " + Startplats.FelFörVärde + ". \n\n" + Startplats.Felmeddelande, "Problem med Startplats", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Transformerar om startplatser vid vid behov, d.v.s. om koordinaterna inte är i SWEREF99 TM. 
            _Koordinatsystem.TransformeraStartplatser(_Startplatser);

            // Bygger upp reservobjekten och kontrollera om allt gick bra.
            _Reservobjekt = ReservObjekt.ByggUppObjekt(xml);
            if (ReservObjekt.FelIXML)
                return;

			// Kontrollerar om order redan finns med sedan innan i Access
			if (Accessdatabas.FinnsOrdernummer(_Företag))
			{
				MessageBox.Show("Beställningen finns redan i Access-databasen. ", "Beställning redan inlagd", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

            //Skriver xml-en till MySql som en backup.
            if (!MySqlCommunicator.OnlyTestData)
            {
                bool success = MySqlCommunicator.BackupOrderToMySql(xml, _Företag.Ordernummer);
                if (!success)
                {
                    MessageBox.Show("Kan inte skapa en backup på ordern i MySql. ", "Problem med MySql", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

			// Skriver företagsuppgifterna (uppgifter från xml:en och en tidsstämpel) till databasen. 
			Accessdatabas.SkrivTillDatabas(_Företag.ByggUppSQL());
            if (Accessdatabas.FelISQL)
                return;

            // Hämtar ordernummret för den just tillagda företagsuppgiften i databasen m.h.a. tidsstämpeln. 
            string ordernummer = _Företag.HämtaOrdernummer(Accessdatabas.LäsIfrånDatabas);
            if (Accessdatabas.FelISQL)
                return;

            // Skriver in alla startplatser till databasen.   
            foreach (Startplats startsplats in _Startplatser)
            {
                // Skriver in startplatsen i databasen. 
                Accessdatabas.SkrivTillDatabas(startsplats.ByggUppSQL(ordernummer));
                if (Accessdatabas.FelISQL)
                    return;

                // Skriver in alla objekt kopplade till startplatsen i databasen. 
                foreach (ObjektIStartplats objekt in startsplats.Objekt)
                {
                    Accessdatabas.SkrivTillDatabas(objekt.ByggUppSQL(startsplats, ordernummer));
                    if (Accessdatabas.FelISQL)
                        return;
                }
            }

            // Skriver in alla reservobjekt i databasen. 
            foreach (ReservObjekt reservobjekt in _Reservobjekt)
            {
                Accessdatabas.SkrivTillDatabas(reservobjekt.ByggUppSQL(ordernummer));
                if (Accessdatabas.FelISQL)
                    return;
            }
  
            // Skriver in information i kolumnen "Ingående Objekt" för alla tillagda startplatser.
            foreach (Startplats startplats in _Startplatser)
            {
                
            }

            if (!MySqlCommunicator.OnlyTestData)
                MessageBox.Show("Beställningen är inlagd i Access-databasen och ordernumret i MySql. ", "Inlagd beställning", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show("Beställningen är inlagd i Access-databasen, inget ordernummer skapat i MySql. ", "Inlagd beställning", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
