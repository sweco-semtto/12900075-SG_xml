using System;
using System.Collections.Generic;
using System.Net;
using System.IO;
using System.Text;

namespace SG_xml
{
    /// <summary>
    /// Statisk klass som är till för att sköta kommunikationen med MySql. 
    /// </summary>
    public class MySqlCommunicator
    {
        /// <summary>
        /// Sökvägen till var skriptet för att få fram dagens datum ligger. 
        /// </summary>
        protected static string _Url_To_Get_Current_Time = "http://www.sg-systemet.com/bestallning/Gettime.php";

        /// <summary>
        /// Sökvägen till var skriptet för få ett nytt ordernummer ligger. 
        /// </summary>
        protected static string _Url_To_New_Ordernumber = "http://www.sg-systemet.com/bestallning/CreateOrdernumber.php";

        /// <summary>
        /// Sökvägen till var skriptet för få ett nytt ordernummer ligger. 
        /// </summary>
        protected static string _Url_To_Backup_Order = "http://www.sg-systemet.com/bestallning/BackupOrder.php";

        /// <summary>
        /// Konstant som anger att inget år har mottagits. 
        /// </summary>
        protected const string NO_YEAR_RECEIVED = "0";

        /// <summary>
        /// Konstant som anger att inget ordernummer har hämtats. 
        /// </summary>
        protected const string NO_ORDERNUMBER = "Inget ordernummer";

        /// <summary>
        /// Anger vilket ordernummer som senast blev satt. 
        /// </summary>
        private static string _Current_Ordernumber = NO_ORDERNUMBER;
        
        /// <summary>
        /// Hämtar ett nytt ordernummer från MySql. 
        /// </summary>
        /// <returns>Returerar ett nytt ordernummer från MySql. </returns>
        public static string GetNewOrdernumber()
        {
            string ans = NO_ORDERNUMBER;

            try
            {
                // Skickar en request och tar emot ett response.
                string timestamp = GetAnswerStringFromPHPResponse(SendRequestToPHP(_Url_To_Get_Current_Time));
                string currentYear = GetCurrentYearFromPHP(timestamp);
                ans = GetAnswerStringFromPHPResponse(SendRequestToPHP(_Url_To_New_Ordernumber, "Year=" + currentYear + "&Timestamp=" + timestamp));

                // Sparar undan vilket ordernummer tills backupen. 
                _Current_Ordernumber = ans;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ans;
        }

        /// <summary>
        /// Skriver order till MySql som en backup. 
        /// </summary>
        /// <param name="orderInXML">Ordern som skall backupas. </param>
        /// <returns>Returnerar true om allt gick bra. </returns>
        public static bool BackupOrderToMySql(string orderInXML)
        {
            bool ans = true;

            // Ersätter å, ä, ö m.fl. till xml-säkra tecken och byter ut & och = så att vi kan skicka argument till PHP. 
            orderInXML = ChangeUTF_8Characters(orderInXML);
            orderInXML = ChangeAmpersandAndEqualsCharacters(orderInXML);

            try
            {
                _Current_Ordernumber = "100";

                // Skickar en request och tar emot ett response.
                string timestamp = GetAnswerStringFromPHPResponse(SendRequestToPHP(_Url_To_Get_Current_Time));
                ans = GetAnswerStringFromPHPResponse(SendRequestToPHP(_Url_To_Backup_Order, "time=" + timestamp + "&ordernummer=" + _Current_Ordernumber + "&xml=" + orderInXML)).Equals("true") ? true : false;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ans;
        }

        /// <summary>
        /// Skickar väg data till php utan någon argument. 
        /// </summary>
        /// <param name="urlToPHP">Sökvägen till php-skriptet. </param>
        private static HttpWebResponse SendRequestToPHP(string urlToPHP)
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlToPHP);
                return (HttpWebResponse)request.GetResponse();
            }
            catch (System.Net.WebException ex)
            {
                throw ex;
            }
            catch (System.Net.ProtocolViolationException ex)
            {
                throw ex;
            }
            catch (ObjectDisposedException ex)
            {
                throw ex;
            }
            catch (InvalidOperationException ex)
            {
                throw ex;
            }
            catch (ArgumentNullException ex)
            {
                throw ex;
            }
            catch (NotSupportedException ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Skickar iväg data till php med argument. 
        /// </summary>
        /// <param name="urlToPHP">Sökvägen till php-skriptet. </param>
        /// <param name="postData">Argumentet som skall skickas</param>
        private static HttpWebResponse SendRequestToPHP(string urlToPHP, string postData)
        {
            try
            {
                //ASCIIEncoding enc = new ASCIIEncoding();
                UTF8Encoding encoding = new UTF8Encoding();
                //byte[] POST = encoding.GetBytes(postData);
                byte[] POST = Encoding.UTF8.GetBytes(postData);
                
                // Bygger upp en request. 
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlToPHP);
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded; encoding='utf-8'";
                //request.ContentType = "text/xml; encoding='utf-8'";
                //request.ContentType = "text/html; Charset=UTF-8";
                request.ContentLength = POST.Length;

                // Data till PHP som en ström
                Stream StreamPOST = request.GetRequestStream();
                StreamPOST.Write(POST, 0, POST.Length);
                StreamPOST.Close();

                // Skickar en request och tar emot ett response. 
                return (HttpWebResponse)request.GetResponse();
            }
            catch (System.Net.WebException ex)
            {
                throw ex;
            }
            catch (System.Net.ProtocolViolationException ex)
            {
                throw ex;
            }
            catch (ObjectDisposedException ex)
            {
                throw ex;
            }
            catch (InvalidOperationException ex)
            {
                throw ex;
            }
            catch (ArgumentNullException ex)
            {
                throw ex;
            }
            catch (NotSupportedException ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// Tar fram svarssträngen från php-svaret. 
        /// </summary>
        /// <param name="response">Svaret från php. </param>
        /// <returns>Returnerar bara själva svarstexten som php skickar. </returns>
        private static string GetAnswerStringFromPHPResponse(HttpWebResponse response)
        {
            string ans = String.Empty;
            try
            {
                StreamReader responseStream = new StreamReader(response.GetResponseStream());
                ans = responseStream.ReadToEnd();
            }
            catch (Exception ex)
            {
            }

            return ans;
        }

        /// <summary>
        /// Hämtar vilket år som vi har nu. 
        /// </summary>
        /// <returns>Returerar året från en tidsstämpel. </returns>
        private static string GetCurrentYearFromPHP(string timestamp)
        {
            // Tar bara fram året från dagens datum, t.ex. 2017-03-08 blir 2017. 
            return timestamp.Substring(0, 4);
        }

        /// <summary>
        /// Ersätter alla UTF-8-tecken som har kommit upp i olika beställningar genom åren till xml-säkra bokstäver. 
        /// </summary>
        /// <param name="text">Texten som skall sökas igenom. </param>
        /// <returns>En xml-säker text innehållandes tecken från UTF-8. </returns>
        private static string ChangeUTF_8Characters(string text)
        {
            string ans = String.Empty;
			
			if (text == null)
				return null;
		
			// Ersätter varje å, ä och ö med html-motsvarigheten. 
            char[] allCharactes = text.ToCharArray();
            for (var i = 0; i < allCharactes.Length; i++)
			{
                char letter = allCharactes[i];
				
				if (letter == 'å')
				{
					ans += "&aring";
				}
				else if (letter == 'Å')
				{
					ans += "&Aring";
				}
				else if (letter == 'ä')
				{
					ans += "&auml";
				}
				else if (letter == 'Ä')
				{
					ans += "&Auml";
				}
				else if (letter == 'ö')
				{
					ans += "&ouml";
				}
				else if (letter == 'Ö')
				{
					ans += "&Ouml";
				}
				else if (letter == 'ü')
				{
					ans += "&uuml";
				}
				else if (letter == 'Ü')
				{
					ans += "&Uuml";
				}
				else if (letter == 'û')
				{
					ans += "&ucirc";
				}
				else if (letter == 'Û')
				{
					ans += "&Ucirc";
				}
				else if (letter == 'é')
				{
					ans += "&egrave";
				}
				else if (letter == 'É')
				{
					ans += "&Egrave";
				}
				else if (letter == 'è')
				{
					ans += "&eacute";
				}
				else if (letter == 'È')
				{
					ans += "&Eacute";
				}
				else if (letter == '&')
				{
					ans += "&amp";
				}
				else if (letter == '<')
				{
					ans += "&lt";
				}
				else if (letter == '>')
				{
					ans += "&gt";
				}
				else if (letter == '"')
				{
					ans += "&quot";
				}
				else if (letter.ToString() == "'")
				{
					ans += "&#39";
				}
				else
				{
					ans += letter;
				}
			}
			
			return ans;
        }

        /// <summary>
        /// Byter ut alla &- och =-tecken så att argumenten kommer vidare till PHP. 
        /// </summary>
        /// <param name="text">Texten som skall modiferias. </param>
        /// <returns>Den modifierade texten. </returns>
        private static string ChangeAmpersandAndEqualsCharacters(string text)
        {
            string ans = String.Empty;

            if (text == null)
				return null;
		
			// Ersätter varje å, ä och ö med html-motsvarigheten. 
            char[] allCharactes = text.ToCharArray();
            for (var i = 0; i < allCharactes.Length; i++)
			{
                char letter = allCharactes[i];

                if (letter == '&')
                    ans += "|ampersand|";
                else if (letter == '=')
                    ans += "|equals|";
                else
                    ans += letter;
            }

            return ans;
        }
    }
}
