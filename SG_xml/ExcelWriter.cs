/****************************************************
 * ExcelWriter.cs Writes Excel files to hard drive with the help of FileIO.cs
 * 
 * 
 * 
 * Fredrik Björklund FEBJ, SWECO Position, 2008
 * Modifierad av Martin Thorbjörnsson MTTO, SWECO Postion, 2008
 * 
 *****************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using CarlosAg.ExcelXmlWriter;
using System.Data;
using System.Xml;



namespace SG_xml
{
    
    class ExcelWriter
    {
        private DataRow drowOfCData;

        /// <summary>
        /// WriteExcelFileFromDataset: Write Excel file from data in database.
        /// </summary>
        /// <param name="myCompanyData"></param>
        /// <param name="path"></param>
        public bool WriteExcelFileFromDataset(DataSet myCompanyData, string path)
        {
            //returns true if repport has been created
            bool repportGenerated = false;

            try
            {
                // Get the table from the data set
                DataTable dCompanyData = myCompanyData.Tables[0];

                //Create FileIO objekt to check file
                FileIO myXMLFile = new FileIO();
                FileIO myXLSFile = new FileIO();


                //Write all order nr
                for (int i = 0; i < dCompanyData.Rows.Count; i++)
                {
                    DataRow drowOfCompayData = dCompanyData.Rows[i];

                    // Only row that have not been deleted
                    if (drowOfCompayData.RowState != DataRowState.Deleted)
                    {
                        string nameOfExcelFile = drowOfCompayData["Företagsnamn"].ToString() + "_" + drowOfCompayData["Region_Förvaltning"].ToString() + "_" + drowOfCompayData["Distrikt_Område"].ToString() + "_" + drowOfCompayData["Ordernr"].ToString();

                        //Fix date, remove time stamp
                  //      string dateForFileName = drowOfCompayData["Beställningsdatum"].ToString();
                  //      dateForFileName = dateForFileName.Remove(dateForFileName.IndexOf(' '));

                        //Create file name for excel file
                  //      string nameOfExcelFile = drowOfCompayData["Företagsnamn"].ToString() + "_" + dateForFileName + "_" + drowOfCompayData["Ordernr"].ToString();
                        drowOfCData = drowOfCompayData;

                        string tempFileName = DateTime.Now.Ticks.ToString();

                        string fullPath = path + "\\" + nameOfExcelFile + ".xls";
                        string fullPathTemp = path + "\\" + tempFileName + ".xls";

                        //Create the excel file
                        myXMLFile.CheckXMLFile(fullPathTemp);

                        //Generate(fullPathTemp);
                        Generate(fullPath);
                        repportGenerated = true;
                   
             /*           if (myXLSFile.CheckXLSFile(fullPathTemp, fullPath, nameOfExcelFile + ".xls"))
                        {
                            XLSWriter myXLSConverter = new XLSWriter();
                            myXLSConverter.WriteXLSFile(fullPath, fullPathTemp);
                            repportGenerated = true;

                        }
                       */

                        if (myXLSFile.ExitXLSExport)
                        {
                            
                            return repportGenerated;
                        }

                     

                        if (myXMLFile.ExitExport)
                            return repportGenerated;
                    }

                    
                }

                return repportGenerated;

            }
            catch (SystemException ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// Generate: Generate Excel file
        /// </summary>
        /// <param name="filename"></param>
        public void Generate(string filename)
        {
            Workbook book = new Workbook();
            // -----------------------------------------------
            //  Properties
            // -----------------------------------------------
            book.Properties.Author = "Lina Samor";
            book.Properties.LastAuthor = "Fredrik Björklund";
            book.Properties.Created = new System.DateTime(2008, 1, 14, 15, 2, 41, 0);
            book.Properties.LastSaved = new System.DateTime(2008, 1, 18, 16, 0, 41, 0);
            book.Properties.Company = "SWECO";
            book.Properties.Version = "11.8132";
            book.ExcelWorkbook.WindowHeight = 13035;
            //book.ExcelWorkbook.WindowWidth = 15195;
            book.ExcelWorkbook.WindowWidth = 20000;
            book.ExcelWorkbook.WindowTopX = 480;
            book.ExcelWorkbook.WindowTopY = 30;
            book.ExcelWorkbook.ProtectWindows = false;
            book.ExcelWorkbook.ProtectStructure = false;
            // -----------------------------------------------
            //  Generate Styles
            // -----------------------------------------------
            this.GenerateStyles(book.Styles);
            // -----------------------------------------------
            //  Generate Blad1 Worksheet
            // -----------------------------------------------
            this.GenerateWorksheetBlad1(book.Worksheets);

            book.Save(filename);
        }

        /// <summary>
        /// GenerateStyles: Generate the styles for excel dokument
        /// </summary>
        /// <param name="styles"></param>
        private void GenerateStyles(WorksheetStyleCollection styles)
        {
            // -----------------------------------------------
            //  Default
            // -----------------------------------------------
            WorksheetStyle Default = styles.Add("Default");
            Default.Name = "Normal";
            Default.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  m19482424
            // -----------------------------------------------
            WorksheetStyle m19482424 = styles.Add("m19482424");
            m19482424.Font.Color = "#0000FF";
            m19482424.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482424.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482424.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482424.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482424.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  m19482434
            // -----------------------------------------------
            WorksheetStyle m19482434 = styles.Add("m19482434");
            m19482434.Font.Color = "#0000FF";
            m19482434.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482434.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482434.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482434.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482434.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  m19482272
            // -----------------------------------------------
            WorksheetStyle m19482272 = styles.Add("m19482272");
            m19482272.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            m19482272.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482272.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482272.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482272.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482272.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  m19482282
            // -----------------------------------------------
            WorksheetStyle m19482282 = styles.Add("m19482282");
            m19482282.Font.Color = "#333399";
            m19482282.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482282.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482282.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482282.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482282.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            m19482282.NumberFormat = "0";
            // -----------------------------------------------
            //  m19482292
            // -----------------------------------------------
            WorksheetStyle m19482292 = styles.Add("m19482292");
            m19482292.Font.Color = "#333399";
            m19482292.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482292.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482292.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482292.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482292.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            m19482292.NumberFormat = "0";
            // -----------------------------------------------
            //  m19482302
            // -----------------------------------------------
            WorksheetStyle m19482302 = styles.Add("m19482302");
            m19482302.Font.Color = "#0000FF";
            m19482302.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482302.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482302.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482302.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482302.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  m19482108
            // -----------------------------------------------
            WorksheetStyle m19482108 = styles.Add("m19482108");
            m19482108.Font.Color = "#0000FF";
            m19482108.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            m19482108.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482108.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482108.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482108.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482108.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  m19482118
            // -----------------------------------------------
            WorksheetStyle m19482118 = styles.Add("m19482118");
            m19482118.Font.Color = "#0000FF";
            m19482118.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482118.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482118.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482118.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482118.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            m19482118.NumberFormat = "0";
            // -----------------------------------------------
            //  m19482128
            // -----------------------------------------------
            WorksheetStyle m19482128 = styles.Add("m19482128");
            m19482128.Font.Color = "#0000FF";
            m19482128.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482128.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482128.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482128.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482128.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            m19482128.NumberFormat = "0";
            // -----------------------------------------------
            //  m19482138
            // -----------------------------------------------
            WorksheetStyle m19482138 = styles.Add("m19482138");
            m19482138.Font.Color = "#0000FF";
            m19482138.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            m19482138.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482138.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482138.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482138.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482138.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  m19482148
            // -----------------------------------------------
            WorksheetStyle m19482148 = styles.Add("m19482148");
            m19482148.Font.Color = "#0000FF";
            m19482148.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482148.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482148.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482148.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482148.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            m19482148.NumberFormat = "0";
            // -----------------------------------------------
            //  m19482158
            // -----------------------------------------------
            WorksheetStyle m19482158 = styles.Add("m19482158");
            m19482158.Font.Color = "#0000FF";
            m19482158.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19482158.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19482158.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19482158.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19482158.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            m19482158.NumberFormat = "0";
            // -----------------------------------------------
            //  m19481956
            // -----------------------------------------------
            WorksheetStyle m19481956 = styles.Add("m19481956");
            m19481956.Font.Size = 11;
            m19481956.Font.Color = "#0000FF";
            m19481956.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19481956.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19481956.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19481956.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            m19481956.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  m19466902
            // -----------------------------------------------
            WorksheetStyle m19466902 = styles.Add("m19466902");
            m19466902.Font.Color = "#0000FF";
            m19466902.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19466902.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19466902.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19466902.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  m19466912
            // -----------------------------------------------
            WorksheetStyle m19466912 = styles.Add("m19466912");
            m19466912.Font.Color = "#0000FF";
            m19466912.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19466912.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19466912.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19466912.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  m19466922
            // -----------------------------------------------
            WorksheetStyle m19466922 = styles.Add("m19466922");
            m19466922.Font.Color = "#0000FF";
            m19466922.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19466922.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            m19466922.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            m19466922.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s21
            // -----------------------------------------------
            WorksheetStyle s21 = styles.Add("s21");
            s21.Font.Size = 12;
            // -----------------------------------------------
            //  s22
            // -----------------------------------------------
            WorksheetStyle s22 = styles.Add("s22");
            s22.Font.Bold = true;
            s22.Font.Size = 16;
            // -----------------------------------------------
            //  s23
            // -----------------------------------------------
            WorksheetStyle s23 = styles.Add("s23");
            s23.Font.Bold = true;
            s23.Font.Size = 12;
            // -----------------------------------------------
            //  s24
            // -----------------------------------------------
            WorksheetStyle s24 = styles.Add("s24");
            s24.Font.Bold = true;
            s24.Font.Size = 14;
            s24.Font.Color = "#FF0000";
            // -----------------------------------------------
            //  s25
            // -----------------------------------------------
            WorksheetStyle s25 = styles.Add("s25");
            s25.Font.Bold = true;
            s25.Font.Size = 12;
            // -----------------------------------------------
            //  s26
            // -----------------------------------------------
            WorksheetStyle s26 = styles.Add("s26");
            s26.Font.Size = 12;
            s26.Font.Color = "#0000FF";
            s26.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s26.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s27
            // -----------------------------------------------
            WorksheetStyle s27 = styles.Add("s27");
            s27.Font.Bold = true;
            s27.Font.Size = 12;
            s27.Font.Color = "#0000FF";
            s27.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s27.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s28
            // -----------------------------------------------
            WorksheetStyle s28 = styles.Add("s28");
            s28.Font.Bold = true;
            s28.Font.Size = 12;
            s28.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s28.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s28.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s28.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s29
            // -----------------------------------------------
            WorksheetStyle s29 = styles.Add("s29");
            s29.Font.Color = "#FF0000";
            s29.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s30
            // -----------------------------------------------
            WorksheetStyle s30 = styles.Add("s30");
            s30.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s31
            // -----------------------------------------------
            WorksheetStyle s31 = styles.Add("s31");
            s31.Font.Size = 12;
            s31.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s32
            // -----------------------------------------------
            WorksheetStyle s32 = styles.Add("s32");
            s32.Font.Size = 12;
            s32.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s32.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s33
            // -----------------------------------------------
            WorksheetStyle s33 = styles.Add("s33");
            s33.Font.Color = "#0000FF";
            s33.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s33.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s34
            // -----------------------------------------------
            WorksheetStyle s34 = styles.Add("s34");
            s34.Font.Color = "#0000FF";
            s34.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s35
            // -----------------------------------------------
            WorksheetStyle s35 = styles.Add("s35");
            s35.Font.Color = "#0000FF";
            s35.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s36
            // -----------------------------------------------
            WorksheetStyle s36 = styles.Add("s36");
            s36.Font.Color = "#0000FF";
            // -----------------------------------------------
            //  s37
            // -----------------------------------------------
            WorksheetStyle s37 = styles.Add("s37");
            s37.Font.Size = 12;
            s37.Font.Color = "#0000FF";
            // -----------------------------------------------
            //  s38
            // -----------------------------------------------
            WorksheetStyle s38 = styles.Add("s38");
            s38.Font.Size = 12;
            s38.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s39
            // -----------------------------------------------
            WorksheetStyle s39 = styles.Add("s39");
            s39.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s40
            // -----------------------------------------------
            WorksheetStyle s40 = styles.Add("s40");
            // -----------------------------------------------
            //  s41
            // -----------------------------------------------
            WorksheetStyle s41 = styles.Add("s41");
            s41.Font.Size = 8;
            // -----------------------------------------------
            //  s42
            // -----------------------------------------------
            WorksheetStyle s42 = styles.Add("s42");
            s42.Font.Size = 8;
            s42.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s43
            // -----------------------------------------------
            WorksheetStyle s43 = styles.Add("s43");
            s43.Font.Size = 12;
            s43.Font.Color = "#0000FF";
            // -----------------------------------------------
            //  s44
            // -----------------------------------------------
            WorksheetStyle s44 = styles.Add("s44");
            s44.Font.Size = 12;
            s44.Font.Color = "#0000FF";
            s44.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s45
            // -----------------------------------------------
            WorksheetStyle s45 = styles.Add("s45");
            s45.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s46
            // -----------------------------------------------
            WorksheetStyle s46 = styles.Add("s46");
            s46.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s46.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s47
            // -----------------------------------------------
            WorksheetStyle s47 = styles.Add("s47");
            s47.Font.Color = "#333399";
            s47.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s47.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s48
            // -----------------------------------------------
            WorksheetStyle s48 = styles.Add("s48");
            s48.Font.Bold = true;
            s48.Font.Size = 11;
            s48.Font.Color = "#0000FF";
            s48.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s48.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s48.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s49
            // -----------------------------------------------
            WorksheetStyle s49 = styles.Add("s49");
            s49.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s49.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s50
            // -----------------------------------------------
            WorksheetStyle s50 = styles.Add("s50");
            s50.Font.Bold = true;
            s50.Font.Size = 11;
            s50.Font.Color = "#0000FF";
            s50.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s50.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s51
            // -----------------------------------------------
            WorksheetStyle s51 = styles.Add("s51");
            s51.Font.Color = "#333399";
            s51.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s51.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s51.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s52
            // -----------------------------------------------
            WorksheetStyle s52 = styles.Add("s52");
            s52.Font.Size = 12;
            s52.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s54
            // -----------------------------------------------
            WorksheetStyle s54 = styles.Add("s54");
            s54.Font.Bold = true;
            s54.Font.Size = 11;
            s54.Font.Color = "#0000FF";
            s54.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s54.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s56
            // -----------------------------------------------
            WorksheetStyle s56 = styles.Add("s56");
            s56.Font.Color = "#333399";
            s56.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s56.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s56.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s57
            // -----------------------------------------------
            WorksheetStyle s57 = styles.Add("s57");
            s57.Font.Color = "#333399";
            s57.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s57.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s57.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s57.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s58
            // -----------------------------------------------
            WorksheetStyle s58 = styles.Add("s58");
            s58.Font.Color = "#333399";
            s58.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s58.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s59
            // -----------------------------------------------
            WorksheetStyle s59 = styles.Add("s59");
            s59.Font.Bold = true;
            s59.Font.Size = 12;
            s59.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s59.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s59.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s59.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s60
            // -----------------------------------------------
            WorksheetStyle s60 = styles.Add("s60");
            s60.Font.Bold = true;
            s60.Font.Size = 12;
            s60.Font.Color = "#FF0000";
            s60.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s60.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s60.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s61
            // -----------------------------------------------
            WorksheetStyle s61 = styles.Add("s61");
            s61.Font.Bold = true;
            s61.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s61.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s61.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s62
            // -----------------------------------------------
            WorksheetStyle s62 = styles.Add("s62");
            s62.Font.Size = 12;
            s62.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s62.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s62.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s63
            // -----------------------------------------------
            WorksheetStyle s63 = styles.Add("s63");
            s63.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s63.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s63.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s64
            // -----------------------------------------------
            WorksheetStyle s64 = styles.Add("s64");
            s64.Font.Bold = true;
            s64.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s64.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s64.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s64.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s65
            // -----------------------------------------------
            WorksheetStyle s65 = styles.Add("s65");
            s65.Font.Size = 12;
            s65.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s65.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s65.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s66
            // -----------------------------------------------
            WorksheetStyle s66 = styles.Add("s66");
            s66.Font.Bold = true;
            s66.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s66.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s66.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s66.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s67
            // -----------------------------------------------
            WorksheetStyle s67 = styles.Add("s67");
            s67.Font.Bold = true;
            s67.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s67.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s67.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s68
            // -----------------------------------------------
            WorksheetStyle s68 = styles.Add("s68");
            s68.Font.Bold = true;
            s68.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s68.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s69
            // -----------------------------------------------
            WorksheetStyle s69 = styles.Add("s69");
            s69.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s69.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s69.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s70
            // -----------------------------------------------
            WorksheetStyle s70 = styles.Add("s70");
            s70.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s70.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s71
            // -----------------------------------------------
            WorksheetStyle s71 = styles.Add("s71");
            s71.Font.Size = 12;
            s71.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s71.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s72
            // -----------------------------------------------
            WorksheetStyle s72 = styles.Add("s72");
            s72.Font.Bold = true;
            s72.Interior.Color = "#CCFFCC";
            s72.Interior.Pattern = StyleInteriorPattern.Solid;
            s72.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s72.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s72.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s73
            // -----------------------------------------------
            WorksheetStyle s73 = styles.Add("s73");
            s73.Font.Bold = true;
            s73.Interior.Color = "#CCFFCC";
            s73.Interior.Pattern = StyleInteriorPattern.Solid;
            s73.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s73.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s74
            // -----------------------------------------------
            WorksheetStyle s74 = styles.Add("s74");
            s74.Font.Bold = true;
            s74.Interior.Color = "#CCFFCC";
            s74.Interior.Pattern = StyleInteriorPattern.Solid;
            s74.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s74.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s74.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s75
            // -----------------------------------------------
            WorksheetStyle s75 = styles.Add("s75");
            s75.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s76
            // -----------------------------------------------
            WorksheetStyle s76 = styles.Add("s76");
            s76.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s76.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s76.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s76.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s77
            // -----------------------------------------------
            WorksheetStyle s77 = styles.Add("s77");
            s77.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s77.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s77.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s78
            // -----------------------------------------------
            WorksheetStyle s78 = styles.Add("s78");
            s78.Font.Bold = true;
            s78.Interior.Color = "#CCFFCC";
            s78.Interior.Pattern = StyleInteriorPattern.Solid;
            s78.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s78.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s78.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s78.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s79
            // -----------------------------------------------
            WorksheetStyle s79 = styles.Add("s79");
            s79.Font.Bold = true;
            s79.Interior.Color = "#CCFFCC";
            s79.Interior.Pattern = StyleInteriorPattern.Solid;
            s79.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s79.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s79.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s79.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s80
            // -----------------------------------------------
            WorksheetStyle s80 = styles.Add("s80");
            s80.Font.Color = "#0000FF";
            s80.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s80.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s80.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s80.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s80.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s81
            // -----------------------------------------------
            WorksheetStyle s81 = styles.Add("s81");
            s81.Font.Color = "#0000FF";
            s81.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s81.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s81.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s81.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s81.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s81.NumberFormat = "0";
            // -----------------------------------------------
            //  s88
            // -----------------------------------------------
            WorksheetStyle s88 = styles.Add("s88");
            s88.Font.Color = "#0000FF";
            s88.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s88.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s89
            // -----------------------------------------------
            WorksheetStyle s89 = styles.Add("s89");
            s89.Font.Color = "#0000FF";
            s89.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s89.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s90
            // -----------------------------------------------
            WorksheetStyle s90 = styles.Add("s90");
            s90.Font.Size = 11;
            s90.Font.Color = "#0000FF";
            s90.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s90.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s90.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s90.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s91
            // -----------------------------------------------
            WorksheetStyle s91 = styles.Add("s91");
            s91.Font.Color = "#000000";
            s91.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s91.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s91.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s91.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s92
            // -----------------------------------------------
            WorksheetStyle s92 = styles.Add("s92");
            s92.Font.Color = "#0000FF";
            s92.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s92.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s92.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s92.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s92.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s92.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s93
            // -----------------------------------------------
            WorksheetStyle s93 = styles.Add("s93");
            s93.Font.Color = "#0000FF";
            s93.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s93.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s93.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s93.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s93.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s93.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s93.NumberFormat = "0";
            // -----------------------------------------------
            //  s94
            // -----------------------------------------------
            WorksheetStyle s94 = styles.Add("s94");
            s94.Font.Color = "#0000FF";
            s94.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s94.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s94.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s94.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s94.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s94.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s94.NumberFormat = "0";
            // -----------------------------------------------
            //  s102
            // -----------------------------------------------
            WorksheetStyle s102 = styles.Add("s102");
            s102.Font.Bold = true;
            s102.Font.Color = "#0000FF";
            s102.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s102.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s102.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s102.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s102.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s103
            // -----------------------------------------------
            WorksheetStyle s103 = styles.Add("s103");
            s103.Font.Size = 11;
            s103.Font.Color = "#0000FF";
            s103.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s103.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s103.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s103.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s104
            // -----------------------------------------------
            WorksheetStyle s104 = styles.Add("s104");
            s104.Font.Bold = true;
            s104.Font.Size = 11;
            s104.Font.Color = "#0000FF";
            s104.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s104.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s104.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s104.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s104.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s105
            // -----------------------------------------------
            WorksheetStyle s105 = styles.Add("s105");
            s105.Font.Size = 12;
            s105.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s105.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s105.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s105.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s106
            // -----------------------------------------------
            WorksheetStyle s106 = styles.Add("s106");
            s106.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s106.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s106.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s106.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s107
            // -----------------------------------------------
            WorksheetStyle s107 = styles.Add("s107");
            s107.Font.Bold = true;
            s107.Font.Size = 12;
            s107.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s107.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s107.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s108
            // -----------------------------------------------
            WorksheetStyle s108 = styles.Add("s108");
            s108.Font.Bold = true;
            s108.Font.Size = 14;
            s108.Font.Color = "#FF0000";
            s108.NumberFormat = "0";
            // -----------------------------------------------
            //  s109
            // -----------------------------------------------
            WorksheetStyle s109 = styles.Add("s109");
            s109.Font.Bold = true;
            s109.Font.Size = 12;
            s109.Font.Color = "#0000FF";
            s109.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s110
            // -----------------------------------------------
            WorksheetStyle s110 = styles.Add("s110");
            s110.Font.Bold = true;
            s110.Font.Color = "#000000";
            s110.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s111
            // -----------------------------------------------
            WorksheetStyle s111 = styles.Add("s111");
            s111.Font.Color = "#0000FF";
            s111.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s111.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s111.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s112
            // -----------------------------------------------
            WorksheetStyle s112 = styles.Add("s112");
            s112.Font.Color = "#0000FF";
            s112.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s112.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s112.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s114
            // -----------------------------------------------
            WorksheetStyle s114 = styles.Add("s114");
            s114.Font.Size = 11;
            s114.Font.Color = "#0000FF";
            s114.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s114.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s116
            // -----------------------------------------------
            WorksheetStyle s116 = styles.Add("s116");
            s116.Font.Bold = true;
            s116.Font.Color = "#333399";
            s116.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s116.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s117
            // -----------------------------------------------
            WorksheetStyle s117 = styles.Add("s117");
            s117.Font.Size = 11;
            s117.Font.Color = "#0000FF";
            s117.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s118
            // -----------------------------------------------
            WorksheetStyle s118 = styles.Add("s118");
            s118.Font.Bold = true;
            s118.Font.Size = 11;
            s118.Font.Color = "#0000FF";
            s118.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s118.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s119
            // -----------------------------------------------
            WorksheetStyle s119 = styles.Add("s119");
            s119.Font.Bold = true;
            s119.Font.Size = 11;
            s119.Font.Color = "#0000FF";
            s119.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s119.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s119.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s120
            // -----------------------------------------------
            WorksheetStyle s120 = styles.Add("s120");
            s120.Font.Bold = true;
            s120.Font.Size = 12;
            s120.Font.Color = "#FF0000";
            // -----------------------------------------------
            //  s121
            // -----------------------------------------------
            WorksheetStyle s121 = styles.Add("s121");
            s121.Font.Bold = true;
            s121.Font.Size = 14;
            s121.Font.Color = "#FF0000";
            s121.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s122
            // -----------------------------------------------
            WorksheetStyle s122 = styles.Add("s122");
            s122.Font.Size = 12;
            s122.Font.Color = "#FF0000";
            // -----------------------------------------------
            //  s123
            // -----------------------------------------------
            WorksheetStyle s123 = styles.Add("s123");
            // -----------------------------------------------
            //  s125
            // -----------------------------------------------
            WorksheetStyle s125 = styles.Add("s125");
            s125.Font.Bold = true;
            s125.Font.Size = 12;
            s125.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s125.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s125.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s129
            // -----------------------------------------------
            WorksheetStyle s129 = styles.Add("s129");
            s129.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s129.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s129.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s129.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s130
            // -----------------------------------------------
            WorksheetStyle s130 = styles.Add("s130");
            s130.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s131
            // -----------------------------------------------
            WorksheetStyle s131 = styles.Add("s131");
            s131.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s131.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s132
            // -----------------------------------------------
            WorksheetStyle s132 = styles.Add("s132");
            s132.Font.Size = 14;
            s132.Font.Color = "#FF0000";
            s132.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s134
            // -----------------------------------------------
            WorksheetStyle s134 = styles.Add("s134");
            s134.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s134.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s134.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s135
            // -----------------------------------------------
            WorksheetStyle s135 = styles.Add("s135");
            s135.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s135.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s136
            // -----------------------------------------------
            WorksheetStyle s136 = styles.Add("s136");
            s136.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s136.Alignment.WrapText = true;
            // -----------------------------------------------
            //  s149
            // -----------------------------------------------
            WorksheetStyle s149 = styles.Add("s149");
            s149.Font.Color = "#0000FF";
            s149.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s149.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s149.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s149.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s149.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s150
            // -----------------------------------------------
            WorksheetStyle s150 = styles.Add("s150");
            s150.Font.Color = "#0000FF";
            s150.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s150.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s150.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s150.NumberFormat = "0";
            // -----------------------------------------------
            //  s151
            // -----------------------------------------------
            WorksheetStyle s151 = styles.Add("s151");
            s151.Font.Size = 12;
            s151.NumberFormat = "0";
            // -----------------------------------------------
            //  s156
            // -----------------------------------------------
            WorksheetStyle s156 = styles.Add("s156");
            s156.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s156.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s156.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s156.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s156.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s156.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s169
            // -----------------------------------------------
            WorksheetStyle s169 = styles.Add("s169");
            s169.Font.Color = "#333399";
            s169.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s169.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s169.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s169.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s169.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s170
            // -----------------------------------------------
            WorksheetStyle s170 = styles.Add("s170");
            s170.Font.Color = "#333399";
            s170.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s170.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s170.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s170.NumberFormat = "0";
            // -----------------------------------------------
            //  s171
            // -----------------------------------------------
            WorksheetStyle s171 = styles.Add("s171");
            s171.Font.Bold = true;
            s171.Font.Size = 12;
            s171.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s172
            // -----------------------------------------------
            WorksheetStyle s172 = styles.Add("s172");
            s172.Font.Bold = true;
            s172.Font.Size = 12;
            s172.Font.Color = "#0000FF";
            s172.NumberFormat = "0";
            // -----------------------------------------------
            //  s173
            // -----------------------------------------------
            WorksheetStyle s173 = styles.Add("s173");
            s173.Font.Bold = true;
            s173.Font.Color = "#000000";
            // -----------------------------------------------
            //  s174
            // -----------------------------------------------
            WorksheetStyle s174 = styles.Add("s174");
            s174.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s174.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s175
            // -----------------------------------------------
            WorksheetStyle s175 = styles.Add("s175");
            s175.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s176
            // -----------------------------------------------
            WorksheetStyle s176 = styles.Add("s176");
            s176.Font.Size = 12;
            s176.Font.Color = "#FF0000";
            s176.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s177
            // -----------------------------------------------
            WorksheetStyle s177 = styles.Add("s177");
            s177.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s177.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s178
            // -----------------------------------------------
            WorksheetStyle s178 = styles.Add("s178");
            s178.Font.Color = "#0000FF";
            s178.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s178.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s178.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s179
            // -----------------------------------------------
            WorksheetStyle s179 = styles.Add("s179");
            s179.Font.Color = "#0000FF";
            s179.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s180
            // -----------------------------------------------
            WorksheetStyle s180 = styles.Add("s180");
            s180.Font.Size = 11;
            s180.Font.Color = "#0000FF";
            s180.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s181
            // -----------------------------------------------
            WorksheetStyle s181 = styles.Add("s181");
            s181.Font.Bold = true;
            s181.Font.Size = 11;
            s181.Font.Color = "#0000FF";
            s181.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s181.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s182
            // -----------------------------------------------
            WorksheetStyle s182 = styles.Add("s182");
            s182.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s182.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s182.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s183
            // -----------------------------------------------
            WorksheetStyle s183 = styles.Add("s183");
            s183.Font.Size = 12;
            s183.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s183.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  s184
            // -----------------------------------------------
            WorksheetStyle s184 = styles.Add("s184");
            s184.Font.Size = 11;
            s184.Font.Color = "#0000FF";
            s184.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s185
            // -----------------------------------------------
            WorksheetStyle s185 = styles.Add("s185");
            s185.Font.Color = "#0000FF";
            s185.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s185.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s185.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s185.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s185.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s185.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s185.NumberFormat = "@";
            // -----------------------------------------------
            //  s193
            // -----------------------------------------------
            WorksheetStyle s193 = styles.Add("s193");
            s193.Font.Color = "#0000FF";
            s193.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s193.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s193.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s193.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s194
            // -----------------------------------------------
            WorksheetStyle s194 = styles.Add("s194");
            s194.Font.Color = "#333399";
            s194.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s194.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s194.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s194.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s195
            // -----------------------------------------------
            WorksheetStyle s195 = styles.Add("s195");
            s195.Font.Bold = true;
            s195.Font.Size = 14;
            // -----------------------------------------------
            //  s196
            // -----------------------------------------------
            WorksheetStyle s196 = styles.Add("s196");
            s196.Font.Bold = true;
            s196.Font.Size = 14;
            s196.NumberFormat = "0";
            // -----------------------------------------------
            //  s197
            // -----------------------------------------------
            WorksheetStyle s197 = styles.Add("s197");
            s197.Font.Bold = true;
            s197.Font.Size = 12;
            s197.Font.Color = "#000000";
            // -----------------------------------------------
            //  s198
            // -----------------------------------------------
            WorksheetStyle s198 = styles.Add("s198");
            s198.Font.Color = "#0000FF";
            s198.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s198.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s198.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s198.NumberFormat = "@";
            // -----------------------------------------------
            //  s200
            // -----------------------------------------------
            WorksheetStyle s200 = styles.Add("s200");
            s200.Font.Color = "#333399";
            s200.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s201
            // -----------------------------------------------
            WorksheetStyle s201 = styles.Add("s201");
            s201.Font.Color = "#339966";
            s201.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s202
            // -----------------------------------------------
            WorksheetStyle s202 = styles.Add("s202");
            s202.Font.Size = 12;
            s202.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s203
            // -----------------------------------------------
            WorksheetStyle s203 = styles.Add("s203");
            s203.Font.Bold = true;
            s203.Font.Size = 12;
            s203.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s203.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s204
            // -----------------------------------------------
            WorksheetStyle s204 = styles.Add("s204");
            s204.Font.Size = 12;
            s204.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s204.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s205
            // -----------------------------------------------
            WorksheetStyle s205 = styles.Add("s205");
            s205.Font.Size = 12;
            s205.Font.Color = "#0000FF";
            s205.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s206
            // -----------------------------------------------
            WorksheetStyle s206 = styles.Add("s206");
            s206.Font.Size = 12;
            s206.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s207
            // -----------------------------------------------
            WorksheetStyle s207 = styles.Add("s207");
            s207.Font.Size = 12;
            s207.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s207.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s208
            // -----------------------------------------------
            WorksheetStyle s208 = styles.Add("s208");
            s208.Font.Size = 12;
            s208.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s208.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s209
            // -----------------------------------------------
            WorksheetStyle s209 = styles.Add("s209");
            s209.Font.Size = 12;
            s209.Font.Color = "#000000";
            // -----------------------------------------------
            //  s210
            // -----------------------------------------------
            WorksheetStyle s210 = styles.Add("s210");
            s210.Font.Color = "#000000";
            // -----------------------------------------------
            //  s211
            // -----------------------------------------------
            WorksheetStyle s211 = styles.Add("s211");
            s211.Font.Size = 12;
            s211.Font.Color = "#000000";
            s211.NumberFormat = "0.0";
            // -----------------------------------------------
            //  s212
            // -----------------------------------------------
            WorksheetStyle s212 = styles.Add("s212");
            s212.Font.Bold = true;
            s212.Interior.Color = "#CCFFCC";
            s212.Interior.Pattern = StyleInteriorPattern.Solid;
            s212.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s212.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s212.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s212.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s212.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s213
            // -----------------------------------------------
            WorksheetStyle s213 = styles.Add("s213");
            s213.Interior.Color = "#CCFFCC";
            s213.Interior.Pattern = StyleInteriorPattern.Solid;
            s213.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s213.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s214
            // -----------------------------------------------
            WorksheetStyle s214 = styles.Add("s214");
            s214.Interior.Color = "#CCFFCC";
            s214.Interior.Pattern = StyleInteriorPattern.Solid;
            s214.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s214.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s214.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s215
            // -----------------------------------------------
            WorksheetStyle s215 = styles.Add("s215");
            s215.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s216
            // -----------------------------------------------
            WorksheetStyle s216 = styles.Add("s216");
            s216.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s216.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s216.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            s216.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s217
            // -----------------------------------------------
            WorksheetStyle s217 = styles.Add("s217");
            s217.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s217.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s217.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s217.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s218
            // -----------------------------------------------
            WorksheetStyle s218 = styles.Add("s218");
            s218.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s218.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s218.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s218.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s218.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
        }


        /// <summary>
        /// GenerateWorksheetBlad1: Generate Worksheet 1
        /// </summary>
        /// <param name="sheets"></param>
        private void GenerateWorksheetBlad1(WorksheetCollection sheets)
        {
            Worksheet sheet = sheets.Add("Blad1");
            sheet.Table.DefaultRowHeight = 15F;
            sheet.Table.ExpandedColumnCount = 14;
            sheet.Table.ExpandedRowCount = 631;
            sheet.Table.FullColumns = 1;
            sheet.Table.FullRows = 1;

            WorksheetColumn column0 = sheet.Table.Columns.Add(45);
            WorksheetColumn column1 = sheet.Table.Columns.Add(45);
            WorksheetColumn column2 = sheet.Table.Columns.Add(60);
            WorksheetColumn column3 = sheet.Table.Columns.Add(30);
            WorksheetColumn column4 = sheet.Table.Columns.Add(50);

            WorksheetColumn column5 = sheet.Table.Columns.Add();
            column5.Index = 6;
            column5.Width = 45;

            WorksheetColumn column6 = sheet.Table.Columns.Add(40);
            WorksheetColumn column7 = sheet.Table.Columns.Add(40);
            WorksheetColumn column8 = sheet.Table.Columns.Add(40);

            WorksheetColumn column9 = sheet.Table.Columns.Add();
            column9.Index = 10;
            column9.Width = 150;

            WorksheetColumn column10 = sheet.Table.Columns.Add(45);
            WorksheetColumn column11 = sheet.Table.Columns.Add(45);
            WorksheetColumn column12 = sheet.Table.Columns.Add(45);
            WorksheetColumn column13 = sheet.Table.Columns.Add(45);

            // -----------------------------------------------
            WorksheetRow Row0 = sheet.Table.Rows.Add();
            Row0.Height = 20;
            Row0.AutoFitHeight = false;
            WorksheetCell cell;
            cell = Row0.Cells.Add();
            cell.StyleID = "s21";
            cell = Row0.Cells.Add();
            cell.StyleID = "s22";
            cell.Data.Type = DataType.String;
            cell.Data.Text = "Återrapportering/Fakturaunderlag SG-systemet";
            cell.Index = 4;
            cell = Row0.Cells.Add();
            cell.StyleID = "s21";
            cell = Row0.Cells.Add();
            cell.StyleID = "s21";
            cell = Row0.Cells.Add();
            cell.StyleID = "s21";
            cell = Row0.Cells.Add();
            cell.StyleID = "s21";
            cell = Row0.Cells.Add();
            cell.StyleID = "s21";
            cell = Row0.Cells.Add();
            cell.StyleID = "s21";
            cell = Row0.Cells.Add();
            cell.StyleID = "s21";
            cell = Row0.Cells.Add();
            cell.StyleID = "s21";
            cell = Row0.Cells.Add();
            cell.StyleID = "s21";
            // -----------------------------------------------
            WorksheetRow Row1 = sheet.Table.Rows.Add();
            Row1.AutoFitHeight = false;
            cell = Row1.Cells.Add();
            cell.StyleID = "s21";
            cell = Row1.Cells.Add();
            cell.StyleID = "s22";
            cell.Index = 4;
            cell = Row1.Cells.Add();
            cell.StyleID = "s23";
            cell.Index = 6;
            cell = Row1.Cells.Add();
            cell.StyleID = "s21";
            cell = Row1.Cells.Add();
            cell.StyleID = "s21";
            cell = Row1.Cells.Add();
            cell.StyleID = "s21";
            cell = Row1.Cells.Add();
            cell.StyleID = "s21";
            cell = Row1.Cells.Add();
            cell.StyleID = "s21";
            cell = Row1.Cells.Add();
            cell.StyleID = "s24";
            cell = Row1.Cells.Add();
            cell.StyleID = "s24";
            // -----------------------------------------------
            WorksheetRow Row2 = sheet.Table.Rows.Add();
            Row2.AutoFitHeight = false;
            Row2.Cells.Add("Ordernummer:", DataType.String, "s25");

            //cell = Row2.Cells.Add();
            //cell.StyleID = "s21";

            cell = Row2.Cells.Add();
            cell.StyleID = "s26";
            cell.Data.Type = DataType.Number;
            cell.Data.Text = drowOfCData["Ordernr"].ToString().Replace(',', '.').ToString();
            cell.Index = 3;    
            
            Row2.Cells.Add("Orderdatum:", DataType.String, "s25");

            //cell = Row2.Cells.Add();
            //cell.StyleID = "s21";

            cell = Row2.Cells.Add();
            cell.StyleID = "s26";
            cell.Data.Type = DataType.String;
            cell.Data.Text = Convert.ToDateTime(drowOfCData["Beställningsdatum"]).ToShortDateString();
            cell.Index = 6;    
            
            //cell = Row2.Cells.Add();
            //cell.StyleID = "s22";
            //cell = Row2.Cells.Add();
            //cell.StyleID = "s21";
            cell = Row2.Cells.Add();
            cell.StyleID = "s21";
            cell = Row2.Cells.Add();
            cell.StyleID = "s21";
            cell = Row2.Cells.Add();
            cell.StyleID = "s21";
            cell = Row2.Cells.Add();
            cell.StyleID = "s21";
            cell = Row2.Cells.Add();
            cell.StyleID = "s21";
            cell = Row2.Cells.Add();
            cell.StyleID = "s21";
            cell = Row2.Cells.Add();
            cell.StyleID = "s21";
            cell = Row2.Cells.Add();
            cell.StyleID = "s21";
            // -----------------------------------------------
            WorksheetRow Row3 = sheet.Table.Rows.Add();
            Row3.AutoFitHeight = false;
            cell = Row3.Cells.Add();
            cell.StyleID = "s25";
            cell = Row3.Cells.Add();
            cell.StyleID = "s27";
            cell.Index = 3;
            cell = Row3.Cells.Add();
            cell.StyleID = "s22";
            cell = Row3.Cells.Add();
            cell.StyleID = "s21";
            cell = Row3.Cells.Add();
            cell.StyleID = "s21";
            cell = Row3.Cells.Add();
            cell.StyleID = "s21";
            cell = Row3.Cells.Add();
            cell.StyleID = "s21";
            cell = Row3.Cells.Add();
            cell.StyleID = "s21";
            cell = Row3.Cells.Add();
            cell.StyleID = "s21";
            cell = Row3.Cells.Add();
            cell.StyleID = "s21";
            cell = Row3.Cells.Add();
            cell.StyleID = "s21";
            cell = Row3.Cells.Add();
            cell.StyleID = "s21";
            // -----------------------------------------------
            WorksheetRow Row4 = sheet.Table.Rows.Add();
            Row4.AutoFitHeight = false;
            Row4.Cells.Add("Företag", DataType.String, "s28");
            cell = Row4.Cells.Add();
            cell.StyleID = "s29";
            cell = Row4.Cells.Add();
            cell.StyleID = "s29";
            Row4.Cells.Add("Fakturaadress", DataType.String, "s30");
            cell = Row4.Cells.Add();
            cell.StyleID = "s30";
            cell = Row4.Cells.Add();
            cell.StyleID = "s30";
            Row4.Cells.Add("Postnr", DataType.String, "s30");
            Row4.Cells.Add("Postadress", DataType.String, "s30");
            cell = Row4.Cells.Add();
            cell.StyleID = "s30";
            Row4.Cells.Add("Beställningsreferens", DataType.String, "s30");
            cell = Row4.Cells.Add();
            cell.StyleID = "s31";
            cell = Row4.Cells.Add();
            cell.StyleID = "s31";
            cell = Row4.Cells.Add();
            cell.StyleID = "s32";
            // -----------------------------------------------
            WorksheetRow RowV4 = sheet.Table.Rows.Add();
            RowV4.AutoFitHeight = false;
            RowV4.Cells.Add(drowOfCData["Företagsnamn"].ToString(), DataType.String, "s33");
            cell = RowV4.Cells.Add();
            cell.StyleID = "s34";
            cell = RowV4.Cells.Add();
            cell.StyleID = "s34";
            RowV4.Cells.Add(drowOfCData["Faktureringsadress"].ToString(), DataType.String, "s35");
            cell = RowV4.Cells.Add();
            cell.StyleID = "s35";
            cell = RowV4.Cells.Add();
            cell.StyleID = "s35";
            RowV4.Cells.Add(drowOfCData["Postnummer"].ToString(), DataType.String, "s36");
            RowV4.Cells.Add(drowOfCData["Ort"].ToString(), DataType.String, "s35");
            cell = RowV4.Cells.Add();
            cell.StyleID = "s35";
            RowV4.Cells.Add(drowOfCData["Beställningsreferens"].ToString(), DataType.String, "s35");
            cell = RowV4.Cells.Add();
            cell.StyleID = "s37";
            cell = RowV4.Cells.Add();
            cell.StyleID = "s21";
            cell = RowV4.Cells.Add();
            cell.StyleID = "s38";
            // -----------------------------------------------
            WorksheetRow Row6 = sheet.Table.Rows.Add();
            Row6.AutoFitHeight = false;
            cell = Row6.Cells.Add();
            cell.StyleID = "s33";
            cell = Row6.Cells.Add();
            cell.StyleID = "s34";
            cell = Row6.Cells.Add();
            cell.StyleID = "s34";
            cell = Row6.Cells.Add();
            cell.StyleID = "s35";
            cell = Row6.Cells.Add();
            cell.StyleID = "s35";
            cell = Row6.Cells.Add();
            cell.StyleID = "s35";
            cell = Row6.Cells.Add();
            cell.StyleID = "s36";
            cell = Row6.Cells.Add();
            cell.StyleID = "s35";
            cell = Row6.Cells.Add();
            cell.StyleID = "s35";
            cell = Row6.Cells.Add();
            cell.StyleID = "s35";
            cell = Row6.Cells.Add();
            cell.StyleID = "s21";
            cell = Row6.Cells.Add();
            cell.StyleID = "s21";
            cell = Row6.Cells.Add();
            cell.StyleID = "s38";
            // -----------------------------------------------
            WorksheetRow Row7 = sheet.Table.Rows.Add();
            Row7.AutoFitHeight = false;
            Row7.Cells.Add("Region/Förvaltning", DataType.String, "s39");
            cell = Row7.Cells.Add();
            cell.StyleID = "s40";
            cell = Row7.Cells.Add();
            cell.StyleID = "s40";
            Row7.Cells.Add("Distrikt/Område", DataType.String, "s40");
            cell = Row7.Cells.Add();
            cell.StyleID = "s40";
            cell = Row7.Cells.Add();
            cell.StyleID = "s40";
            Row7.Cells.Add("VAT", DataType.String, "s40");
            cell = Row7.Cells.Add();
            cell.StyleID = "s40";
            cell = Row7.Cells.Add();
            cell.StyleID = "s40";
            cell = Row7.Cells.Add();
            cell.StyleID = "s40";
            cell = Row7.Cells.Add();
            cell.StyleID = "s41";
            cell = Row7.Cells.Add();
            cell.StyleID = "s41";
            cell = Row7.Cells.Add();
            cell.StyleID = "s42";
            // -----------------------------------------------
            WorksheetRow Row8 = sheet.Table.Rows.Add();
            Row8.AutoFitHeight = false;
            Row8.Cells.Add(drowOfCData["Region_Förvaltning"].ToString(), DataType.String, "s33");
            cell = Row8.Cells.Add();
            cell.StyleID = "s34";
            cell = Row8.Cells.Add();
            cell.StyleID = "s34";
            Row8.Cells.Add(drowOfCData["Distrikt_Område"].ToString(), DataType.String, "s35");
            cell = Row8.Cells.Add();
            cell.StyleID = "s35";
            cell = Row8.Cells.Add();
            cell.StyleID = "s35";
            Row8.Cells.Add(drowOfCData["VAT"].ToString(), DataType.String, "s35");
            cell = Row8.Cells.Add();
            cell.StyleID = "s35";
            cell = Row8.Cells.Add();
            cell.StyleID = "s36";
            cell = Row8.Cells.Add();
            cell.StyleID = "s36";
            cell = Row8.Cells.Add();
            cell.StyleID = "s43";
            cell = Row8.Cells.Add();
            cell.StyleID = "s43";
            cell = Row8.Cells.Add();
            cell.StyleID = "s44";
            // -----------------------------------------------
            WorksheetRow Row9 = sheet.Table.Rows.Add();
            Row9.AutoFitHeight = false;
            cell = Row9.Cells.Add();
            cell.StyleID = "s33";
            cell = Row9.Cells.Add();
            cell.StyleID = "s45";
            cell = Row9.Cells.Add();
            cell.StyleID = "s45";
            cell = Row9.Cells.Add();
            cell.StyleID = "s35";
            cell = Row9.Cells.Add();
            cell.StyleID = "s45";
            cell = Row9.Cells.Add();
            cell.StyleID = "s45";
            cell = Row9.Cells.Add();
            cell.StyleID = "s35";
            cell = Row9.Cells.Add();
            cell.StyleID = "s45";
            cell = Row9.Cells.Add();
            cell.StyleID = "s36";
            cell = Row9.Cells.Add();
            cell.StyleID = "s36";
            cell = Row9.Cells.Add();
            cell.StyleID = "s43";
            cell = Row9.Cells.Add();
            cell.StyleID = "s43";
            cell = Row9.Cells.Add();
            cell.StyleID = "s44";
            // -----------------------------------------------
            WorksheetRow Row10 = sheet.Table.Rows.Add();
            Row10.AutoFitHeight = false;
            Row10.Cells.Add("Kontaktman 1", DataType.String, "s39");
            cell = Row10.Cells.Add();
            cell.StyleID = "s40";
            cell = Row10.Cells.Add();
            cell.StyleID = "s40";
            cell = Row10.Cells.Add();
            cell.StyleID = "s40";
            cell = Row10.Cells.Add();
            cell.StyleID = "s40";
            Row10.Cells.Add("Kontaktman 2", DataType.String, "s40");
            cell = Row10.Cells.Add();
            cell.StyleID = "s40";
            cell = Row10.Cells.Add();
            cell.StyleID = "s40";
            cell = Row10.Cells.Add();
            cell.StyleID = "s40";
            cell = Row10.Cells.Add();
            cell.StyleID = "s46";
            cell = Row10.Cells.Add();
            cell.StyleID = "s41";
            cell = Row10.Cells.Add();
            cell.StyleID = "s41";
            cell = Row10.Cells.Add();
            cell.StyleID = "s42";
            // -----------------------------------------------
            WorksheetRow Row11 = sheet.Table.Rows.Add();
            Row11.AutoFitHeight = false;
            Row11.Cells.Add(drowOfCData["Kontaktperson1"].ToString(), DataType.String, "s33");
            cell = Row11.Cells.Add();
            cell.StyleID = "s35";
            cell = Row11.Cells.Add();
            cell.StyleID = "s35";
            cell = Row11.Cells.Add();
            cell.StyleID = "s35";
            cell = Row11.Cells.Add();
            cell.StyleID = "s35";
            Row11.Cells.Add(drowOfCData["Kontaktperson2"].ToString(), DataType.String, "s35");
            cell = Row11.Cells.Add();
            cell.StyleID = "s36";
            cell = Row11.Cells.Add();
            cell.StyleID = "s35";
            cell = Row11.Cells.Add();
            cell.StyleID = "s35";
            cell = Row11.Cells.Add();
            cell.StyleID = "s35";
            cell = Row11.Cells.Add();
            cell.StyleID = "s21";
            cell = Row11.Cells.Add();
            cell.StyleID = "s21";
            cell = Row11.Cells.Add();
            cell.StyleID = "s38";
            // -----------------------------------------------
            WorksheetRow Row12 = sheet.Table.Rows.Add();
            Row12.AutoFitHeight = false;
            Row12.Cells.Add("Tfn arb", DataType.String, "s39");
            Row12.Cells.Add(drowOfCData["TelefonArb1"].ToString(), DataType.String, "s35");
            cell = Row12.Cells.Add();
            cell.StyleID = "s35";
            cell = Row12.Cells.Add();
            cell.StyleID = "s35";
            cell = Row12.Cells.Add();
            cell.StyleID = "s40";
            Row12.Cells.Add("Tfn arb", DataType.String, "s40");
            Row12.Cells.Add(drowOfCData["TelefonArb2"].ToString(), DataType.String, "s35");
            cell = Row12.Cells.Add();
            cell.StyleID = "s40";
            cell = Row12.Cells.Add();
            cell.StyleID = "s40";
            cell = Row12.Cells.Add();
            cell.StyleID = "s40";
            cell = Row12.Cells.Add();
            cell.StyleID = "s21";
            cell = Row12.Cells.Add();
            cell.StyleID = "s21";
            cell = Row12.Cells.Add();
            cell.StyleID = "s38";
            // -----------------------------------------------
            WorksheetRow Row13 = sheet.Table.Rows.Add();
            Row13.AutoFitHeight = false;
            Row13.Cells.Add("Tfn mobil", DataType.String, "s39");
            Row13.Cells.Add(drowOfCData["TelefonMobil1"].ToString(), DataType.String, "s35");
            cell = Row13.Cells.Add();
            cell.StyleID = "s35";
            cell = Row13.Cells.Add();
            cell.StyleID = "s35";
            cell = Row13.Cells.Add();
            cell.StyleID = "s35";
            Row13.Cells.Add("Tfn mobil", DataType.String, "s40");
            Row13.Cells.Add(drowOfCData["TelefonMobil2"].ToString(), DataType.String, "s35");
            cell = Row13.Cells.Add();
            cell.StyleID = "s45";
            cell = Row13.Cells.Add();
            cell.StyleID = "s47";
            cell = Row13.Cells.Add();
            cell.StyleID = "s47";
            cell = Row13.Cells.Add();
            cell.StyleID = "s21";
            cell = Row13.Cells.Add();
            cell.StyleID = "s21";
            cell = Row13.Cells.Add();
            cell.StyleID = "s38";
            // -----------------------------------------------
            WorksheetRow Row14 = sheet.Table.Rows.Add();
            Row14.AutoFitHeight = false;
            Row14.Cells.Add("Tfn bost", DataType.String, "s39");
            Row14.Cells.Add(drowOfCData["TelefonHem1"].ToString(), DataType.String, "s35");
            cell = Row14.Cells.Add();
            cell.StyleID = "s35";
            cell = Row14.Cells.Add();
            cell.StyleID = "s35";
            cell = Row14.Cells.Add();
            cell.StyleID = "s35";
            Row14.Cells.Add("Tfn bost", DataType.String, "s40");
            Row14.Cells.Add(drowOfCData["TelefonHem2"].ToString(), DataType.String, "s35");
            cell = Row14.Cells.Add();
            cell.StyleID = "s45";
            cell = Row14.Cells.Add();
            cell.StyleID = "s47";
            cell = Row14.Cells.Add();
            cell.StyleID = "s47";
            cell = Row14.Cells.Add();
            cell.StyleID = "s21";
            cell = Row14.Cells.Add();
            cell.StyleID = "s21";
            cell = Row14.Cells.Add();
            cell.StyleID = "s38";
            // -----------------------------------------------
            WorksheetRow Row15 = sheet.Table.Rows.Add();
            Row15.AutoFitHeight = false;
            
            /* added */
            cell = Row15.Cells.Add("e-mail", DataType.String, "s39");
            Row15.Cells.Add(drowOfCData["Epostadress1"].ToString(), DataType.String, "s35");
            cell = Row15.Cells.Add();
            cell.StyleID = "s49";
            cell = Row15.Cells.Add();
            cell.StyleID = "s49";
            cell = Row15.Cells.Add();
            cell.StyleID = "s50";
            cell = Row15.Cells.Add("e-mail", DataType.String, "s40");
            Row15.Cells.Add(drowOfCData["Epostadress2"].ToString(), DataType.String, "s35");
            /* added */
            
            //cell = Row15.Cells.Add();
            //cell.StyleID = "s48";
            //cell = Row15.Cells.Add();
            //cell.StyleID = "s49";
            //cell = Row15.Cells.Add();
            //cell.StyleID = "s49";
            //cell = Row15.Cells.Add();
            //cell.StyleID = "s50";
            //cell = Row15.Cells.Add();
            //cell.StyleID = "s49";
            //cell = Row15.Cells.Add();
            //cell.StyleID = "s49";
            cell = Row15.Cells.Add();
            cell.StyleID = "s49";
            cell = Row15.Cells.Add();
            cell.StyleID = "s51";
            cell = Row15.Cells.Add();
            cell.StyleID = "s51";
            cell = Row15.Cells.Add();
            cell.StyleID = "s52";
            cell = Row15.Cells.Add();
            cell.StyleID = "s52";
            cell = Row15.Cells.Add();
            cell.StyleID = "s38";
            // -----------------------------------------------
            WorksheetRow Row16 = sheet.Table.Rows.Add();
            Row16.AutoFitHeight = false;
            cell = Row16.Cells.Add();
            cell.StyleID = "s54";
            cell.MergeAcross = 2;
            cell = Row16.Cells.Add();
            cell.StyleID = "s54";
            cell.MergeAcross = 2;
            cell = Row16.Cells.Add();
            cell.StyleID = "s54";
            cell.MergeAcross = 1;
            cell = Row16.Cells.Add();
            cell.StyleID = "s56";
            cell.MergeAcross = 1;
            cell = Row16.Cells.Add();
            cell.StyleID = "s56";
            cell.MergeAcross = 1;
            cell = Row16.Cells.Add();
            cell.StyleID = "s57";
            cell = Row16.Cells.Add();
            cell.StyleID = "s58";
            // -----------------------------------------------
            WorksheetRow Row17 = sheet.Table.Rows.Add();
            Row17.AutoFitHeight = false;
            Row17.Cells.Add("Gödslingsobjekt", DataType.String, "s59");
            cell = Row17.Cells.Add();
            cell.StyleID = "s60";
            cell = Row17.Cells.Add();
            cell.StyleID = "s60";
            cell = Row17.Cells.Add();
            cell.StyleID = "s61";
            cell = Row17.Cells.Add();
            cell.StyleID = "s61";
            cell = Row17.Cells.Add();
            cell.StyleID = "s62";
            cell = Row17.Cells.Add();
            cell.StyleID = "s31";
            Row17.Cells.Add("Netto-", DataType.String, "s63");
            Row17.Cells.Add("Skog-", DataType.String, "s63");

            /* added*/
            cell = Row17.Cells.Add();
            cell.StyleID = "s31";
// HÄR
            //cell = Row17.Cells.Add();
            //cell.StyleID = "s64";
            Row17.Cells.Add("Spridn.gruppens noteringar", DataType.String, "s212");

            cell = Row17.Cells.Add();
            cell.StyleID = "s213";
            cell = Row17.Cells.Add();
            cell.StyleID = "s214";
            // -----------------------------------------------
            WorksheetRow Row18 = sheet.Table.Rows.Add();
            Row18.AutoFitHeight = false;
            Row18.Cells.Add("Objekt", DataType.String, "s69");
            Row18.Cells.Add("Start-", DataType.String, "s217");
            cell = Row18.Cells.Add("Avdelning", DataType.String, "s217");
            cell.MergeAcross = 1;
            cell = Row18.Cells.Add("Avdelning", DataType.String, "s217");
            cell.MergeAcross = 1;
            Row18.Cells.Add("Giva ", DataType.String, "s217");
            Row18.Cells.Add("areal", DataType.String, "s217");
            Row18.Cells.Add("CAN", DataType.String, "s217");
            Row18.Cells.Add("Kommentar", DataType.String, "s217");
            Row18.Cells.Add("Spritt", DataType.String, "s72");
            Row18.Cells.Add("Väg-Obj.", DataType.String, "s73");
            Row18.Cells.Add("Datum", DataType.String, "s74");
            // -----------------------------------------------
            WorksheetRow Row19 = sheet.Table.Rows.Add();
            Row19.AutoFitHeight = false;
            Row19.Cells.Add("Nr", DataType.String, "s218");
            Row19.Cells.Add("plats", DataType.String, "s218");
            cell = Row19.Cells.Add("Nr", DataType.String, "s218");
            cell.MergeAcross = 1;
            cell = Row19.Cells.Add("Namn", DataType.String, "s218");
            cell.MergeAcross = 1;
            Row19.Cells.Add("KgN/ha", DataType.String, "s218");
            Row19.Cells.Add("hektar", DataType.String, "s218");
            Row19.Cells.Add("ton", DataType.String, "s218");
            cell = Row19.Cells.Add();
            cell.StyleID = "s71";
            Row19.Cells.Add("ton", DataType.String, "s78");
            Row19.Cells.Add("m", DataType.String, "s73");
            cell = Row19.Cells.Add();
            cell.StyleID = "s79";
            cell = Row19.Cells.Add();

            string OrderNR = drowOfCData["Ordernr"].ToString();
            DataSet myStartPlaceDataSet = Accessdatabas.LäsIfrånDatabas("Select * from Startplats where Ordernr = " + OrderNR + " order by Startplats");
            DataSet myObjectDataSet;

             // Get the table from the data set
            int[] startPlacesSorted = SortNumbersAsWell(myStartPlaceDataSet.Tables[0]);
            DataTable dStartData = myStartPlaceDataSet.Tables[0];


            DataTable dObjectData;

            WorksheetRow Row20;

            int numberOfObjects = 1;

            for (int i = 0; i < dStartData.Rows.Count; i++)
            {
                DataRow drowOfStartData = dStartData.Rows[startPlacesSorted[i]];

                // Only row that have not been deleted
                if (drowOfStartData.RowState != DataRowState.Deleted)
                {
                    string startPlatsID = drowOfStartData["Startplats"].ToString();
                    myObjectDataSet = Accessdatabas.LäsIfrånDatabas("Select * from Objekt where Ordernr = " + OrderNR + " and Startplats = '" + startPlatsID + "'");
                    dObjectData = myObjectDataSet.Tables[0];

                    for (int j = 0; j < dObjectData.Rows.Count; j++)
                    {


                        DataRow drowOfObjectData = dObjectData.Rows[j];

                        // Only row that have not been deleted
                        if (drowOfObjectData.RowState != DataRowState.Deleted)
                        {
                            numberOfObjects++;

                            Row20 = sheet.Table.Rows.Add();
                            Row20.AutoFitHeight = false;
                            Row20.Cells.Add(drowOfObjectData["Objektnr"].ToString().Replace(',', '.').ToString(), DataType.Number, "s80");
                            Row20.Cells.Add(drowOfObjectData["Startplats"].ToString(), DataType.String, "s80");
                            cell = Row20.Cells.Add(drowOfObjectData["Avdnr"].ToString(), DataType.String, "s81");
                            cell.MergeAcross = 1;
                            cell = Row20.Cells.Add();
                            cell.StyleID = "m19466902";
                            cell.Data.Type = DataType.String;
                            cell.Data.Text = drowOfObjectData["Avdnamn"].ToString();
                            cell.MergeAcross = 1;
                            Row20.Cells.Add(drowOfObjectData["Giva_KgN_ha"].ToString().Replace(',', '.').ToString(), DataType.Number, "s89");
                            Row20.Cells.Add(drowOfObjectData["Areal_ha"].ToString().Replace(',', '.').ToString(), DataType.Number, "s90");
                            Row20.Cells.Add(drowOfObjectData["Skog_CAN_ton"].ToString().Replace(',', '.').ToString(), DataType.Number, "s90");
                           
                            /* added */
                            //LSAM added 2009-01-02
                            cell = Row20.Cells.Add(drowOfObjectData["Kommentar"].ToString(), DataType.String, "m19482158");
                            //cell.StyleID = "s91";

                            cell = Row20.Cells.Add();
                            cell.StyleID = "s91";
                            cell.Data.Type = DataType.String;
                            cell = Row20.Cells.Add();
                            cell.StyleID = "s91";
                            cell.Data.Type = DataType.String;
                            cell = Row20.Cells.Add();
                            cell.StyleID = "s91";
                            cell.Data.Type = DataType.String;
                            //cell = Row20.Cells.Add();
                            //cell.StyleID = "s75";   
                        }
  
                    }
                }
            }

            WorksheetRow Row23 = sheet.Table.Rows.Add();
            Row23.AutoFitHeight = false;
            cell = Row23.Cells.Add();
            cell.StyleID = "s92";
            cell = Row23.Cells.Add();
            cell.StyleID = "s93";
            cell = Row23.Cells.Add();
            cell.StyleID = "s94";
            cell.MergeAcross = 1;
            cell = Row23.Cells.Add();
            cell.StyleID = "m19481956";
            cell.MergeAcross = 1;
            cell = Row23.Cells.Add();
            cell.StyleID = "s102";
            cell = Row23.Cells.Add();
            cell.StyleID = "s103";
            cell = Row23.Cells.Add();
            cell.StyleID = "s104";
            cell = Row23.Cells.Add();
            cell.StyleID = "s105";
            cell = Row23.Cells.Add();
            cell.StyleID = "s106";
            cell = Row23.Cells.Add();
            cell.StyleID = "s106";
            cell = Row23.Cells.Add();
            cell.StyleID = "s106";
            //cell = Row23.Cells.Add();
            //cell.StyleID = "s75";
            // -----------------------------------------------
            WorksheetRow Row24 = sheet.Table.Rows.Add();

            string sumThisNumbers = "-" + numberOfObjects.ToString();

            Row24.AutoFitHeight = false;
            Row24.Cells.Add("Summa", DataType.String, "s107");
            cell = Row24.Cells.Add();
            cell.StyleID = "s24";
            cell = Row24.Cells.Add();
            cell.StyleID = "s24";
            cell = Row24.Cells.Add();
            cell.StyleID = "s21";
            cell = Row24.Cells.Add();
            cell.StyleID = "s24";
            cell = Row24.Cells.Add();
            cell.StyleID = "s108";
            cell = Row24.Cells.Add();
            cell.StyleID = "s21";
            cell = Row24.Cells.Add();
            cell.StyleID = "s109";
            cell.Data.Type = DataType.Number;
            cell.Data.Text = "11.8";
            cell.Formula = "=SUM(R[" + sumThisNumbers + "]C:R[-1]C)";
            cell = Row24.Cells.Add();
            cell.StyleID = "s109";
            cell.Data.Type = DataType.Number;
            cell.Data.Text = "6.5";
            cell.Formula = "=SUM(R[" + sumThisNumbers + "]C:R[-2]C)";
            
            /* added */
            cell = Row24.Cells.Add();
            cell.StyleID = "s21";

            cell = Row24.Cells.Add();
            cell.StyleID = "s110";
            cell.Data.Type = DataType.Number;
            cell.Data.Text = "0";
            cell.Formula = "=SUM(R[" + sumThisNumbers + "]C:R[-1]C)";
            cell = Row24.Cells.Add();
            cell.StyleID = "s21";
            cell = Row24.Cells.Add();
            cell.StyleID = "s75";
            cell.Index = 13;
            // -----------------------------------------------
            // -----------------------------------------------
            WorksheetRow Row25 = sheet.Table.Rows.Add();
            Row25.AutoFitHeight = false;
            cell = Row25.Cells.Add();
            cell.StyleID = "s76";
            cell = Row25.Cells.Add();
            cell.StyleID = "s111";
            cell = Row25.Cells.Add();
            cell.StyleID = "s112";
            cell = Row25.Cells.Add();
            cell.StyleID = "s114";
            cell.MergeAcross = 2;
            cell = Row25.Cells.Add();
            cell.StyleID = "s116";
            cell = Row25.Cells.Add();
            cell.StyleID = "s117";
            cell = Row25.Cells.Add();
            cell.StyleID = "s118";
            cell = Row25.Cells.Add();
            cell.StyleID = "s118";
            cell = Row25.Cells.Add();
            cell.StyleID = "s118";
            cell = Row25.Cells.Add();
            cell.StyleID = "s118";
            cell = Row25.Cells.Add();
            cell.StyleID = "s119";
            // -----------------------------------------------
            WorksheetRow Row26 = sheet.Table.Rows.Add();
            Row26.AutoFitHeight = false;
            cell = Row26.Cells.Add();
            cell.StyleID = "s120";
            cell = Row26.Cells.Add();
            cell.StyleID = "s24";
            cell = Row26.Cells.Add();
            cell.StyleID = "s24";
            cell = Row26.Cells.Add();
            cell.StyleID = "s21";
            cell = Row26.Cells.Add();
            cell.StyleID = "s24";
            cell = Row26.Cells.Add();
            cell.StyleID = "s108";
            cell = Row26.Cells.Add();
            cell.StyleID = "s21";
            cell = Row26.Cells.Add();
            cell.StyleID = "s121";
            cell = Row26.Cells.Add();
            cell.StyleID = "s121";
            cell = Row26.Cells.Add();
            cell.StyleID = "s122";
            cell = Row26.Cells.Add();
            cell.StyleID = "s21";
            cell = Row26.Cells.Add();
            cell.StyleID = "s123";
            cell.Index = 13;
            // -----------------------------------------------
            WorksheetRow Row27 = sheet.Table.Rows.Add();
            Row27.AutoFitHeight = false;
            cell = Row27.Cells.Add();
            cell.StyleID = "s125";
            cell.Data.Type = DataType.String;
            cell.Data.Text = "Startplatser";
            cell.MergeAcross = 6;
            Row27.Cells.Add("Netto-", DataType.String, "s63");
            Row27.Cells.Add("Skog-", DataType.String, "s63");
            cell = Row27.Cells.Add();
            cell.StyleID = "s130";
            Row27.Cells.Add("Transportörens noteringar", DataType.String, "s212");
            //cell = Row27.Cells.Add();
            //cell.StyleID = "s31";
            cell = Row27.Cells.Add();
            cell.StyleID = "s213";
            cell = Row27.Cells.Add();
            cell.StyleID = "s214";
            // -----------------------------------------------
            WorksheetRow Row28 = sheet.Table.Rows.Add();
            Row28.AutoFitHeight = false;
            cell = Row28.Cells.Add();
            cell.StyleID = "s217";
            cell.MergeAcross = 1;
            //cell.StyleID = "s132";
            //cell = Row28.Cells.Add();
            //cell.StyleID = "s21";
            cell = Row28.Cells.Add();
            cell.Data.Type = DataType.String;
            cell.Data.Text = "Koordinat enligt Rikets Nät";
            cell.StyleID = "s217";
            cell.MergeAcross = 4;
            Row28.Cells.Add("areal", DataType.String, "s217");
            Row28.Cells.Add("CAN", DataType.String, "s217");    
            Row28.Cells.Add("", DataType.String, "s70");
            Row28.Cells.Add("Utställt", DataType.String, "s72");
            cell = Row28.Cells.Add();
            cell.StyleID = "s216";
            cell = Row28.Cells.Add();
            cell.StyleID = "s216";
            cell.Index = 13;
            // -----------------------------------------------
            WorksheetRow Row29 = sheet.Table.Rows.Add();
            Row29.AutoFitHeight = false;
            cell = Row29.Cells.Add("Startplats", DataType.String, "s218");
            cell.MergeAcross = 1;
            cell = Row29.Cells.Add();
            cell.Data.Type = DataType.String;
            cell.Data.Text = "Koordinat nordlig";
            cell.StyleID = "s218";
            cell.MergeAcross = 1;
            cell = Row29.Cells.Add();
            cell.Data.Type = DataType.String;
            cell.Data.Text = "Koordinat ostlig";
            cell.StyleID = "s218";
            cell.MergeAcross = 2;
            Row29.Cells.Add("hektar", DataType.String, "s218");
            Row29.Cells.Add("ton", DataType.String, "s218");           
            cell = Row29.Cells.Add();
            cell.StyleID = "s215";
            Row29.Cells.Add("ton", DataType.String, "s72");
            cell = Row29.Cells.Add();
            cell.StyleID = "s216";
            cell = Row29.Cells.Add();
            cell.StyleID = "s216";
            cell.Index = 13;


            WorksheetRow Row30;

            for (int i = 0; i < dStartData.Rows.Count; i++)
            {
                DataRow drowOfStartData = dStartData.Rows[startPlacesSorted[i]];

                // Only row that have not been deleted
                if (drowOfStartData.RowState != DataRowState.Deleted)
                {
                    Row30 = sheet.Table.Rows.Add();
                    Row30.AutoFitHeight = false;
                    cell = Row30.Cells.Add();
                    cell.StyleID = "m19482108";
                    cell.Data.Type = DataType.String;
                    cell.Data.Text = drowOfStartData["Startplats"].ToString();
                    cell.MergeAcross = 1;
                    cell = Row30.Cells.Add();
                    cell.StyleID = "m19482118";
                    cell.Data.Type = DataType.Number;
                    cell.Data.Text = drowOfStartData["Nordligkoordinat_startplats"].ToString();
                    cell.MergeAcross = 1;
                    cell = Row30.Cells.Add();
                    cell.StyleID = "m19482128";
                    cell.Data.Type = DataType.Number;
                    cell.Data.Text = drowOfStartData["Ostligkoordinat_startplats"].ToString();
                    cell.MergeAcross = 2;
                    Row30.Cells.Add(drowOfStartData["Areal_ha_startplats"].ToString().Replace(',', '.').ToString(), DataType.Number, "s149");
                    Row30.Cells.Add(drowOfStartData["Skog_CAN_ton_startplats"].ToString().Replace(',', '.').ToString(), DataType.Number, "s150");
                    cell = Row30.Cells.Add();
// HÄR 5
                    cell.StyleID = "s216";
                    cell = Row30.Cells.Add();
                    cell.StyleID = "s216";
                    cell = Row30.Cells.Add();
                    cell.StyleID = "s216";
                    cell = Row30.Cells.Add();
                    cell.StyleID = "s216";
                    //cell = Row30.Cells.Add();
                    /*cell.StyleID = "s91";
                    cell = Row30.Cells.Add();
                    cell.StyleID = "s151";
                    cell = Row30.Cells.Add();
                    cell.StyleID = "s75";*/
                    cell.Index = 13;       
                     
                }
            }

            // -----------------------------------------------
            WorksheetRow Row32 = sheet.Table.Rows.Add();
            Row32.AutoFitHeight = false;
            cell = Row32.Cells.Add();
            cell.StyleID = "m19482272";
            cell.MergeAcross = 1;
            cell = Row32.Cells.Add();
            cell.StyleID = "m19482282";
            cell.MergeAcross = 1;
            cell = Row32.Cells.Add();
            cell.StyleID = "m19482292";
            cell.MergeAcross = 2;
            cell = Row32.Cells.Add();
            cell.StyleID = "s169";
            cell = Row32.Cells.Add();
            cell.StyleID = "s170";
            cell = Row32.Cells.Add();
            cell.StyleID = "s216";
            cell = Row32.Cells.Add();
            cell.StyleID = "s216";
            cell = Row32.Cells.Add();
            cell.StyleID = "s216";
            cell = Row32.Cells.Add();
            cell.StyleID = "s216";
            cell.Index = 13;

            // -----------------------------------------------
            WorksheetRow Row33 = sheet.Table.Rows.Add();

            int numberOfRowsForStarData = dStartData.Rows.Count;
            numberOfRowsForStarData++;

            string sumThisStartNumbers = "-" + numberOfRowsForStarData.ToString();

            Row33.AutoFitHeight = false;
            Row33.Cells.Add("Summa", DataType.String, "s171");
            cell = Row33.Cells.Add();
            cell.StyleID = "s24";
            cell = Row33.Cells.Add();
            cell.StyleID = "s24";
            cell = Row33.Cells.Add();
            cell.StyleID = "s21";
            cell = Row33.Cells.Add();
            cell.StyleID = "s24";
            cell = Row33.Cells.Add();
            cell.StyleID = "s108";
            cell = Row33.Cells.Add();
            cell.StyleID = "s24";
            cell = Row33.Cells.Add();
            cell.StyleID = "s109";
            cell.Data.Type = DataType.Number;
            cell.Data.Text = "108.2";
            cell.Formula = "=SUM(R[" + sumThisStartNumbers + "]C:R[-1]C)";
            cell = Row33.Cells.Add();
            cell.StyleID = "s172";
            cell.Data.Type = DataType.Number;
            cell.Data.Text = "60";
            cell.Formula = "=SUM(R[" + sumThisStartNumbers + "]C:R[-1]C)";

            cell = Row33.Cells.Add();
            cell.StyleID = "s151";

            cell = Row33.Cells.Add();
            cell.StyleID = "s173";
            cell.Data.Type = DataType.Number;
            cell.Data.Text = "0";
            cell.Formula = "=SUM(R[" + sumThisStartNumbers + "]C:R[-1]C)";
            cell = Row33.Cells.Add();
            cell.StyleID = "s151";
            cell = Row33.Cells.Add();
            cell.StyleID = "s75";
            cell.Index = 13;

            // -----------------------------------------------
            WorksheetRow Row34 = sheet.Table.Rows.Add();
            Row34.AutoFitHeight = false;
            cell = Row34.Cells.Add();
            cell.StyleID = "s174";
            cell = Row34.Cells.Add();
            cell.StyleID = "s175";
            cell = Row34.Cells.Add();
            cell.StyleID = "s175";
            cell = Row34.Cells.Add();
            cell.StyleID = "s175";
            cell = Row34.Cells.Add();
            cell.StyleID = "s175";
            cell = Row34.Cells.Add();
            cell.StyleID = "s175";
            cell = Row34.Cells.Add();
            cell.StyleID = "s175";
            cell = Row34.Cells.Add();
            cell.StyleID = "s175";
            cell = Row34.Cells.Add();
            cell.StyleID = "s175";
            cell = Row34.Cells.Add();
            cell.StyleID = "s176";
            cell = Row34.Cells.Add();
            cell.StyleID = "s52";
            cell = Row34.Cells.Add();
            cell.StyleID = "s175";
            cell = Row34.Cells.Add();
            cell.StyleID = "s177";

            // -----------------------------------------------
            WorksheetRow Row35 = sheet.Table.Rows.Add();
            Row35.AutoFitHeight = false;
            cell = Row35.Cells.Add();
            cell.StyleID = "s120";
            cell = Row35.Cells.Add();
            cell.StyleID = "s24";
            cell = Row35.Cells.Add();
            cell.StyleID = "s24";
            cell = Row35.Cells.Add();
            cell.StyleID = "s21";
            cell = Row35.Cells.Add();
            cell.StyleID = "s24";
            cell = Row35.Cells.Add();
            cell.StyleID = "s108";
            cell = Row35.Cells.Add();
            cell.StyleID = "s24";
            cell = Row35.Cells.Add();
            cell.StyleID = "s121";
            cell = Row35.Cells.Add();
            cell.StyleID = "s108";
            cell = Row35.Cells.Add();
            cell.StyleID = "s122";
            cell = Row35.Cells.Add();
            cell.StyleID = "s21";

            // -----------------------------------------------
            WorksheetRow Row36 = sheet.Table.Rows.Add();
            Row36.AutoFitHeight = false;
            cell = Row36.Cells.Add("Reservobjekt", DataType.String, "s59");
            cell.MergeAcross = 5;
            cell = Row36.Cells.Add();
            cell.StyleID = "s31";
            Row36.Cells.Add("Netto-", DataType.String, "s63");
            cell = Row36.Cells.Add();
            cell.StyleID = "s31";
            cell.MergeAcross = 1;
            Row36.Cells.Add("Spridn.gruppens noteringar", DataType.String, "s212");
            cell = Row36.Cells.Add();
            cell.StyleID = "s213";
            cell = Row36.Cells.Add();
            cell.StyleID = "s214";

            // -----------------------------------------------
            WorksheetRow Row37 = sheet.Table.Rows.Add();
            Row37.AutoFitHeight = false;
            Row37.Cells.Add("Objekt", DataType.String, "s217");
            cell = Row37.Cells.Add("Avdelning", DataType.String, "s217");
            cell.MergeAcross = 1;
            cell = Row37.Cells.Add("Avdelning", DataType.String, "s217");
            cell.MergeAcross = 2;
            Row37.Cells.Add("Giva ", DataType.String, "s217");
            Row37.Cells.Add("areal", DataType.String, "s217");         
            cell = Row37.Cells.Add("Kommentar", DataType.String, "s217");
            cell.MergeAcross = 1;
            Row37.Cells.Add("Spritt", DataType.String, "s72");
            Row37.Cells.Add("Väg-Obj.", DataType.String, "s73");
            Row37.Cells.Add("Datum", DataType.String, "s74");

            // -----------------------------------------------
            WorksheetRow Row38 = sheet.Table.Rows.Add();
            Row38.AutoFitHeight = false;
            Row38.Cells.Add("Nr", DataType.String, "s218");
            cell = Row38.Cells.Add("Nr", DataType.String, "s218");
            cell.MergeAcross = 1;
            cell = Row38.Cells.Add("Namn", DataType.String, "s218");
            cell.MergeAcross = 2;
            Row38.Cells.Add("KgN/ha", DataType.String, "s218");
            Row38.Cells.Add("hektar", DataType.String, "s218");
            cell = Row38.Cells.Add();
            cell.MergeAcross = 1;
            Row38.Cells.Add("ton", DataType.String, "s78");
            Row38.Cells.Add("m", DataType.String, "s73");
            cell = Row38.Cells.Add();
            cell.StyleID = "s79";
            
            DataSet myReservObjectDataSet = Accessdatabas.LäsIfrånDatabas("Select * from Reservobjekt where Ordernr = " + OrderNR);

             // Get the table from the data set
            DataTable dReservObjectData = myReservObjectDataSet.Tables[0];

            
            WorksheetRow Row39;

            for (int i = 0; i < dReservObjectData.Rows.Count; i++)
            {
                DataRow drowOfReservObjectData = dReservObjectData.Rows[i];

                // Only row that have not been deleted
                if (drowOfReservObjectData.RowState != DataRowState.Deleted)
                {
                    // -----------------------------------------------
                    Row39 = sheet.Table.Rows.Add();
                    Row39.AutoFitHeight = false;
                    Row39.Cells.Add(drowOfReservObjectData["Objektnr"].ToString(), DataType.String, "s92");
                    cell = Row39.Cells.Add(drowOfReservObjectData["Avdnr"].ToString(), DataType.String, "s185");
                    cell.MergeAcross = 1;
                    cell = Row39.Cells.Add();
                    cell.StyleID = "m19482302";
                    cell.Data.Type = DataType.String;
                    cell.Data.Text = drowOfReservObjectData["Avdnamn"].ToString();
                    cell.MergeAcross = 2;
                    Row39.Cells.Add(drowOfReservObjectData["Giva_KgN_ha"].ToString().Replace(',', '.').ToString(), DataType.Number, "s193");
                    Row39.Cells.Add(drowOfReservObjectData["Areal_ha"].ToString().Replace(',', '.').ToString(), DataType.Number, "s149");
                    //Row39.Cells.Add("0", DataType.Number, "s91");

                    /* added */
                    //LSAM added 2009-01-02
                    cell = Row39.Cells.Add(drowOfReservObjectData["Kommentar"].ToString(), DataType.String, "m19482158");
                    cell.MergeAcross = 1;
                    //cell.StyleID = "s91";

                    cell = Row39.Cells.Add();
                    cell.StyleID = "s91";
                    cell = Row39.Cells.Add();
                    cell.StyleID = "s91";
                    cell = Row39.Cells.Add();
                    cell.StyleID = "s91";
                    //cell = Row39.Cells.Add();
                    //cell.StyleID = "s216";

                    //cell.Index = 13;

                }
            }
             

            // -----------------------------------------------
            WorksheetRow Row40 = sheet.Table.Rows.Add();
            Row40.AutoFitHeight = false;
            cell = Row40.Cells.Add();
            cell.StyleID = "s156";
            cell = Row40.Cells.Add();
            cell.StyleID = "s185";
            cell.MergeAcross = 1;
            cell = Row40.Cells.Add();
            cell.StyleID = "m19482424";
            cell.MergeAcross = 2;
            cell = Row40.Cells.Add();
            cell.StyleID = "s194";
            cell = Row40.Cells.Add();
            cell.StyleID = "s169";

            /* added */
            cell = Row40.Cells.Add();
            cell.MergeAcross = 1;
            cell.StyleID = "s91";

            cell = Row40.Cells.Add();
            cell.StyleID = "s91";
            cell.Data.Type = DataType.String;
            cell = Row40.Cells.Add();
            cell.StyleID = "s91";
            cell.Data.Type = DataType.String;
            cell = Row40.Cells.Add();
            cell.StyleID = "s91";
            cell.Data.Type = DataType.String;
            //cell = Row40.Cells.Add();
            //cell.StyleID = "s216";
            cell.Index = 13;
            // -----------------------------------------------
            WorksheetRow Row41 = sheet.Table.Rows.Add();
            Row41.AutoFitHeight = false;
            cell = Row41.Cells.Add();
            cell.StyleID = "s156";
            cell = Row41.Cells.Add();
            cell.StyleID = "s185";
            cell.MergeAcross = 1;
            cell = Row41.Cells.Add();
            cell.StyleID = "m19482434";
            cell.MergeAcross = 2;
            cell = Row41.Cells.Add();
            cell.StyleID = "s194";
            cell = Row41.Cells.Add();
            cell.StyleID = "s169";

            /* added */
            cell = Row41.Cells.Add();
            cell.MergeAcross = 1;
            cell.StyleID = "s91";
            
            cell = Row41.Cells.Add();
            cell.StyleID = "s91";
            cell.Data.Type = DataType.String;
            cell = Row41.Cells.Add();
            cell.StyleID = "s91";
            cell.Data.Type = DataType.String;
            cell = Row41.Cells.Add();
            cell.StyleID = "s91";
            cell.Data.Type = DataType.String;
           // cell = Row41.Cells.Add();
           // cell.StyleID = "s216";
            cell.Index = 13;
            // -----------------------------------------------
            WorksheetRow Row42 = sheet.Table.Rows.Add();

            int numberOfRowsForReservData = dReservObjectData.Rows.Count;
            numberOfRowsForReservData++;
            numberOfRowsForReservData++;
            string sumThisReservNumbers = "-" + numberOfRowsForReservData.ToString();


            Row42.AutoFitHeight = false;
            Row42.Cells.Add("Summa reservareal", DataType.String, "s171");
            cell = Row42.Cells.Add();
            cell.StyleID = "s195";
            cell = Row42.Cells.Add();
            cell.StyleID = "s195";
            cell = Row42.Cells.Add();
            cell.StyleID = "s196";
            cell = Row42.Cells.Add();
            cell.StyleID = "s196";
            cell = Row42.Cells.Add();
            cell.StyleID = "s196";
            cell = Row42.Cells.Add();
            cell.StyleID = "s23";
            cell = Row42.Cells.Add();
            cell.StyleID = "s109";
            cell.Data.Type = DataType.Number;
            cell.Data.Text = "33";
            cell.Formula = "=SUM(R[" + sumThisReservNumbers + "]C:R[-1]C)";

            /* added */
            cell = Row42.Cells.Add();
            cell.StyleID = "s23";

            cell = Row42.Cells.Add();
            cell.StyleID = "s110";

            cell = Row42.Cells.Add();
            cell.StyleID = "s110";

            cell.Data.Type = DataType.Number;
            cell.Data.Text = "0";
            cell.Formula = "=SUM(R[" + sumThisReservNumbers + "]C:R[-1]C)";
            cell = Row42.Cells.Add();
            cell.StyleID = "s197";
            //cell = Row42.Cells.Add();
            //cell.StyleID = "s197";
            cell = Row42.Cells.Add();
            cell.StyleID = "s75";
            cell.Index = 13;
            // -----------------------------------------------
            WorksheetRow Row43 = sheet.Table.Rows.Add();
            Row43.AutoFitHeight = false;
            cell = Row43.Cells.Add();
            cell.StyleID = "s76";
            cell = Row43.Cells.Add();
            cell.StyleID = "s198";
            cell = Row43.Cells.Add();
            cell.StyleID = "s88";
            cell.MergeAcross = 3;
            cell = Row43.Cells.Add();
            cell.StyleID = "s200";
            cell = Row43.Cells.Add();
            cell.StyleID = "s116";
            cell = Row43.Cells.Add();
            cell.StyleID = "s201";
            cell = Row43.Cells.Add();
            cell.StyleID = "s201";
            cell = Row43.Cells.Add();
            cell.StyleID = "s52";
            cell = Row43.Cells.Add();
            cell.StyleID = "s175";
            cell = Row43.Cells.Add();
            cell.StyleID = "s177";
            // -----------------------------------------------
            WorksheetRow Row44 = sheet.Table.Rows.Add();
            Row44.AutoFitHeight = false;
            cell = Row44.Cells.Add();
            cell.StyleID = "s21";
            cell = Row44.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row44.Cells.Add();
            cell.StyleID = "s21";
            cell = Row44.Cells.Add();
            cell.StyleID = "s202";
            cell = Row44.Cells.Add();
            cell.StyleID = "s21";
            cell = Row44.Cells.Add();
            cell.StyleID = "s21";
            cell = Row44.Cells.Add();
            cell.StyleID = "s21";
            cell = Row44.Cells.Add();
            cell.StyleID = "s21";
            cell = Row44.Cells.Add();
            cell.StyleID = "s21";
            // -----------------------------------------------
            WorksheetRow Row45 = sheet.Table.Rows.Add();
            Row45.AutoFitHeight = false;
            Row45.Cells.Add("Kommentar från beställare", DataType.String, "s203");
            cell = Row45.Cells.Add();
            cell.StyleID = "s130";
            cell = Row45.Cells.Add();
            cell.StyleID = "s130";
            cell = Row45.Cells.Add();
            cell.StyleID = "s31";
            cell = Row45.Cells.Add();
            cell.StyleID = "s31";
            cell = Row45.Cells.Add();
            cell.StyleID = "s204";
            cell = Row45.Cells.Add();
            cell.StyleID = "s31";
            cell = Row45.Cells.Add();
            cell.StyleID = "s31";
            cell = Row45.Cells.Add();
            cell.StyleID = "s31";
            cell = Row45.Cells.Add();
            cell.StyleID = "s31";
            cell = Row45.Cells.Add();
            cell.StyleID = "s31";
            cell = Row45.Cells.Add();
            cell.StyleID = "s130";
            cell = Row45.Cells.Add();
            cell.StyleID = "s131";
            // -----------------------------------------------
            WorksheetRow Row46 = sheet.Table.Rows.Add();
            Row46.AutoFitHeight = false;
            Row46.Cells.Add(drowOfCData["Kommentar"].ToString(), DataType.String, "s205");
            cell = Row46.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row46.Cells.Add();
            cell.StyleID = "s21";
            cell = Row46.Cells.Add();
            cell.StyleID = "s202";
            cell = Row46.Cells.Add();
            cell.StyleID = "s21";
            cell = Row46.Cells.Add();
            cell.StyleID = "s21";
            cell = Row46.Cells.Add();
            cell.StyleID = "s21";
            cell = Row46.Cells.Add();
            cell.StyleID = "s21";
            cell = Row46.Cells.Add();
            cell.StyleID = "s21";
            cell = Row46.Cells.Add();
            cell.StyleID = "s75";
            cell.Index = 13;
            // -----------------------------------------------
            WorksheetRow Row47 = sheet.Table.Rows.Add();
            Row47.AutoFitHeight = false;
            cell = Row47.Cells.Add();
            cell.StyleID = "s206";
            cell = Row47.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row47.Cells.Add();
            cell.StyleID = "s21";
            cell = Row47.Cells.Add();
            cell.StyleID = "s202";
            cell = Row47.Cells.Add();
            cell.StyleID = "s21";
            cell = Row47.Cells.Add();
            cell.StyleID = "s21";
            cell = Row47.Cells.Add();
            cell.StyleID = "s21";
            cell = Row47.Cells.Add();
            cell.StyleID = "s21";
            cell = Row47.Cells.Add();
            cell.StyleID = "s21";
            cell = Row47.Cells.Add();
            cell.StyleID = "s75";
            cell.Index = 13;
            // -----------------------------------------------
            WorksheetRow Row48 = sheet.Table.Rows.Add();
            Row48.AutoFitHeight = false;
            cell = Row48.Cells.Add();
            cell.StyleID = "s207";
            cell = Row48.Cells.Add();
            cell.StyleID = "s175";
            cell = Row48.Cells.Add();
            cell.StyleID = "s175";
            cell = Row48.Cells.Add();
            cell.StyleID = "s52";
            cell = Row48.Cells.Add();
            cell.StyleID = "s52";
            cell = Row48.Cells.Add();
            cell.StyleID = "s208";
            cell = Row48.Cells.Add();
            cell.StyleID = "s52";
            cell = Row48.Cells.Add();
            cell.StyleID = "s52";
            cell = Row48.Cells.Add();
            cell.StyleID = "s52";
            cell = Row48.Cells.Add();
            cell.StyleID = "s52";
            cell = Row48.Cells.Add();
            cell.StyleID = "s52";
            cell = Row48.Cells.Add();
            cell.StyleID = "s175";
            cell = Row48.Cells.Add();
            cell.StyleID = "s177";
            // -----------------------------------------------
            WorksheetRow Row49 = sheet.Table.Rows.Add();
            Row49.AutoFitHeight = false;
            cell = Row49.Cells.Add();
            cell.StyleID = "s21";
            cell = Row49.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row49.Cells.Add();
            cell.StyleID = "s21";
            cell = Row49.Cells.Add();
            cell.StyleID = "s202";
            cell = Row49.Cells.Add();
            cell.StyleID = "s21";
            cell = Row49.Cells.Add();
            cell.StyleID = "s21";
            cell = Row49.Cells.Add();
            cell.StyleID = "s21";
            cell = Row49.Cells.Add();
            cell.StyleID = "s21";
            cell = Row49.Cells.Add();
            cell.StyleID = "s21";
            // -----------------------------------------------
            WorksheetRow Row50 = sheet.Table.Rows.Add();
            Row50.AutoFitHeight = false;
            Row50.Cells.Add("Kommentar från spridningsgrupp", DataType.String, "s23");
            cell = Row50.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row50.Cells.Add();
            cell.StyleID = "s21";
            cell = Row50.Cells.Add();
            cell.StyleID = "s202";
            cell = Row50.Cells.Add();
            cell.StyleID = "s21";
            cell = Row50.Cells.Add();
            cell.StyleID = "s21";
            cell = Row50.Cells.Add();
            cell.StyleID = "s21";
            cell = Row50.Cells.Add();
            cell.StyleID = "s21";
            cell = Row50.Cells.Add();
            cell.StyleID = "s21";
            // -----------------------------------------------
            WorksheetRow Row51 = sheet.Table.Rows.Add();
            Row51.AutoFitHeight = false;
            cell = Row51.Cells.Add();
            cell.StyleID = "s209";
            cell = Row51.Cells.Add();
            cell.StyleID = "s210";
            cell = Row51.Cells.Add();
            cell.StyleID = "s210";
            cell = Row51.Cells.Add();
            cell.StyleID = "s209";
            cell = Row51.Cells.Add();
            cell.StyleID = "s209";
            cell = Row51.Cells.Add();
            cell.StyleID = "s211";
            cell = Row51.Cells.Add();
            cell.StyleID = "s210";
            cell = Row51.Cells.Add();
            cell.StyleID = "s210";
            cell = Row51.Cells.Add();
            cell.StyleID = "s210";
            cell = Row51.Cells.Add();
            cell.StyleID = "s210";
            cell = Row51.Cells.Add();
            cell.StyleID = "s210";
            cell = Row51.Cells.Add();
            cell.StyleID = "s210";
            cell = Row51.Cells.Add();
            cell.StyleID = "s210";
            // -----------------------------------------------
            WorksheetRow Row52 = sheet.Table.Rows.Add();
            Row52.AutoFitHeight = false;
            cell = Row52.Cells.Add();
            cell.StyleID = "s209";
            cell = Row52.Cells.Add();
            cell.StyleID = "s210";
            cell = Row52.Cells.Add();
            cell.StyleID = "s210";
            cell = Row52.Cells.Add();
            cell.StyleID = "s209";
            cell = Row52.Cells.Add();
            cell.StyleID = "s209";
            cell = Row52.Cells.Add();
            cell.StyleID = "s211";
            cell = Row52.Cells.Add();
            cell.StyleID = "s210";
            cell = Row52.Cells.Add();
            cell.StyleID = "s210";
            cell = Row52.Cells.Add();
            cell.StyleID = "s210";
            cell = Row52.Cells.Add();
            cell.StyleID = "s210";
            cell = Row52.Cells.Add();
            cell.StyleID = "s210";
            cell = Row52.Cells.Add();
            cell.StyleID = "s210";
            cell = Row52.Cells.Add();
            cell.StyleID = "s210";
            // -----------------------------------------------
            WorksheetRow Row53 = sheet.Table.Rows.Add();
            Row53.AutoFitHeight = false;
            cell = Row53.Cells.Add();
            cell.StyleID = "s209";
            cell = Row53.Cells.Add();
            cell.StyleID = "s210";
            cell = Row53.Cells.Add();
            cell.StyleID = "s210";
            cell = Row53.Cells.Add();
            cell.StyleID = "s209";
            cell = Row53.Cells.Add();
            cell.StyleID = "s209";
            cell = Row53.Cells.Add();
            cell.StyleID = "s211";
            cell = Row53.Cells.Add();
            cell.StyleID = "s210";
            cell = Row53.Cells.Add();
            cell.StyleID = "s210";
            cell = Row53.Cells.Add();
            cell.StyleID = "s210";
            cell = Row53.Cells.Add();
            cell.StyleID = "s210";
            cell = Row53.Cells.Add();
            cell.StyleID = "s210";
            cell = Row53.Cells.Add();
            cell.StyleID = "s210";
            cell = Row53.Cells.Add();
            cell.StyleID = "s210";
            // -----------------------------------------------
            WorksheetRow Row54 = sheet.Table.Rows.Add();
            Row54.AutoFitHeight = false;
            cell = Row54.Cells.Add();
            cell.StyleID = "s209";
            cell = Row54.Cells.Add();
            cell.StyleID = "s210";
            cell = Row54.Cells.Add();
            cell.StyleID = "s210";
            cell = Row54.Cells.Add();
            cell.StyleID = "s209";
            cell = Row54.Cells.Add();
            cell.StyleID = "s209";
            cell = Row54.Cells.Add();
            cell.StyleID = "s211";
            cell = Row54.Cells.Add();
            cell.StyleID = "s210";
            cell = Row54.Cells.Add();
            cell.StyleID = "s210";
            cell = Row54.Cells.Add();
            cell.StyleID = "s210";
            cell = Row54.Cells.Add();
            cell.StyleID = "s210";
            cell = Row54.Cells.Add();
            cell.StyleID = "s210";
            cell = Row54.Cells.Add();
            cell.StyleID = "s210";
            cell = Row54.Cells.Add();
            cell.StyleID = "s210";
            // -----------------------------------------------
            WorksheetRow Row55 = sheet.Table.Rows.Add();
            Row55.AutoFitHeight = false;
            cell = Row55.Cells.Add();
            cell.StyleID = "s209";
            cell = Row55.Cells.Add();
            cell.StyleID = "s210";
            cell = Row55.Cells.Add();
            cell.StyleID = "s210";
            cell = Row55.Cells.Add();
            cell.StyleID = "s209";
            cell = Row55.Cells.Add();
            cell.StyleID = "s209";
            cell = Row55.Cells.Add();
            cell.StyleID = "s211";
            cell = Row55.Cells.Add();
            cell.StyleID = "s210";
            cell = Row55.Cells.Add();
            cell.StyleID = "s210";
            cell = Row55.Cells.Add();
            cell.StyleID = "s210";
            cell = Row55.Cells.Add();
            cell.StyleID = "s210";
            cell = Row55.Cells.Add();
            cell.StyleID = "s210";
            cell = Row55.Cells.Add();
            cell.StyleID = "s210";
            cell = Row55.Cells.Add();
            cell.StyleID = "s210";
            // -----------------------------------------------
            WorksheetRow Row56 = sheet.Table.Rows.Add();
            Row56.AutoFitHeight = false;
            cell = Row56.Cells.Add();
            cell.StyleID = "s209";
            cell = Row56.Cells.Add();
            cell.StyleID = "s210";
            cell = Row56.Cells.Add();
            cell.StyleID = "s210";
            cell = Row56.Cells.Add();
            cell.StyleID = "s209";
            cell = Row56.Cells.Add();
            cell.StyleID = "s209";
            cell = Row56.Cells.Add();
            cell.StyleID = "s211";
            cell = Row56.Cells.Add();
            cell.StyleID = "s210";
            cell = Row56.Cells.Add();
            cell.StyleID = "s210";
            cell = Row56.Cells.Add();
            cell.StyleID = "s210";
            cell = Row56.Cells.Add();
            cell.StyleID = "s210";
            cell = Row56.Cells.Add();
            cell.StyleID = "s210";
            cell = Row56.Cells.Add();
            cell.StyleID = "s210";
            cell = Row56.Cells.Add();
            cell.StyleID = "s210";
            // -----------------------------------------------
            WorksheetRow Row57 = sheet.Table.Rows.Add();
            Row57.AutoFitHeight = false;
            cell = Row57.Cells.Add();
            cell.StyleID = "s209";
            cell = Row57.Cells.Add();
            cell.StyleID = "s210";
            cell = Row57.Cells.Add();
            cell.StyleID = "s210";
            cell = Row57.Cells.Add();
            cell.StyleID = "s209";
            cell = Row57.Cells.Add();
            cell.StyleID = "s209";
            cell = Row57.Cells.Add();
            cell.StyleID = "s211";
            cell = Row57.Cells.Add();
            cell.StyleID = "s210";
            cell = Row57.Cells.Add();
            cell.StyleID = "s210";
            cell = Row57.Cells.Add();
            cell.StyleID = "s210";
            cell = Row57.Cells.Add();
            cell.StyleID = "s210";
            cell = Row57.Cells.Add();
            cell.StyleID = "s210";
            cell = Row57.Cells.Add();
            cell.StyleID = "s210";
            cell = Row57.Cells.Add();
            cell.StyleID = "s210";
            // -----------------------------------------------
            WorksheetRow Row58 = sheet.Table.Rows.Add();
            Row58.AutoFitHeight = false;
            cell = Row58.Cells.Add();
            cell.StyleID = "s209";
            cell = Row58.Cells.Add();
            cell.StyleID = "s210";
            cell = Row58.Cells.Add();
            cell.StyleID = "s210";
            cell = Row58.Cells.Add();
            cell.StyleID = "s209";
            cell = Row58.Cells.Add();
            cell.StyleID = "s209";
            cell = Row58.Cells.Add();
            cell.StyleID = "s211";
            cell = Row58.Cells.Add();
            cell.StyleID = "s210";
            cell = Row58.Cells.Add();
            cell.StyleID = "s210";
            cell = Row58.Cells.Add();
            cell.StyleID = "s210";
            cell = Row58.Cells.Add();
            cell.StyleID = "s210";
            cell = Row58.Cells.Add();
            cell.StyleID = "s210";
            cell = Row58.Cells.Add();
            cell.StyleID = "s210";
            cell = Row58.Cells.Add();
            cell.StyleID = "s210";
            // -----------------------------------------------
            WorksheetRow Row59 = sheet.Table.Rows.Add();
            Row59.AutoFitHeight = false;
            cell = Row59.Cells.Add();
            cell.StyleID = "s21";
            cell = Row59.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row59.Cells.Add();
            cell.StyleID = "s21";
            cell = Row59.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row60 = sheet.Table.Rows.Add();
            Row60.AutoFitHeight = false;
            cell = Row60.Cells.Add();
            cell.StyleID = "s21";
            cell = Row60.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row60.Cells.Add();
            cell.StyleID = "s21";
            cell = Row60.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row61 = sheet.Table.Rows.Add();
            Row61.AutoFitHeight = false;
            cell = Row61.Cells.Add();
            cell.StyleID = "s21";
            cell = Row61.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row61.Cells.Add();
            cell.StyleID = "s21";
            cell = Row61.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row62 = sheet.Table.Rows.Add();
            Row62.AutoFitHeight = false;
            cell = Row62.Cells.Add();
            cell.StyleID = "s21";
            cell = Row62.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row62.Cells.Add();
            cell.StyleID = "s21";
            cell = Row62.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row63 = sheet.Table.Rows.Add();
            Row63.AutoFitHeight = false;
            cell = Row63.Cells.Add();
            cell.StyleID = "s21";
            cell = Row63.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row63.Cells.Add();
            cell.StyleID = "s21";
            cell = Row63.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row64 = sheet.Table.Rows.Add();
            Row64.AutoFitHeight = false;
            cell = Row64.Cells.Add();
            cell.StyleID = "s21";
            cell = Row64.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row64.Cells.Add();
            cell.StyleID = "s21";
            cell = Row64.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row65 = sheet.Table.Rows.Add();
            Row65.AutoFitHeight = false;
            cell = Row65.Cells.Add();
            cell.StyleID = "s21";
            cell = Row65.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row65.Cells.Add();
            cell.StyleID = "s21";
            cell = Row65.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row66 = sheet.Table.Rows.Add();
            Row66.AutoFitHeight = false;
            cell = Row66.Cells.Add();
            cell.StyleID = "s21";
            cell = Row66.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row66.Cells.Add();
            cell.StyleID = "s21";
            cell = Row66.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row67 = sheet.Table.Rows.Add();
            Row67.AutoFitHeight = false;
            cell = Row67.Cells.Add();
            cell.StyleID = "s21";
            cell = Row67.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row67.Cells.Add();
            cell.StyleID = "s21";
            cell = Row67.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row68 = sheet.Table.Rows.Add();
            Row68.AutoFitHeight = false;
            cell = Row68.Cells.Add();
            cell.StyleID = "s21";
            cell = Row68.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row68.Cells.Add();
            cell.StyleID = "s21";
            cell = Row68.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row69 = sheet.Table.Rows.Add();
            Row69.AutoFitHeight = false;
            cell = Row69.Cells.Add();
            cell.StyleID = "s21";
            cell = Row69.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row69.Cells.Add();
            cell.StyleID = "s21";
            cell = Row69.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row70 = sheet.Table.Rows.Add();
            Row70.AutoFitHeight = false;
            cell = Row70.Cells.Add();
            cell.StyleID = "s21";
            cell = Row70.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row70.Cells.Add();
            cell.StyleID = "s21";
            cell = Row70.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row71 = sheet.Table.Rows.Add();
            Row71.AutoFitHeight = false;
            cell = Row71.Cells.Add();
            cell.StyleID = "s21";
            cell = Row71.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row71.Cells.Add();
            cell.StyleID = "s21";
            cell = Row71.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row72 = sheet.Table.Rows.Add();
            Row72.AutoFitHeight = false;
            cell = Row72.Cells.Add();
            cell.StyleID = "s21";
            cell = Row72.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row72.Cells.Add();
            cell.StyleID = "s21";
            cell = Row72.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row73 = sheet.Table.Rows.Add();
            Row73.AutoFitHeight = false;
            cell = Row73.Cells.Add();
            cell.StyleID = "s21";
            cell = Row73.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row73.Cells.Add();
            cell.StyleID = "s21";
            cell = Row73.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row74 = sheet.Table.Rows.Add();
            Row74.AutoFitHeight = false;
            cell = Row74.Cells.Add();
            cell.StyleID = "s21";
            cell = Row74.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row74.Cells.Add();
            cell.StyleID = "s21";
            cell = Row74.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row75 = sheet.Table.Rows.Add();
            Row75.AutoFitHeight = false;
            cell = Row75.Cells.Add();
            cell.StyleID = "s21";
            cell = Row75.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row75.Cells.Add();
            cell.StyleID = "s21";
            cell = Row75.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row76 = sheet.Table.Rows.Add();
            Row76.AutoFitHeight = false;
            cell = Row76.Cells.Add();
            cell.StyleID = "s21";
            cell = Row76.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row76.Cells.Add();
            cell.StyleID = "s21";
            cell = Row76.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row77 = sheet.Table.Rows.Add();
            Row77.AutoFitHeight = false;
            cell = Row77.Cells.Add();
            cell.StyleID = "s21";
            cell = Row77.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row77.Cells.Add();
            cell.StyleID = "s21";
            cell = Row77.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row78 = sheet.Table.Rows.Add();
            Row78.AutoFitHeight = false;
            cell = Row78.Cells.Add();
            cell.StyleID = "s21";
            cell = Row78.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row78.Cells.Add();
            cell.StyleID = "s21";
            cell = Row78.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row79 = sheet.Table.Rows.Add();
            Row79.AutoFitHeight = false;
            cell = Row79.Cells.Add();
            cell.StyleID = "s21";
            cell = Row79.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row79.Cells.Add();
            cell.StyleID = "s21";
            cell = Row79.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row80 = sheet.Table.Rows.Add();
            Row80.AutoFitHeight = false;
            cell = Row80.Cells.Add();
            cell.StyleID = "s21";
            cell = Row80.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row80.Cells.Add();
            cell.StyleID = "s21";
            cell = Row80.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row81 = sheet.Table.Rows.Add();
            Row81.AutoFitHeight = false;
            cell = Row81.Cells.Add();
            cell.StyleID = "s21";
            cell = Row81.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row81.Cells.Add();
            cell.StyleID = "s21";
            cell = Row81.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row82 = sheet.Table.Rows.Add();
            Row82.AutoFitHeight = false;
            cell = Row82.Cells.Add();
            cell.StyleID = "s21";
            cell = Row82.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row82.Cells.Add();
            cell.StyleID = "s21";
            cell = Row82.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row83 = sheet.Table.Rows.Add();
            Row83.AutoFitHeight = false;
            cell = Row83.Cells.Add();
            cell.StyleID = "s21";
            cell = Row83.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row83.Cells.Add();
            cell.StyleID = "s21";
            cell = Row83.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row84 = sheet.Table.Rows.Add();
            Row84.AutoFitHeight = false;
            cell = Row84.Cells.Add();
            cell.StyleID = "s21";
            cell = Row84.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row84.Cells.Add();
            cell.StyleID = "s21";
            cell = Row84.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row85 = sheet.Table.Rows.Add();
            Row85.AutoFitHeight = false;
            cell = Row85.Cells.Add();
            cell.StyleID = "s21";
            cell = Row85.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row85.Cells.Add();
            cell.StyleID = "s21";
            cell = Row85.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row86 = sheet.Table.Rows.Add();
            Row86.AutoFitHeight = false;
            cell = Row86.Cells.Add();
            cell.StyleID = "s21";
            cell = Row86.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row86.Cells.Add();
            cell.StyleID = "s21";
            cell = Row86.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row87 = sheet.Table.Rows.Add();
            Row87.AutoFitHeight = false;
            cell = Row87.Cells.Add();
            cell.StyleID = "s21";
            cell = Row87.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row87.Cells.Add();
            cell.StyleID = "s21";
            cell = Row87.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row88 = sheet.Table.Rows.Add();
            Row88.AutoFitHeight = false;
            cell = Row88.Cells.Add();
            cell.StyleID = "s21";
            cell = Row88.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row88.Cells.Add();
            cell.StyleID = "s21";
            cell = Row88.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row89 = sheet.Table.Rows.Add();
            Row89.AutoFitHeight = false;
            cell = Row89.Cells.Add();
            cell.StyleID = "s21";
            cell = Row89.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row89.Cells.Add();
            cell.StyleID = "s21";
            cell = Row89.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row90 = sheet.Table.Rows.Add();
            Row90.AutoFitHeight = false;
            cell = Row90.Cells.Add();
            cell.StyleID = "s21";
            cell = Row90.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row90.Cells.Add();
            cell.StyleID = "s21";
            cell = Row90.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row91 = sheet.Table.Rows.Add();
            Row91.AutoFitHeight = false;
            cell = Row91.Cells.Add();
            cell.StyleID = "s21";
            cell = Row91.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row91.Cells.Add();
            cell.StyleID = "s21";
            cell = Row91.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row92 = sheet.Table.Rows.Add();
            Row92.AutoFitHeight = false;
            cell = Row92.Cells.Add();
            cell.StyleID = "s21";
            cell = Row92.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row92.Cells.Add();
            cell.StyleID = "s21";
            cell = Row92.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row93 = sheet.Table.Rows.Add();
            Row93.AutoFitHeight = false;
            cell = Row93.Cells.Add();
            cell.StyleID = "s21";
            cell = Row93.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row93.Cells.Add();
            cell.StyleID = "s21";
            cell = Row93.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row94 = sheet.Table.Rows.Add();
            Row94.AutoFitHeight = false;
            cell = Row94.Cells.Add();
            cell.StyleID = "s21";
            cell = Row94.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row94.Cells.Add();
            cell.StyleID = "s21";
            cell = Row94.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row95 = sheet.Table.Rows.Add();
            Row95.AutoFitHeight = false;
            cell = Row95.Cells.Add();
            cell.StyleID = "s21";
            cell = Row95.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row95.Cells.Add();
            cell.StyleID = "s21";
            cell = Row95.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row96 = sheet.Table.Rows.Add();
            Row96.AutoFitHeight = false;
            cell = Row96.Cells.Add();
            cell.StyleID = "s21";
            cell = Row96.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row96.Cells.Add();
            cell.StyleID = "s21";
            cell = Row96.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row97 = sheet.Table.Rows.Add();
            Row97.AutoFitHeight = false;
            cell = Row97.Cells.Add();
            cell.StyleID = "s21";
            cell = Row97.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row97.Cells.Add();
            cell.StyleID = "s21";
            cell = Row97.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row98 = sheet.Table.Rows.Add();
            Row98.AutoFitHeight = false;
            cell = Row98.Cells.Add();
            cell.StyleID = "s21";
            cell = Row98.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row98.Cells.Add();
            cell.StyleID = "s21";
            cell = Row98.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row99 = sheet.Table.Rows.Add();
            Row99.AutoFitHeight = false;
            cell = Row99.Cells.Add();
            cell.StyleID = "s21";
            cell = Row99.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row99.Cells.Add();
            cell.StyleID = "s21";
            cell = Row99.Cells.Add();
            cell.StyleID = "s202";
            // -----------------------------------------------
            WorksheetRow Row100 = sheet.Table.Rows.Add();
            Row100.AutoFitHeight = false;
            cell = Row100.Cells.Add();
            cell.StyleID = "s21";
            cell = Row100.Cells.Add();
            cell.StyleID = "s21";
            cell.Index = 4;
            cell = Row100.Cells.Add();
            cell.StyleID = "s21";
            cell = Row100.Cells.Add();
            cell.StyleID = "s202";
  
            // -----------------------------------------------
            //  Options
            // -----------------------------------------------
            sheet.Options.Selected = true;
            sheet.Options.ProtectObjects = false;
            sheet.Options.ProtectScenarios = false;
            sheet.Options.PageSetup.Layout.Orientation = Orientation.Landscape;
            sheet.Options.PageSetup.PageMargins.Bottom = 0.984252F;
            sheet.Options.PageSetup.PageMargins.Left = 0.7874016F;
            sheet.Options.PageSetup.PageMargins.Right = 0.7874016F;
            sheet.Options.PageSetup.PageMargins.Top = 0.984252F;
            sheet.Options.Print.PaperSizeIndex = 9;
            sheet.Options.Print.ValidPrinterInfo = true;

        }

        /// <summary>
        /// Sorterar raderna i datatabellen i bokstavsordning men tar även hänsyn till siffror som siffror. T.ex. A2 sorteras före A10. 
        /// </summary>
        /// <param name="myStartPlaceDataTable"></param>
        /// <returns></returns>
        private int[] SortNumbersAsWell(DataTable startPlaces)
        {
            // Sorteringslistan som håller reda på i vilken ordning startplatserna skall komma. 
            int[] ans = new int[startPlaces.Rows.Count];
            for (int i = 0; i < ans.Length; i++)
                ans[i] = i;

            for (int i = 0; i < startPlaces.Rows.Count; i++)
            {
                for (int j = i + 1; j < startPlaces.Rows.Count; j++)
                {
                    char[] startPlace1 = startPlaces.Rows[ans[i]]["Startplats"].ToString().ToUpper().ToCharArray();
                    char[] startPlace2 = startPlaces.Rows[ans[j]]["Startplats"].ToString().ToUpper().ToCharArray();

                    // Jämför  
                    bool StartPlacesAreTheSame = true;
                    for (int charNo = 0; charNo < Math.Min(startPlace1.Length, startPlace2.Length); charNo++)
                    {
                        if (char.IsDigit(startPlace1[charNo]) && char.IsDigit(startPlace2[charNo]))
                        {
                            int number1 = 0;
                            int number2 = 0;

                            for (int k = charNo; k < startPlace1.Length && char.IsDigit(startPlace1[k]); k++)
                                number1 = number1 * 10 + int.Parse(startPlace1[k].ToString());

                            for (int k = charNo; k < startPlace2.Length && char.IsDigit(startPlace2[k]); k++)
                                number2 = number2 * 10 + int.Parse(startPlace2[k].ToString());

                            if (number1 < number2)
                            {
                                StartPlacesAreTheSame = false;
                                break;
                            }

                            if (number1 > number2)
                            {
                                StartPlacesAreTheSame = false;
                                Swop(i, j, ref ans);
                                break;
                            }

                            if ((int)startPlace1[charNo] < (int)startPlace2[charNo])
                            {
                                StartPlacesAreTheSame = false;
                                break;
                            }

                            if ((int)startPlace1[charNo] > (int)startPlace2[charNo])
                            {

                                StartPlacesAreTheSame = false;
                                Swop(i, j, ref ans);
                                break;

                            }
                        }
                    }

                    if (StartPlacesAreTheSame && startPlace2.Length < startPlace1.Length)
                    {
                        Swop(i, j, ref ans);
                        break;
                    }
                }
            }

            return ans;
        }

        /// <summary>
        /// Byter plats på två datarader i en lista med datarader. 
        /// </summary>
        /// <param name="datarows">Listan med datarader. </param>
        /// <param name="index1">Första indexet som skall byta plats. </param>
        /// <param name="index2">Andra indexet som skall byta plats. </param>
        public void Swop(int index1, int index2, ref int[] ans)
        {
            int tmp = ans[index1];
            ans[index1] = ans[index2];
            ans[index2] = tmp;

            /*
            DataRow row1 = startPlaces.Rows[index1];
            DataRow row2 = startPlaces.Rows[index2];

            DataRow tmpRow1 = startPlaces.NewRow();
            tmpRow1.ItemArray = row1.ItemArray;
            DataRow tmpRow2 = startPlaces.NewRow();
            tmpRow2.ItemArray = row2.ItemArray;

            startPlaces.Rows.RemoveAt(index1);
            startPlaces.Rows.InsertAt(tmpRow1, index2);
            startPlaces.Rows.RemoveAt(index2);
            startPlaces.Rows.InsertAt(tmpRow2, index1);
            */

            /*
            DataRow tmpRow;

            tmpRow = datarows[index1];
            datarows[index1] = datarows[index2];
            datarows[index2] = tmpRow;
             */
        }
    }
}