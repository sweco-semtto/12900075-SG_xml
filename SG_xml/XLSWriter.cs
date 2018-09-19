using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using System.Windows.Forms;


namespace SG_xml
{
    class XLSWriter
    {
        public XLSWriter()
        {

        }

        //public void WriteXLSFile(string file)
        public void WriteXLSFile(string path, string pathTemp)
        {
            System.Globalization.CultureInfo myNewCulture = new System.Globalization.CultureInfo("en-US");
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            System.Threading.Thread.CurrentThread.CurrentCulture = myNewCulture;
            Microsoft.Office.Interop.Excel.Workbook book = null;

            try
            {



                XlFileFormat inFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel9795;

                app.Visible = true;
                app.UserControl = true;
                
                
               // System.Globalization.CultureInfo myNewCulture = new System.Globalization.CultureInfo("en-US");

                
                
//                book.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, Missing.Value, book, Missing.Value, myNewCulture, );

                book = app.Workbooks.Open(pathTemp, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                //book = app.Workbooks.Open(pathTemp, null, true, null, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, null, null, null);
                
                
                app.DisplayAlerts = false;
                // sheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Wor ksheets.get_Item[1];

                XlFileFormat format = Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel5;
                //XlSaveAsAccessMode mySave = new XlSaveAsAccessMode();
                XlSaveAsAccessMode acessMode = Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange;

                XlSaveConflictResolution ConflictResolution = Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges;

                //77file = file.Length;
                //ConflictOption 
                book.SaveAs(path, format, Type.Missing, Type.Missing, Type.Missing, Type.Missing, acessMode, ConflictResolution, Type.Missing, Type.Missing, Missing.Value, Missing.Value);


                //Conflci Microsoft.Office.Interop.Excel.confli

                book.Close(Type.Missing, Type.Missing, Type.Missing);
               
                FileIO myFileIO = new FileIO();
                myFileIO.DeleteThisFile(pathTemp);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(book);

                book = null;
                
                app.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;

/*
                book = app.Workbooks.Open(file, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);


                // sheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Wor ksheets.get_Item[1];

                XlFileFormat format = Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel9795;
                //XlSaveAsAccessMode mySave = new XlSaveAsAccessMode();
                XlSaveAsAccessMode acessMode = Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange;

                //77file = file.Length;

                book.SaveAs(file.Remove(file.Length - 4) + ".xls", format, Type.Missing, Type.Missing, Type.Missing, Type.Missing, acessMode, Type.Missing, Type.Missing, Type.Missing, Missing.Value, Missing.Value);

                book.Close(Type.Missing, Type.Missing, Type.Missing);

                FileIO myFileIO = new FileIO();
                myFileIO.DeleteThisFile(file);     
 * */
            
            }
            catch (SystemException ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
                book = null;
                app.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;

                FileIO myFileIO = new FileIO();
                myFileIO.DeleteThisFile(pathTemp);
                //throw ex;
            }
        }
    }


   
}
