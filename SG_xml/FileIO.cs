/****************************************************
 * FileIO.cs Handle IO for XML writer 
 * 
 * 
 * 
 * Fredrik Björklund FEBJ, SWECO Position, 2008
 * 
 *****************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace SG_xml
{
    class FileIO
    {

        /// <summary>
        /// CheckXMLFile: Check if file exist
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public bool CheckXMLFile(string filePath)
        {
            _exitExport = false;

            //Check if the file exist
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
                return true;
            }
            else
                return false;
        
        }

        /// <summary>
        /// CheckFile: Check if file exist
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public bool CheckXLSFile(string filePathTemp, string filePath, string FileName)
        {
            _exitExport = false;

            //Check if the file exist
            if (File.Exists(filePath))
            {
                //Check if question about overwriting has to be asked
                if (_OverwriteAllXLSQuestion)
                {
                    //Set question variable, do not ask this question again
                    _OverwriteAllXLSQuestion = false;

                    //Ask the question and set the variables
                    DialogResult deleteAllFilesDialog = MessageBox.Show("Skriva över alla filer?", "Varning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (deleteAllFilesDialog == DialogResult.Yes)
                    {
                        _OverwriteAllXLS = true;
                        File.Delete(filePath);
                        return true;
                    }
                }
                //If overwrite all, delete all files and write new one
                if (_OverwriteAllXLS)
                {
                    File.Delete(filePath);
                    return true;
                }
                else
                {
                    //The file exist, overwrite it question
                    DialogResult deleteFileDialog = MessageBox.Show("Filen " + FileName + " existerar redan, skriva över filen?", "Varning", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                    if (deleteFileDialog == DialogResult.Yes)
                    {
                        //Owrite one file has been told, ask about overwrite all
                        if (doOneTimeXLS)
                        {
                            doOneTimeXLS = false;
                            _OverwriteAllXLSQuestion = true;
                        }

                        //delete file
                        File.Delete(filePath);
                        return true;
                    }
                    else if (deleteFileDialog == DialogResult.No)
                    {
                        try
                        {
                            File.Delete(filePathTemp);
                            return false;
                        }
                        catch (SystemException ex)
                        {
                        }
                    }
                    else
                    {
                        try
                        {
                            File.Delete(filePathTemp);
                            _exitXLSExport = true;
                        }
                        catch (SystemException ex)
                        {
                            _exitXLSExport = true;
                        }

                        return false;
                    }
                }
            }
            return true;
        }

        public void DeleteThisFile(string file)
        {
            try
            {
                File.Delete(file);
            }
            catch (SystemException ex)
            {
                throw ex;
            }
        }

        private bool _exitExport = false;
        private bool _exitXLSExport = false;

        private bool doOneTime = true;
        private bool doOneTimeXLS = true;

        /// <summary>
        /// ExitExport: Stop create repports 
        /// </summary>
        public bool ExitExport
        {
            get
            {
                return _exitExport;
            }
        }

        /// <summary>
        /// ExitExport: Stop create repports 
        /// </summary>
        public bool ExitXLSExport
        {
            get
            {
                return _exitXLSExport;
            }
        }

        private bool _OverwriteAll = false;
        private bool _OverwriteAllXLS = false;

        /// <summary>
        /// OverWriteAll: Over write all variable. 
        /// </summary>
        public bool OverWriteAll
        {
            get
            {
                return _OverwriteAll;
            }
            set
            {
                _OverwriteAll = value;
            }
        }

        /// <summary>
        /// OverWriteAll: Over write all variable. 
        /// </summary>
        public bool OverWriteAllXLS
        {
            get
            {
                return _OverwriteAllXLS;
            }
            set
            {
                _OverwriteAllXLS = value;
            }
        }

        private bool _OverwriteAllQuestion = false;
        private bool _OverwriteAllXLSQuestion = false;

        /// <summary>
        /// OverwriteAllQuestion: Over write all question variable. 
        /// </summary>
        public bool OverwriteAllQuestion
        {
            get
            {
                return _OverwriteAllQuestion;
            }
            set
            {
                _OverwriteAllQuestion = value;
            }
        }

        /// <summary>
        /// OverwriteAllQuestion: Over write all question variable. 
        /// </summary>
        public bool OverwriteAllXLSQuestion
        {
            get
            {
                return _OverwriteAllXLSQuestion;
            }
            set
            {
                _OverwriteAllXLSQuestion = value;
            }
        }

        /// <summary>
        /// ChechPath: Check if path exists
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool ChechPath(string filePath)
        {

            if (Directory.Exists(filePath))
                return true;
            else
            {
                DialogResult createFilePathDialog = MessageBox.Show("Sökvägen " + filePath + " existerar inte, skapa sökvägen?", "Varning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (createFilePathDialog == DialogResult.Yes)
                {
                    try
                    {
                        Directory.CreateDirectory(filePath);
                        return true;
                    }
                    catch(SystemException ex)
                    {
                        throw ex;
                    }
                }
                else
                    return false;
            }
        }

        /// <summary>
        /// CheckPathAfterCreate: Check the path
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool CheckPathAfterCreate(string filePath)
        {
            return Directory.Exists(filePath);
        }
    }

    

}
