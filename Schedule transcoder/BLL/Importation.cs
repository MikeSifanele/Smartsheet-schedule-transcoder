using System;
using ExcelDataReader;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace BLL
{
    public class Importation
    {
        DataSet myExcel_Spreadsheet_DataSet;
        IExcelDataReader myExcel_Data_Reader = null;
        public DataSet Load_Excel_Spreadsheet_DataSet(string Excel_Spreadsheet_File_Path)
        {
            myExcel_Spreadsheet_DataSet = null;

            try
            {
                FileStream myFileStream = File.Open(Excel_Spreadsheet_File_Path, FileMode.Open, FileAccess.Read);
                try
                {
                    myExcel_Data_Reader = ExcelReaderFactory.CreateBinaryReader(myFileStream);
                }
                catch
                {
                    myExcel_Data_Reader = ExcelReaderFactory.CreateOpenXmlReader(myFileStream);
                }

                myExcel_Spreadsheet_DataSet = myExcel_Data_Reader.AsDataSet();

                myExcel_Data_Reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occured {ex.Message}","Error");
                Application.Restart();
            }            

            return myExcel_Spreadsheet_DataSet;
        }
    }
}
