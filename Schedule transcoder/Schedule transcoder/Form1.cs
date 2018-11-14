using System;
using System.Threading;
using System.Data;
using System.IO;
using System.Windows.Forms;
using BLL;

namespace Schedule_transcoder
{
    public partial class Form1 : Form
    {
        Initial_Components myApp = null;

        string Excel_Spreadsheet_File_Path = "", myFileName = "";
        DataSet myExcel_Spreadsheet_DataSet;
        DataTable myTable;
        Importation myImport = new Importation();
        Exportation myExport = new Exportation();

        public Form1()
        {
            InitializeComponent();
        }

        private void pbStatus_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //2018-10-17(13:19)            


            lblStatus.Text = "Status.. Import an Excel file.";

            cmbExcelWorksheets.Items.Insert(0, "<Please select a Worksheet>");
            cmbExcelWorksheets.Items.Insert(1, "SSCCatchUpScheduleRpt");
            cmbExcelWorksheets.SelectedIndex = 1;

            cmbTeams.Items.Insert(0, "<Please select your Team>");
            cmbTeams.Items.Insert(1, "M-Net");
            cmbTeams.Items.Insert(2, "Third Party");
            cmbTeams.Items.Insert(3, "Third Party & M-Net");
            cmbTeams.Items.Insert(4, "Third Party & M-Net Merged");

            cmbTeams.SelectedIndex = 4;
        }
        public string Get_Excel_Spreadsheet_File_Path()
        {
            Excel_Spreadsheet_File_Path = "";
            myFileName = "";

            try
            {
                using (OpenFileDialog ImportDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xls*", ValidateNames = true })
                {
                    DialogResult myResult = ImportDialog.ShowDialog();

                    if (myResult == DialogResult.OK)
                    {
                        Excel_Spreadsheet_File_Path = ImportDialog.FileName;
                        myFileName = ImportDialog.SafeFileName;
                    }
                    else
                    {
                        MessageBox.Show("Excel Spreadsheet file not imported.\nPlease try again", "Error");
                    }
                }
            }
            catch (IOException)
            {
                MessageBox.Show("Excel Spreadsheet file not imported.\nPlease close the file you are trying to import and try again", "Error - file open");
                Application.ExitThread();
            }

            return Excel_Spreadsheet_File_Path;
        }
        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                pbStatus.Value = 0;
                pbOverrall.Value = 0;

                pbOverrall.Increment(33);

                //Clear Excel Spreadsheet DataSet
                myExcel_Spreadsheet_DataSet = null;
                pbStatus.Increment(150);

                //Import Excel Spreadsheet file
                myExcel_Spreadsheet_DataSet = myImport.Load_Excel_Spreadsheet_DataSet(Get_Excel_Spreadsheet_File_Path());
                pbStatus.Increment(150);

                cmbExcelWorksheets.Items.Clear();
                //Load Dropdown List with Worksheet titles from the imported Excel Spreadsheet.
                foreach (DataTable myExcelWorksheets in myExcel_Spreadsheet_DataSet.Tables)
                {
                    cmbExcelWorksheets.Items.Add(myExcelWorksheets.TableName);
                }

                pbStatus.Increment(150);

                cmbExcelWorksheets.Items.Insert(0, "<Please select a Worksheet>");
                pbStatus.Increment(150);

                for (int i = 0; i < cmbExcelWorksheets.Items.Count; i++)
                {
                    if (cmbExcelWorksheets.Items[i].ToString() == "SSCCatchUpScheduleRpt")
                    {
                        cmbExcelWorksheets.SelectedIndex = i;
                        break;
                    }
                }

                //Update status
                lblStatus.Text = "File uploaded... select a Worksheet.";
                pbStatus.Increment(150);
                pbStatus.Increment(250);

            }
            catch
            {
                MessageBox.Show("Please select a file to import.", "Import file");
                Application.ExitThread();
            }
        }

        private void cmbActiveSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            pbOverrall.Increment(33);
            lblStatus.Text = "Worksheet selected.. Export file.";
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                pbStatus.Value = 200;

                pbStatus.Step = 20;

                myTable = myExcel_Spreadsheet_DataSet.Tables[cmbExcelWorksheets.SelectedItem.ToString()];

                myTable.Rows.RemoveAt(0);

                lblStatus.Text = "File Exporting.. Please wait.";
                if (cmbTeams.SelectedIndex == 2)
                {
                    if (myExport.Export_To_Excel(myTable, Excel_Spreadsheet_File_Path, myFileName, "TRSA", "TAFR", cmbTeams.SelectedItem.ToString(), ""))
                    {
                        pbStatus.Value = 1000;
                        pbOverrall.Value = 100;
                        lblStatus.Text = "File Export successful.. Complete.";
                    }
                }
                else if (cmbTeams.SelectedIndex == 1)
                {
                    if (myExport.Export_To_Excel(myTable, Excel_Spreadsheet_File_Path, myFileName, "CRSA", "CAFR", cmbTeams.SelectedItem.ToString(), ""))
                    {
                        pbStatus.Value = 1000;
                        pbOverrall.Value = 100;
                        lblStatus.Text = "File Export successful.. Complete.";
                    }
                }
                else if (cmbTeams.SelectedIndex == 3)
                {
                    if (myExport.Export_To_Excel(myTable, Excel_Spreadsheet_File_Path, myFileName, "TRSA", "TAFR", cmbTeams.SelectedItem.ToString(), "1"))
                    {
                        pbStatus.Value = 350;
                    }

                    if (myExport.Export_To_Excel(myTable, Excel_Spreadsheet_File_Path, myFileName, "CRSA", "CAFR", cmbTeams.SelectedItem.ToString(), "2"))
                    {
                        pbStatus.Value = 350;
                    }

                    pbStatus.Value = 1000;
                    pbOverrall.Value = 100;

                    lblStatus.Text = "File Export successful.. Complete.";
                }
                else if (cmbTeams.SelectedIndex == 4)
                {
                    
                    if (myExport.Export_To_Excel(myTable, Excel_Spreadsheet_File_Path, myFileName, "TRSA", "TAFR", cmbTeams.SelectedItem.ToString(), "11"))
                    {
                        pbStatus.Value = 350;
                    }

                    if (myExport.Export_To_Excel(myTable, Excel_Spreadsheet_File_Path, myFileName, "CRSA", "CAFR", cmbTeams.SelectedItem.ToString(), "22"))
                    {
                        pbStatus.Value = 350;
                    }
                                        
                    pbStatus.Value = 1000;
                    pbOverrall.Value = 100;

                    lblStatus.Text = "File Export successful.. Complete. ";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occured.\n{ex.Message}.", "Error - PL (Export)");
                Application.ExitThread();
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            myApp = new Initial_Components();

            myApp.Load();
        }

        private void btnIntegrityTest_Click(object sender, EventArgs e)
        {
            try
            {
                int Matches = 0, duplicates = 0, overrall = 0;


                for (int i = 1; i < myExcel_Spreadsheet_DataSet.Tables["Results Third Party"].Rows.Count; i++)
                {

                    for (int j = 1; j < myExcel_Spreadsheet_DataSet.Tables["Third Party Results"].Rows.Count; j++)
                    {
                        if (myExcel_Spreadsheet_DataSet.Tables["Third Party Results"].Rows[j][2].ToString() == myExcel_Spreadsheet_DataSet.Tables["Results Third Party"].Rows[i][2].ToString() && myExcel_Spreadsheet_DataSet.Tables["Third Party Results"].Rows[j][2].ToString().Trim(' ') != "")
                        {
                            Matches++;
                            duplicates++;
                        }
                    }
                    if (duplicates > 1)
                    {
                        overrall++;
                    }

                    duplicates = 0;
                }

                MessageBox.Show($"Matches found: \t{Matches}\nOut of:\t[{myExcel_Spreadsheet_DataSet.Tables["Results Third Party"].Rows.Count}] Results from the desired output.\nWith [{overrall}] overrall Duplicates.", "Info.");

            }
            catch
            {
                
            }
        }
    }
}
