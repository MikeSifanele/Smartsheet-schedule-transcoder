using System;
using System.Data;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Drawing;

namespace BLL
{
    public class Exportation
    {
        StreamReader myReader;
        System.Data.DataTable table1 = null;
        System.Data.DataTable table2 = null;
        public bool Export_To_Excel(System.Data.DataTable myTable_From_Excel, string myFilePath, string myFileName, string SA_Region, string A_Region, string TeamName, string ExportingBoth)
        {
            bool results = false;

            using (System.Data.DataTable myTable = myTable_From_Excel)
            {
                System.Data.DataTable myExcel_Worksheet_Input = myTable;
                
                string[] myStart_Date = Get_Start_Date(myTable);
                string[] myIS20_Platform = Get_IS20_Platform();
                string[] myE36B_Platform = Get_E36B_Platform();
                string[] Bouquet = Get_Bouquet();
                string[] myPlatforms = Get_Platforms(myTable, myIS20_Platform, myE36B_Platform, SA_Region, A_Region);
                string[] myBouquets = Get_Bouquets(myTable, Bouquet, SA_Region);

                System.Data.DataTable DataTable_Results = Get_myExcel_Spreadsheet_DataTable_Results(myTable, myPlatforms, myBouquets, myStart_Date, SA_Region, A_Region);

                DataTable_Results = Get_Combined_TRSA_and_TAFR_Regions(DataTable_Results, A_Region);

                DataTable_Results = Get_Table_Without_Duplicates(DataTable_Results);
                
                if (ExportingBoth == "1")
                {
                    table1 = DataTable_Results;

                    table1.Rows.Add("", "", "", "", "", "", DateTime.Now);
                    table1.Rows.Add("M-Net", "", "", "", "", "", DateTime.Now);
                    table1.Rows.Add("", "", "", "", "", "", DateTime.Now);

                    return true;
                }
                else if (ExportingBoth == "2")
                {
                    table2 = DataTable_Results;

                    table1.Merge(table2);

                    table2 = null;

                    Write_To_Excel(table1, myFilePath, TeamName);

                    table1 = null;

                    return true;
                }
                else if (ExportingBoth == "11")
                {
                    
                    table1 = Get_Team_Table(DataTable_Results,"Third Party");
                    
                    return true;
                }
                else if (ExportingBoth == "22")
                {
                
                    table2 = Get_Team_Table(DataTable_Results, "M-Net");
                    
                    table1.Merge(table2);
                    
                    table1.DefaultView.Sort = "StartDate ASC, Title ASC";
                    table1 = table1.DefaultView.ToTable();
                    
                    table2 = null;

                    Write_To_Excel(table1, myFilePath,TeamName);

                    table1 = null;
                    
                    return true;
                }

                results = Write_Results_To_Excel(DataTable_Results, myExcel_Worksheet_Input, myFilePath, myFileName, TeamName, ExportingBoth);

            };

            return results;
        }
        public void Write_To_Excel(System.Data.DataTable myExcel_Table, string myFilePath,string sheetName)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = null;
            Workbook xlWorkbook = null;
            Sheets xlSheets = null;
            Worksheet xlNewSheet = null;

            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();


                // Uncomment the line below if you want to see what's happening in Excel
                xlApp.Visible = true;

                xlWorkbook = xlApp.Workbooks.Open(myFilePath, 0, false, 5, "", "",
                        false, XlPlatform.xlWindows, "",
                        true, false, 0, true, false, false);

                xlSheets = xlWorkbook.Sheets as Sheets;

                // The first argument below inserts the new worksheet as the first one
                xlNewSheet = (Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);

                xlNewSheet.Name = sheetName;
                //xlNewSheet.Range[xlWorkbook.,]

                xlNewSheet.Cells[0 + 1, 0 + 1] = "IS20 Platforms";
                xlNewSheet.Cells[0 + 1, 0 + 2] = "E36B Platforms";
                xlNewSheet.Cells[0 + 1, 0 + 3] = "Title";
                xlNewSheet.Cells[0 + 1, 0 + 4] = "GenrefNo";
                xlNewSheet.Cells[0 + 1, 0 + 5] = "UID";
                xlNewSheet.Cells[0 + 1, 0 + 6] = "Start Date";
                xlNewSheet.Cells[0 + 1, 0 + 7] = "Team";

                

                for (int r = 0; r < myExcel_Table.Rows.Count; r++)
                {

                    for (int c = 0; c < myExcel_Table.Columns.Count; c++)
                    {

                        try
                        {
                            if (myExcel_Table.Rows[r][3].ToString().ToUpper() == "" && myExcel_Table.Rows[r][4].ToString().ToUpper() == "" && c > 2)
                            {
                                xlNewSheet.Cells[r + 2, c + 1] = "";
                            }
                            else
                            {
                                xlNewSheet.Cells[r + 2, c + 1] = myExcel_Table.Rows[r][c].ToString();
                            }

                            
                        }
                        catch 
                        {
                            
                        }
                    }

                    if (myExcel_Table.Rows[r][0].ToString().ToUpper() == "M-NET")
                    {

                        xlNewSheet.Cells[r + 1, 1].EntireRow.Interior.Color = Color.Orange;
                        xlNewSheet.Cells[r + 2, 1].EntireRow.Interior.Color = Color.Orange;
                        xlNewSheet.Cells[r + 3, 1].EntireRow.Interior.Color = Color.Orange;

                    }

                }

                myExcel_Table = null;

                xlWorkbook.Save();
                xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                xlApp.Quit();


            }
            finally
            {
                //relea(xlNewSheet);
                //Marshal.ReleaseComObject(xlSheets);
                //Marshal.ReleaseComObject(xlWorkbook);
                //Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }

        }
        public bool Write_Results_To_Excel(System.Data.DataTable myExcel_Worksheet_Results, System.Data.DataTable myExcel_Worksheet_Input, string myFilePath, string myFileName, string TeamName, string ExportingBoth)
        {
            bool results = false;

            Microsoft.Office.Interop.Excel.Application xlApp = null;
            Workbook xlWorkbook = null;
            Sheets xlSheets = null;
            Worksheet xlNewSheet = null;

            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();


                // Uncomment the line below if you want to see what's happening in Excel
                //xlApp.Visible = true;

                xlWorkbook = xlApp.Workbooks.Open(myFilePath, 0, false, 5, "", "",
                        false, XlPlatform.xlWindows, "",
                        true, false, 0, true, false, false);

                xlSheets = xlWorkbook.Sheets as Sheets;

                // The first argument below inserts the new worksheet as the first one
                xlNewSheet = (Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);

                if (ExportingBoth == "1")
                {
                    xlNewSheet.Name = $"Third Party Results";
                }
                else if (ExportingBoth == "2")
                {
                    xlNewSheet.Name = $"M-Net Results";
                }
                else
                {
                    xlNewSheet.Name = $"{TeamName} Results";
                }


                xlNewSheet.Cells[0 + 1, 0 + 1] = "IS20 Platforms";
                xlNewSheet.Cells[0 + 1, 0 + 2] = "E36B Platforms";
                xlNewSheet.Cells[0 + 1, 0 + 3] = "Title";
                xlNewSheet.Cells[0 + 1, 0 + 4] = "GenrefNo";
                xlNewSheet.Cells[0 + 1, 0 + 5] = "UID";
                xlNewSheet.Cells[0 + 1, 0 + 6] = "Start Date";

                for (int r = 0; r < myExcel_Worksheet_Results.Rows.Count; r++)
                {
                    if (ExportingBoth == "1" && r >= (myExcel_Worksheet_Results.Rows.Count - 3) && myExcel_Worksheet_Results.Rows[r][3].ToString() == "" && myExcel_Worksheet_Results.Rows[r][4].ToString() == "")
                    {
                        break;
                    }
                    for (int c = 0; c < myExcel_Worksheet_Results.Columns.Count; c++)
                    {

                        try
                        {
                            xlNewSheet.Cells[r + 2, c + 1] = myExcel_Worksheet_Results.Rows[r][c].ToString();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"An error occured: {ex.Message}", "Error - Write_Results_To_Excel");
                            System.Windows.Forms.Application.ExitThread();
                        }
                    }

                }


                var saveFileDialoge = new SaveFileDialog();
                saveFileDialoge.FileName = $"{myFileName.Trim(".xlsx".ToCharArray())}-{TeamName.Trim('*')}-Results";
                saveFileDialoge.DefaultExt = ".xlsx";

                if (ExportingBoth == "1")
                {
                    xlWorkbook.Save();
                    xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    xlApp.Quit();
                    results = true;

                }
                else if (saveFileDialoge.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        xlWorkbook.SaveAs(saveFileDialoge.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                        xlApp.Quit();
                        results = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error - Saving Excel spreadsheet");
                        xlApp.Quit();
                        System.Windows.Forms.Application.ExitThread();
                    }
                }

            }
            finally
            {
                //relea(xlNewSheet);
                //Marshal.ReleaseComObject(xlSheets);
                //Marshal.ReleaseComObject(xlWorkbook);
                //Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }

            return results;
        }
        public System.Data.DataTable Get_myExcel_Spreadsheet_DataTable_Results(System.Data.DataTable myExcel_Worksheet, string[] myPlatforms, string[] Bouquets, string[] myStart_Date, string SA_Region, string A_Region)
        {
            System.Data.DataTable myExcel_Spreadsheet_DataTable_Results = new System.Data.DataTable();

            try
            {
                myExcel_Spreadsheet_DataTable_Results.Columns.Add(new DataColumn("Platforms", typeof(string)));  //0
                myExcel_Spreadsheet_DataTable_Results.Columns.Add(new DataColumn("Title", typeof(string)));      //1
                myExcel_Spreadsheet_DataTable_Results.Columns.Add(new DataColumn("GenerefNo", typeof(string)));  //2
                myExcel_Spreadsheet_DataTable_Results.Columns.Add(new DataColumn("UID", typeof(string)));        //3
                myExcel_Spreadsheet_DataTable_Results.Columns.Add(new DataColumn("Start Date", typeof(string))); //4
                myExcel_Spreadsheet_DataTable_Results.Columns.Add(new DataColumn("Region", typeof(string)));     //5
                myExcel_Spreadsheet_DataTable_Results.Columns.Add(new DataColumn("Bouquet", typeof(string)));    //6

                for (int Row_No = 0; Row_No < myExcel_Worksheet.Rows.Count; Row_No++)
                {
                    if (myExcel_Worksheet.Rows[Row_No][4].ToString().ToUpper() == SA_Region || myExcel_Worksheet.Rows[Row_No][4].ToString().ToUpper() == A_Region)
                    {
                        myExcel_Spreadsheet_DataTable_Results.Rows.Add(myPlatforms[Row_No], myExcel_Worksheet.Rows[Row_No][0], myExcel_Worksheet.Rows[Row_No][8], myExcel_Worksheet.Rows[Row_No][12], myStart_Date[Row_No], myExcel_Worksheet.Rows[Row_No][4].ToString(), Bouquets[Row_No]);
                    }
                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error - Getting myExcel Spreadsheet DataTable Results");
                System.Windows.Forms.Application.ExitThread();
            }
            myExcel_Worksheet = null;

            return myExcel_Spreadsheet_DataTable_Results;
        }
        public string[] Get_Platforms(System.Data.DataTable myExcel_Worksheet, string[] myIS20_Platform, string[] myE36B_Platform, string SA_Region, string A_Region)
        {
            string[] myPlatforms = new string[myExcel_Worksheet.Rows.Count];


            string myIS20_Platforms = "";
            string myE36B_Platforms = "";


            //load Platforms.
            try
            {
                //Read rows.
                for (int Row_No = 0; Row_No < myExcel_Worksheet.Rows.Count; Row_No++)
                {
                    if (myExcel_Worksheet.Rows[Row_No][4].ToString() == SA_Region)
                    {
                        //Read columns.
                        for (int i = 0; i < myIS20_Platform.Length; i++)
                        {
                            if (myIS20_Platform[i].Trim(' ') == "Exit")
                            {
                                break;
                            }

                            if (myExcel_Worksheet.Rows[Row_No][Convert.ToInt32(myIS20_Platform[i].Split('|')[0])].ToString().ToUpper() == "YES")
                            {
                                //Load IS20 platform                                        
                                myIS20_Platforms += myIS20_Platform[i].Split('|')[1].Trim(' ') + '/';

                            }
                        }

                        myPlatforms[Row_No] = myIS20_Platforms.TrimEnd('/').Trim(' ') + $"|{myExcel_Worksheet.Rows[Row_No][12].ToString()}";
                    }
                    else if (myExcel_Worksheet.Rows[Row_No][4].ToString() == A_Region)
                    {
                        for (int i = 0; i < myE36B_Platform.Length; i++)
                        {
                            if (myE36B_Platform[i].Trim(' ') == "Exit")
                            {
                                break;
                            }

                            if (myExcel_Worksheet.Rows[Row_No][Convert.ToInt32(myE36B_Platform[i].Split('|')[0])].ToString().ToUpper() == "YES")
                            {
                                //Load E36B platform
                                myE36B_Platforms += myE36B_Platform[i].Split('|')[1].Trim(' ') + '/';

                            }

                        }

                        myPlatforms[Row_No] = myE36B_Platforms.TrimEnd('/').Trim(' ') + $"|{myExcel_Worksheet.Rows[Row_No][12].ToString()}";
                    }

                    //Reset Platforms
                    myIS20_Platforms = "";
                    myE36B_Platforms = "";

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error - Getting Platforms");
                System.Windows.Forms.Application.ExitThread();
            }

            myExcel_Worksheet = null;

            return myPlatforms;
        }
        public string[] Get_IS20_Platform()
        {
            string[] myIS20_Platform = new string[20];
            int myPlatform_No = 0;

            try
            {
                myReader = new StreamReader("Third Party IS20 Platforms.txt");
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show($"{ex.Message}\n\nLocate file?", "Warning", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    //Locate file
                    try
                    {
                        using (OpenFileDialog ImportDialog = new OpenFileDialog() { Filter = "Text Document|*.txt", ValidateNames = true, InitialDirectory = Path.GetDirectoryName(System.Windows.Forms.Application.StartupPath).TrimEnd("\\Schedule transcoder\\bin".ToCharArray()) })
                        {
                            DialogResult myResult = ImportDialog.ShowDialog();

                            if (myResult == DialogResult.OK)
                            {
                                myReader = new StreamReader($"{ImportDialog.FileName}");
                            }
                            else
                            {
                                MessageBox.Show("Text document file not imported.\nPlease try again", "Error");
                            }
                        }
                    }
                    catch
                    {

                    }
                }
                else if (dialogResult == DialogResult.No)
                {
                    System.Windows.Forms.Application.ExitThread();
                }
            }

            while (myReader.EndOfStream == false)
            {
                myIS20_Platform[myPlatform_No] = myReader.ReadLine();

                if (myIS20_Platform[myPlatform_No].Trim(' ') == "")
                {
                    break;
                }
                ++myPlatform_No;
            }

            myReader.Close();

            myIS20_Platform[myPlatform_No] = "Exit";

            return myIS20_Platform;
        }
        public string[] Get_E36B_Platform()
        {
            int myPlatform_No = 0;
            string[] myE36B_Platform = new string[20];

            try
            {
                myReader = new StreamReader("Third Party E36B Platforms.txt");
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show($"{ex.Message}\n\nLocate file?", "Warning", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    //Locate file
                    try
                    {
                        using (OpenFileDialog ImportDialog = new OpenFileDialog() { Filter = "Text Document|*.txt", ValidateNames = true, InitialDirectory = Path.GetDirectoryName(System.Windows.Forms.Application.StartupPath).TrimEnd("\\Schedule transcoder\\bin".ToCharArray()) })
                        {
                            DialogResult myResult = ImportDialog.ShowDialog();

                            if (myResult == DialogResult.OK)
                            {
                                myReader = new StreamReader($"{ImportDialog.FileName}");
                            }
                            else
                            {
                                MessageBox.Show("Text document file not imported.\nPlease try again", "Error");
                            }
                        }
                    }
                    catch
                    {

                    }
                }
                else if (dialogResult == DialogResult.No)
                {
                    System.Windows.Forms.Application.ExitThread();
                }
            }

            while (myReader.EndOfStream == false)
            {
                myE36B_Platform[myPlatform_No] = myReader.ReadLine();

                if (myE36B_Platform[myPlatform_No].Trim(' ') == "")
                {
                    break;
                }
                ++myPlatform_No;
            }

            myReader.Close();

            myE36B_Platform[myPlatform_No] = "Exit";

            return myE36B_Platform;
        }
        public string[] Get_Start_Date(System.Data.DataTable myExcel_Worksheet)
        {
            string[] myStart_Date = new string[myExcel_Worksheet.Rows.Count];

            try
            {
                for (int Row_No = 0; Row_No < myExcel_Worksheet.Rows.Count; Row_No++)
                {
                    if (myExcel_Worksheet.Rows[Row_No][39].ToString().Trim(' ') != "")
                    {
                        myStart_Date[Row_No] = myExcel_Worksheet.Rows[Row_No][39].ToString() + $"|{myExcel_Worksheet.Rows[Row_No][12].ToString()}";
                    }
                    else if (myExcel_Worksheet.Rows[Row_No][86].ToString().Trim(' ') != "")
                    {
                        myStart_Date[Row_No] = myExcel_Worksheet.Rows[Row_No][86].ToString() + $"|{myExcel_Worksheet.Rows[Row_No][12].ToString()}";
                    }
                    else if (myExcel_Worksheet.Rows[Row_No][69].ToString().Trim(' ') != "")
                    {
                        myStart_Date[Row_No] = myExcel_Worksheet.Rows[Row_No][69].ToString() + $"|{myExcel_Worksheet.Rows[Row_No][12].ToString()}";
                    }
                    else if (myExcel_Worksheet.Rows[Row_No][73].ToString().Trim(' ') != "")
                    {
                        myStart_Date[Row_No] = myExcel_Worksheet.Rows[Row_No][73].ToString() + $"|{myExcel_Worksheet.Rows[Row_No][12].ToString()}";
                    }
                    else if (myExcel_Worksheet.Rows[Row_No][35].ToString().Trim(' ') != "")
                    {
                        myStart_Date[Row_No] = myExcel_Worksheet.Rows[Row_No][35].ToString() + $"|{myExcel_Worksheet.Rows[Row_No][12].ToString()}";
                    }
                    else if (myExcel_Worksheet.Rows[Row_No][106].ToString().Trim(' ') != "")
                    {
                        myStart_Date[Row_No] = myExcel_Worksheet.Rows[Row_No][106].ToString() + $"|{myExcel_Worksheet.Rows[Row_No][12].ToString()}";
                    }
                    else
                    {
                        myStart_Date[Row_No] = "1997/10/01 00:00" + $" |{myExcel_Worksheet.Rows[Row_No][12].ToString()}";
                    }                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error - Getting Start Dates");
                System.Windows.Forms.Application.ExitThread();
            }

            myExcel_Worksheet = null;

            return myStart_Date;
        }
        public System.Data.DataTable Get_Combined_TRSA_and_TAFR_Regions(System.Data.DataTable myExcel_Worksheet, string A_Region)
        {
            int Region_Index_to_Combine_with = 0;
            bool Found_Region_to_Combine_with = false;
            System.Data.DataTable Combined_TRSA_and_TAFR_Regions = new System.Data.DataTable(); ;

            Combined_TRSA_and_TAFR_Regions.Columns.Add(new DataColumn("IS20 Platforms", typeof(string)));
            Combined_TRSA_and_TAFR_Regions.Columns.Add(new DataColumn("E36B Platforms", typeof(string)));
            Combined_TRSA_and_TAFR_Regions.Columns.Add(new DataColumn("Other Bouquets", typeof(string)));
            Combined_TRSA_and_TAFR_Regions.Columns.Add(new DataColumn("Title", typeof(string)));
            Combined_TRSA_and_TAFR_Regions.Columns.Add(new DataColumn("GenerefNo", typeof(string)));
            Combined_TRSA_and_TAFR_Regions.Columns.Add(new DataColumn("UID", typeof(string)));
            Combined_TRSA_and_TAFR_Regions.Columns.Add(new DataColumn("StartDate", typeof(DateTime)));

            try
            {
                for (int i = 0; i < myExcel_Worksheet.Rows.Count; i++)
                {
                    try
                    {
                       
                        for (int j = 0; j < myExcel_Worksheet.Rows.Count; j++)
                        {
                            if (myExcel_Worksheet.Rows[i][2].ToString() == myExcel_Worksheet.Rows[j][2].ToString() && myExcel_Worksheet.Rows[i][3].ToString() == myExcel_Worksheet.Rows[j][3].ToString() && j != i)
                            {
                                //A match was found
                                if (myExcel_Worksheet.Rows[i][5].ToString() != myExcel_Worksheet.Rows[j][5].ToString())
                                {
                                    //Played on both Regions

                                    if (myExcel_Worksheet.Rows[j][5].ToString().ToUpper().Trim(' ') == A_Region)
                                    {

                                        if (!Found_Region_to_Combine_with)
                                        {
                                            Region_Index_to_Combine_with = j;

                                            Found_Region_to_Combine_with = true;

                                            break;
                                        }

                                    }
                                }

                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error - Finding matches");
                    }

                    try
                    {

                        if (Found_Region_to_Combine_with)
                        {

                            Combined_TRSA_and_TAFR_Regions.Rows.Add(myExcel_Worksheet.Rows[i][0].ToString().Split('|')[0].Trim('/'), myExcel_Worksheet.Rows[Region_Index_to_Combine_with][0].ToString().Split('|')[0].Trim('/'), $"{myExcel_Worksheet.Rows[i][6].ToString().Split('|')[0].Trim('/')}/{myExcel_Worksheet.Rows[Region_Index_to_Combine_with][6].ToString().Split('|')[0].Trim('/')}".Trim('/'), myExcel_Worksheet.Rows[i][1], myExcel_Worksheet.Rows[i][2], myExcel_Worksheet.Rows[i][3], Convert.ToDateTime(myExcel_Worksheet.Rows[i][4].ToString().Split('|')[0]));
                            myExcel_Worksheet.Rows[Region_Index_to_Combine_with][0] = "IGNORE";
                            Found_Region_to_Combine_with = false;
                        }
                        else if (myExcel_Worksheet.Rows[i][0].ToString() != "IGNORE" && myExcel_Worksheet.Rows[i][1].ToString() != "IGNORE")
                        {
                            if (myExcel_Worksheet.Rows[i][0].ToString().ToUpper().Contains("E36B"))
                            {
                                Combined_TRSA_and_TAFR_Regions.Rows.Add("", myExcel_Worksheet.Rows[i][0].ToString().Split('|')[0].Trim('/'), myExcel_Worksheet.Rows[i][6].ToString().Split('|')[0].Trim('/'), myExcel_Worksheet.Rows[i][1], myExcel_Worksheet.Rows[i][2], myExcel_Worksheet.Rows[i][3], Convert.ToDateTime(myExcel_Worksheet.Rows[i][4].ToString().Split('|')[0]));

                            }
                            else
                            {
                                Combined_TRSA_and_TAFR_Regions.Rows.Add(myExcel_Worksheet.Rows[i][0].ToString().Split('|')[0].Trim('/'), "", myExcel_Worksheet.Rows[i][6].ToString().Split('|')[0].Trim('/'), myExcel_Worksheet.Rows[i][1], myExcel_Worksheet.Rows[i][2], myExcel_Worksheet.Rows[i][3], Convert.ToDateTime(myExcel_Worksheet.Rows[i][4].ToString().Split('|')[0]));

                            }
                        }
                    }
                    catch
                    {

                    }                   
                    
                }

                //100191549
                for (int r = 0; r < Combined_TRSA_and_TAFR_Regions.Rows.Count; r++)
                {

                    if (Combined_TRSA_and_TAFR_Regions.Rows[r][0].ToString().Contains("IGNORE") || Combined_TRSA_and_TAFR_Regions.Rows[r][1].ToString().Contains("IGNORE"))
                    {
                        
                        Combined_TRSA_and_TAFR_Regions.Rows.RemoveAt(r);
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error - Getting Combined TRSA and TAFR Regions");
                System.Windows.Forms.Application.ExitThread();
            }

            try
            {
                Combined_TRSA_and_TAFR_Regions.DefaultView.Sort = "StartDate ASC, Title ASC";
                Combined_TRSA_and_TAFR_Regions = Combined_TRSA_and_TAFR_Regions.DefaultView.ToTable();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error - Sorting table");
            }

            myExcel_Worksheet = null;

            return Combined_TRSA_and_TAFR_Regions;
        }
        public System.Data.DataTable Get_Table_Without_Duplicates(System.Data.DataTable myExcel_Worksheet)
        {
            try
            {
                for (int r = 1; r < myExcel_Worksheet.Rows.Count; r++)
                {
                    //Remove duplicates 

                    if (myExcel_Worksheet.Rows[r][4].ToString() == myExcel_Worksheet.Rows[r - 1][4].ToString() && myExcel_Worksheet.Rows[r][5].ToString() == myExcel_Worksheet.Rows[r - 1][5].ToString())
                    {
                       
                            for (int ix = 0; ix < myExcel_Worksheet.Rows[r][2].ToString().Split('/').Length; ix++)
                            {

                            if (myExcel_Worksheet.Rows[r][2].ToString().Split('/')[ix].Contains("E36B"))
                                {
                                    myExcel_Worksheet.Rows[r][1] = myExcel_Worksheet.Rows[r][1].ToString() + '/' + myExcel_Worksheet.Rows[r][2].ToString().Split('/')[ix];
                                    myExcel_Worksheet.Rows[r][1] = myExcel_Worksheet.Rows[r][1].ToString().Trim('/');
                                }
                                else
                                {
                                    myExcel_Worksheet.Rows[r][0] = myExcel_Worksheet.Rows[r][0].ToString() + '/' + myExcel_Worksheet.Rows[r][2].ToString().Split('/')[ix];
                                    myExcel_Worksheet.Rows[r][0] = myExcel_Worksheet.Rows[r][0].ToString().Trim('/');
                                }

                            }

                        if (myExcel_Worksheet.Rows[r][0].ToString().Trim(' ') == "")
                        {                            
                                myExcel_Worksheet.Rows.RemoveAt(r);

                        }
                        else
                        {
                            myExcel_Worksheet.Rows.RemoveAt(r - 1);

                        }

                    }
                    
                }
                myExcel_Worksheet.Columns.RemoveAt(2);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "Error - Removing duplicates");
                System.Windows.Forms.Application.ExitThread();
            }


            return myExcel_Worksheet;
        }
        public string[] Get_Bouquet()
        {
            string[] Bouquets = new string[20];
            int myBouquet_No = 0;

            try
            {
                myReader = new StreamReader("Bouquets.txt");
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show($"{ex.Message}\n\nLocate file?", "Warning", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    //Locate file
                    try
                    {
                        using (OpenFileDialog ImportDialog = new OpenFileDialog() { Filter = "Text Document|*.txt", ValidateNames = true, InitialDirectory = Path.GetDirectoryName(System.Windows.Forms.Application.StartupPath).TrimEnd("\\Schedule transcoder\\bin".ToCharArray()) })
                        {
                            DialogResult myResult = ImportDialog.ShowDialog();

                            if (myResult == DialogResult.OK)
                            {
                                myReader = new StreamReader($"{ImportDialog.FileName}");
                            }
                            else
                            {
                                MessageBox.Show("Text document file not imported.\nPlease try again", "Error");
                            }
                        }
                    }
                    catch
                    {

                    }
                }
                else if (dialogResult == DialogResult.No)
                {
                    System.Windows.Forms.Application.ExitThread();
                }
            }

            while (myReader.EndOfStream == false)
            {
                Bouquets[myBouquet_No] = myReader.ReadLine();

                if (Bouquets[myBouquet_No].Trim(' ') == "")
                {
                    break;
                }
                ++myBouquet_No;
            }

            myReader.Close();

            Bouquets[myBouquet_No] = "Exit";

            return Bouquets;
        }
        public string[] Get_Bouquets(System.Data.DataTable myExcel_Worksheet, string[] Bouquet, string SA_Region)
        {
            string myBouquet = "";
            string[] Bouquets = new string[myExcel_Worksheet.Rows.Count];
            try
            {
                for (int i = 0; i < myExcel_Worksheet.Rows.Count; i++)
                {
                    for (int c = 0; c < Bouquet.Length; c++)
                    {
                        if (Bouquet[c].Trim(' ') == "Exit")
                        {
                            break;
                        }

                        if (myExcel_Worksheet.Rows[i][Convert.ToInt32(Bouquet[c].Split('|')[0])].ToString().ToUpper() == "YES")
                        {
                            //Load Bouquet          

                            if (myExcel_Worksheet.Rows[i][4].ToString() == SA_Region)
                            {
                                myBouquet += Bouquet[c].Split('|')[1].Split('-')[0].Trim(' ') + '/';
                            }
                            else
                            {
                                myBouquet += Bouquet[c].Split('|')[1].Split('-')[1].Trim(' ') + '/';
                            }


                        }
                    }
                    //Load bouquet                    

                    if ((myExcel_Worksheet.Rows[i][155].ToString().ToUpper().Contains("FREE")))
                    {
                        myBouquet += "FREE";
                    }

                    Bouquets[i] = myBouquet.TrimEnd('/') + $"|{myExcel_Worksheet.Rows[i][12]}";
                    myBouquet = "";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error - Adding Bouquet");
            }

            myExcel_Worksheet = null;

            return Bouquets;
        }
        public System.Data.DataTable Get_Team_Table(System.Data.DataTable myTable, string myTeam)
        {
            System.Data.DataTable Team_Table = new System.Data.DataTable();

            Team_Table.Columns.Add(new DataColumn("IS20 Platforms", typeof(string)));
            Team_Table.Columns.Add(new DataColumn("E36B Platforms", typeof(string)));
            Team_Table.Columns.Add(new DataColumn("Title", typeof(string)));
            Team_Table.Columns.Add(new DataColumn("GenerefNo", typeof(string)));
            Team_Table.Columns.Add(new DataColumn("UID", typeof(string)));
            Team_Table.Columns.Add(new DataColumn("StartDate", typeof(string)));
            Team_Table.Columns.Add(new DataColumn("Team", typeof(string)));

            try
            {
                for (int r = 0; r < myTable.Rows.Count; r++)
                {
                    //"1997/10/01 00:00"
                    if(myTable.Rows[r][5].ToString().Contains("1997/10/01 00:00"))
                    {
                        Team_Table.Rows.Add(myTable.Rows[r][0], myTable.Rows[r][1], myTable.Rows[r][2], myTable.Rows[r][3], myTable.Rows[r][4], "Start date not found.", myTeam);

                    }
                    else
                    {
                        Team_Table.Rows.Add(myTable.Rows[r][0], myTable.Rows[r][1], myTable.Rows[r][2], myTable.Rows[r][3], myTable.Rows[r][4], myTable.Rows[r][5], myTeam);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Getting Team Table");
            }
            
           
            return Team_Table;
        }
    }
}
