using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;


namespace Solidworksaddin
{
    class Excel_Functions
    {

        /// <summary>
        /// Write Data in Template
        /// </summary>
        /// <param name="Header"></param>
        /// <param name="Data"></param>
        public static void Excel_BOM(ModelDoc2 swModel, List<BOM.BOM_Part_Informations> Standard_Parts, List<BOM.BOM_Part_Informations> Custom_Parts, int projectnumber) //String[] Header, String[,] Data
        {

            string path_to_temp = "";
            Microsoft.Office.Interop.Excel.Application excel_app = new Microsoft.Office.Interop.Excel.Application();
            Range worksheet_range;
            // Make Excel visible (optional).
            excel_app.Visible = false;
            excel_app.DisplayAlerts = false;

            if (File.Exists(SwAddin.path_to_template))
            {
                path_to_temp = SwAddin.path_to_template;
            }
            else if (File.Exists(SwAddin.path_to_template_desktop))
            {
                path_to_temp = SwAddin.path_to_template_desktop;
            }
            // Open the workbook read-only.

            if (File.Exists(path_to_temp))
            {

                /* for (int i = 0; i < informations.Length; i++)
                {
                    Debug.Print(informations[i]);
                }*/

                try
                {
                    Microsoft.Office.Interop.Excel.Workbooks workbooks = excel_app.Workbooks;

                    Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Open(
                    path_to_temp,
                    Type.Missing, false, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                    // Get the first worksheet.
                    Microsoft.Office.Interop.Excel.Sheets sheets = workbook.Sheets;
                    Microsoft.Office.Interop.Excel.Worksheet sheet_standard = sheets[SwAddin.standard_part_sheetname];
                    Microsoft.Office.Interop.Excel.Worksheet sheet_custom = sheets[SwAddin.custom_part_sheetname];
                    String path = swModel.GetPathName();
                    String[] informations = path.Split('\\');

                    String[] name = informations[informations.Length - 1].Split('.');
                    String[] excel_path = path.Split('.');

                    if (sheet_custom != null)
                    {


                        sheet_custom.Cells[3, 3] = informations[2];
                        sheet_custom.Cells[4, 3] = name[0] + " " + swModel.SummaryInfo[0];
                        sheet_custom.Cells[5, 3] = informations[1];
                        sheet_custom.Cells[6, 3] = DateTime.Now.Date;

                        for (int cus = 0; cus < Custom_Parts.Count; cus++)
                        {
                            sheet_custom.Cells[SwAddin.excel_template_start_row + cus, SwAddin.excel_template_item_number] = Custom_Parts[cus].item_number;
                            sheet_custom.Cells[SwAddin.excel_template_start_row + cus, SwAddin.excel_template_description] = Custom_Parts[cus].description;
                            sheet_custom.Cells[SwAddin.excel_template_start_row + cus, SwAddin.excel_template_quantity] = Custom_Parts[cus].quantity;
                            sheet_custom.Cells[SwAddin.excel_template_start_row + cus, SwAddin.excel_template_part_number] = Custom_Parts[cus].part_number;
                            sheet_custom.Cells[SwAddin.excel_template_start_row + cus, SwAddin.excel_template_order_number] = Custom_Parts[cus].order_number;
                            sheet_custom.Cells[SwAddin.excel_template_start_row + cus, SwAddin.excel_template_manufacturer] = Custom_Parts[cus].manufacturer;
                            sheet_custom.Cells[SwAddin.excel_template_start_row + cus, SwAddin.excel_template_storage_location] = Custom_Parts[cus].storage_location;
                            sheet_custom.Cells[SwAddin.excel_template_start_row + cus, SwAddin.excel_valid_template_order_number] = Custom_Parts[cus].valid_order_number;
                        }
                        if (Custom_Parts.Count != 0)
                        {
                            worksheet_range = sheet_custom.get_Range("A" + SwAddin.excel_template_start_row.ToString(), "H" + (SwAddin.excel_template_start_row + Custom_Parts.Count - 1).ToString());
                            worksheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
                        }
                    }

                    if (sheet_standard != null)
                    {


                        sheet_standard.Cells[3, 3] = informations[2];
                        sheet_standard.Cells[4, 3] = name[0] + " " + swModel.SummaryInfo[0]; ;
                        sheet_standard.Cells[5, 3] = informations[1];
                        sheet_standard.Cells[6, 3] = DateTime.Now.Date;

                        for (int sta = 0; sta < Standard_Parts.Count; sta++)
                        {
                            sheet_standard.Cells[SwAddin.excel_template_start_row + sta, SwAddin.excel_template_item_number] = Standard_Parts[sta].item_number;
                            sheet_standard.Cells[SwAddin.excel_template_start_row + sta, SwAddin.excel_template_description] = Standard_Parts[sta].description;
                            sheet_standard.Cells[SwAddin.excel_template_start_row + sta, SwAddin.excel_template_quantity] = Standard_Parts[sta].quantity;
                            sheet_standard.Cells[SwAddin.excel_template_start_row + sta, SwAddin.excel_template_part_number] = Standard_Parts[sta].part_number;
                            sheet_standard.Cells[SwAddin.excel_template_start_row + sta, SwAddin.excel_template_order_number] = Standard_Parts[sta].order_number;
                            sheet_standard.Cells[SwAddin.excel_template_start_row + sta, SwAddin.excel_template_manufacturer] = Standard_Parts[sta].manufacturer;
                            sheet_standard.Cells[SwAddin.excel_template_start_row + sta, SwAddin.excel_template_storage_location] = Standard_Parts[sta].storage_location;
                            sheet_standard.Cells[SwAddin.excel_template_start_row + sta, SwAddin.excel_valid_template_order_number] = Standard_Parts[sta].valid_order_number;
                        }
                        if (Standard_Parts.Count != 0)
                        {
                            worksheet_range = sheet_standard.get_Range("A" + SwAddin.excel_template_start_row.ToString(), "H" + (SwAddin.excel_template_start_row + Standard_Parts.Count - 1).ToString());
                            worksheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
                        }
                    }



                    workbook.SaveAs(excel_path[0] + "_bom.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                    workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, excel_path[0] + "_bom.pdf");
                    // Close the workbook without saving changes.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet_custom);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet_standard);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
                    workbook.Close(0);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                    excel_app.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_app);

                    foreach (Process process in Process.GetProcessesByName("Excel"))
                    {
                        if (!string.IsNullOrEmpty(process.ProcessName) && process.StartTime.AddSeconds(+10) > DateTime.Now)
                        {
                            process.Kill();
                        }
                    }





                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + ex.StackTrace);
                }
                finally
                {

                }
            }
            else
            {
                MessageBox.Show("No Template found");
            }
        }


        public static void Excel_Search(List<BOM.BOM_Part_Informations> Standard_parts)
        {

            string path_to_db = "";
            Microsoft.Office.Interop.Excel.Application excel_app = new Microsoft.Office.Interop.Excel.Application();


            // Make Excel visible (optional).
            excel_app.Visible = false;


            if (File.Exists(SwAddin.path_to_database))
            {
                path_to_db = SwAddin.path_to_database;
            }
            else if (File.Exists(SwAddin.path_to_database_desktop))
            {
                path_to_db = SwAddin.path_to_database_desktop;
            }

            if (File.Exists(path_to_db))
            {
                // Open the workbook read-only.



                try
                {

                    Microsoft.Office.Interop.Excel.Workbook workbook = excel_app.Workbooks.Open(
                    path_to_db,
                    Type.Missing, true, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);


                    Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Stock"];
                    int NumCols = 10;
                    int start_row = 4;
                    int end_row = 4000;
                    string[] Fields = new string[NumCols];
                    string[,] search_array = new string[end_row, NumCols];
                    string search_value = "";
                    int index = 0;
                    //string[] temp_search = new string[10];
                    Microsoft.Office.Interop.Excel.Range range = sheet.get_Range("A" + start_row.ToString(), "L" + end_row.ToString());
                    object[,] values = (object[,])range.Value2;
                    int NumRow = 1;
                    while (NumRow < values.GetLength(0))
                    {
                        for (int c = 1; c <= NumCols; c++)
                        {
                            Fields[c - 1] = Convert.ToString(values[NumRow, c]);
                            search_array[NumRow - 1, c - 1] = Convert.ToString(values[NumRow, c]);
                        }
                        NumRow++;
                    }

                    for (int i = 0; i < Standard_parts.Count; i++)
                    {
                        for (int j = 0; j < end_row; j++)
                        {

                            if (search_array[j, SwAddin.db_description].Contains(Standard_parts[i].part_number) || search_array[j, SwAddin.db_technical_data].Contains(Standard_parts[i].description))
                            {

                                //  index = j + start_row;
                            }

                        }


                    }

                    // MessageBox.Show(search_array[6, 4]);
                    // Close the workbook without saving changes.
                    workbook.Close(false, Type.Missing, Type.Missing);

                    // Close the Excel server.
                    excel_app.Quit();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + ex.StackTrace);
                }
            }
            else
            {
                MessageBox.Show("No Database found");
            }
        }


    }
}
