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
    public class BOM
    {


        

        /// <summary>
        /// Compare two strings
        /// </summary>
        public class CaseInsensitiveComparer : IEqualityComparer<BOM_Part_Informations>
        {
            public bool Equals(BOM_Part_Informations x, BOM_Part_Informations y)
            {
                return x.manufacturer.Equals(y.manufacturer, StringComparison.OrdinalIgnoreCase);
            }

            public int GetHashCode(BOM_Part_Informations obj)
            {
                return obj.manufacturer.ToLowerInvariant().GetHashCode();
            }
        }

        /// <summary>
        /// The Class for a Part with important informations
        /// </summary>
        public class BOM_Part_Informations
        {
            public string item_number { get; set; }
            public string part_number { get; set; }
            public string description { get; set; }
            public string quantity { get; set; }
            public string storage_location { get; set; }
            public string manufacturer { get; set; }
            public string order_number { get; set; }
            public string valid_order_number { get; set; }
            public bool IsStandard { get; set; }
        }

        public class Websearch
        {
            string _id;
            string _url;
            string _nomatchkeyword;
            

            public Websearch(string id, string url, string nomatchkeyword)
            {
                this._id = id;
                this._url = url;
                this._nomatchkeyword = nomatchkeyword;
                
            }

            public string Id { get { return _id; } }
            public string Url { get { return _url; } }
            public string Nomatchkeyword { get { return _nomatchkeyword; } }
            
        }


        /// <summary>
        /// Write Data in Template
        /// </summary>
        /// <param name="Header"></param>
        /// <param name="Data"></param>
        public static void Excel_BOM(ModelDoc2 swModel, List<BOM_Part_Informations> Standard_Parts, List<BOM_Part_Informations> Custom_Parts, int projectnumber) //String[] Header, String[,] Data
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

        /// <summary>
        /// Get all suppliers from th Standard parts
        /// </summary>
        /// <param name="Standard_Parts"></param>
        /// <param name="companies"></param>
        public static void Get_Companies(List<BOM_Part_Informations> Standard_Parts, List<string> companies)
        {
            try
            {
                var sorted = Standard_Parts.Distinct(new CaseInsensitiveComparer());
                   
                

                if (sorted != null)
                {
                    foreach (var company in sorted)
                    {
                        if (!string.IsNullOrEmpty(company.manufacturer))
                        {
                            companies.Add(company.manufacturer);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Creates Baskets for the supplier webpages
        /// </summary>
        /// <param name="Standard_Parts"></param>
        /// <param name="projectpath"></param>
        /// <param name="projectnumber"></param>
        public static void Create_Project_Basket_by_Company(List<BOM_Part_Informations> Standard_Parts, string projectpath, int projectnumber)
        {
            try
            {
                var companies = Standard_Parts.Distinct(new CaseInsensitiveComparer());
                   
                string basket_path = "";
                int pos_nr = 1;

                if (companies != null)
                {
                    foreach (var company in companies)
                    {
                        if (!string.IsNullOrEmpty(company.manufacturer))
                        {
                            pos_nr = 1;
                            basket_path = projectpath + projectnumber.ToString() + "_" + "Basket" + "_" + company.manufacturer + "_" + DateTime.Today.ToString("yyyyMMdd") + ".csv";
                          //  MessageBox.Show(basket_path);
                            if (!File.Exists(basket_path))
                            {
                                File.Create(basket_path).Close();
                            }
                            string delimter = ";";
                            string capsulate = "\"";


                            int length = Standard_Parts.Count;

                            using (System.IO.TextWriter writer = File.CreateText(basket_path))
                            {
                                for (int index = 0; index < length; index++)
                                {
                                    if (string.Compare(Standard_Parts[index].manufacturer, company.manufacturer, true) == 0)
                                    {
                                        writer.WriteLine(string.Join(delimter, capsulate + pos_nr + capsulate, capsulate + Standard_Parts[index].order_number + capsulate, capsulate + Standard_Parts[index].quantity + capsulate )); //capsulate + projectnumber + capsulate
                                        pos_nr++;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
            }
        }



        /// <summary>
        /// Returns Lists of Custom_parts and Standard_Parts
        /// </summary>
        /// <param name="swModel"></param>
        /// <param name="swTableAnn"></param>
        /// <param name="ConfigName"></param>
        /// <param name="Standard_Parts"></param>
        /// <param name="Custom_Parts"></param>
        public static void Get_Sorted_Part_Data(ModelDoc2 swModel, BomFeature swBomFeat, List<BOM_Part_Informations> Standard_Parts, List<BOM_Part_Informations> Custom_Parts, string projectpath)
        {
            try
            {
                int nNumRow = 0;
                int J = 0;
                int I = 0;
                int numStandard_Part = 1;
                int numCustom_Part = 1;
                int quantity = 0;
                int index_description = 0;
                int index_article_number = 0;
                int index_supplier = 0;


                BOM_Part_Informations part_informations;


               
                string ItemNumber = null;
                string PartNumber = null;

                // Debug.Print("   Table Title        " + swTableAnn.Title);
                Feature swFeat = default(Feature);
                object[] vTableArr = null;
                object vTable = null;
                string[] vConfigArray = null;
                object vConfig = null;
                string ConfigName = null;
                string partconfig = null;

                TableAnnotation swTable = default(TableAnnotation);
                Annotation swAnnotation = default(Annotation);
                object visibility = null;


                swFeat = swBomFeat.GetFeature();
                vTableArr = (object[])swBomFeat.GetTableAnnotations();

                foreach (TableAnnotation vTable_loopVariable in vTableArr)
                {
                    vTable = vTable_loopVariable;
                    swTable = (TableAnnotation)vTable;
                    vConfigArray = (string[])swBomFeat.GetConfigurations(true, ref visibility);

                    foreach (object vConfig_loopVariable in vConfigArray)
                    {
                        vConfig = vConfig_loopVariable;
                        ConfigName = (string)vConfig;



                        //    MessageBox.Show(ConfigName);


                        // swTable.SaveAsPDF(@"C:\Users\alex\Desktop\test.pdf");




                        nNumRow = swTable.RowCount;

                        BomTableAnnotation swBOMTableAnn = default(BomTableAnnotation);
                        swBOMTableAnn = (BomTableAnnotation)swTable;
                        //swTable.GetColumnTitle
                        for (int h = 0; h < swTable.ColumnCount; h++)
                        {
                            switch (swTable.GetColumnTitle(h))
                            {
                                case "Benennung":
                                    index_description = h;
                                    break;
                                case "Artikelnummer":
                                    index_article_number = h;
                                    break;
                                case "Lieferant":
                                    index_supplier = h;
                                    break;
                                default:
                                    break;
                            }
                        }
                        if (index_supplier != 0 || index_supplier != 0 || index_article_number != 0) //Standard BOM Template
                        {

                            for (int n = 0; n <= nNumRow - 1; n++)
                            {
                                // Debug.Print("   Row Number " + J + " Component Count  : " + swBOMTableAnn.GetComponentsCount2(J, ConfigName, out ItemNumber, out PartNumber));
                                //  Debug.Print("       Item Number  : " + ItemNumber);
                                // Debug.Print("       Part Number  : " + PartNumber);
                                // MessageBox.Show("bubu");
                                object[] vPtArr = null;
                                Component2 swComp = null;
                                object pt = null;
                                quantity = swBOMTableAnn.GetComponentsCount2(n, ConfigName, out ItemNumber, out PartNumber);

                                vPtArr = (object[])swBOMTableAnn.GetComponents2(n, ConfigName);

                                if (((vPtArr != null)))
                                {
                                    for (I = 0; I <= vPtArr.GetUpperBound(0); I++)
                                    {
                                        pt = vPtArr[I];
                                        swComp = (Component2)pt;
                                        if ((swComp != null))
                                        {
                                            part_informations = new BOM_Part_Informations();

                                            part_informations.manufacturer = swTable.get_Text(n, index_supplier);
                                            part_informations.order_number = swTable.get_Text(n, index_article_number);

                                            part_informations.part_number = PartNumber;
                                            part_informations.quantity = quantity.ToString();

                                            //Custom part
                                            if (swComp.GetPathName().Contains(projectpath))
                                            {
                                                part_informations.description = swComp.ReferencedConfiguration;
                                                part_informations.item_number = numCustom_Part.ToString();
                                                numCustom_Part++;

                                                Custom_Parts.Add(part_informations);
                                                break;
                                            }

                                            part_informations.description = swTable.get_Text(n, index_description);
                                            part_informations.item_number = numStandard_Part.ToString();
                                            numStandard_Part++;
                                            Standard_Parts.Add(part_informations);
                                            break;

                                        }
                                        else
                                        {
                                            Debug.Print("  Could not get component.");
                                        }
                                    }
                                }
                            }

                        }
                        else //No Standard BOM Template
                        {
                            for (J = 0; J <= nNumRow - 1; J++)
                            {
                                // Debug.Print("   Row Number " + J + " Component Count  : " + swBOMTableAnn.GetComponentsCount2(J, ConfigName, out ItemNumber, out PartNumber));
                                //  Debug.Print("       Item Number  : " + ItemNumber);
                                // Debug.Print("       Part Number  : " + PartNumber);

                                object[] vPtArr = null;
                                Component2 swComp = null;
                                object pt = null;
                                quantity = swBOMTableAnn.GetComponentsCount2(J, ConfigName, out ItemNumber, out PartNumber);

                                vPtArr = (object[])swBOMTableAnn.GetComponents2(J, ConfigName);

                                if (((vPtArr != null)))
                                {
                                    for (I = 0; I <= vPtArr.GetUpperBound(0); I++)
                                    {
                                        pt = vPtArr[I];
                                        swComp = (Component2)pt;
                                        if ((swComp != null))
                                        {
                                            part_informations = new BOM_Part_Informations();

                                            part_informations.description = swComp.ReferencedConfiguration;
                                            part_informations.part_number = PartNumber;
                                            part_informations.quantity = quantity.ToString();
                                            //Custom part
                                            if (swComp.GetPathName().Contains(projectpath))
                                            {

                                                part_informations.item_number = numCustom_Part.ToString();
                                                numCustom_Part++;

                                                Custom_Parts.Add(part_informations);
                                                break;
                                            }

                                            part_informations.item_number = numStandard_Part.ToString();
                                            numStandard_Part++;
                                            Standard_Parts.Add(part_informations);
                                            break;

                                        }
                                        else
                                        {
                                            Debug.Print("  Could not get component.");
                                        }
                                    }
                                }
                            }


                        }
                        break;
                    }

                }
                swAnnotation = swTable.GetAnnotation();
                swAnnotation.Select3(false, null);
                swModel.EditDelete();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);

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


        /// <summary>
        /// Create a List to search items for the order number by company
        /// </summary>
        public static void Create_XML_Websearch_File()
        {

            List<Websearch> websearchlist = new List<Websearch>();
            websearchlist.Add(new Websearch("Festo","https://www.festo.com/net/de_de/SupportPortal/InternetSearch.aspx?q={0}","WarningMessage"));
            websearchlist.Add(new Websearch("Hanser", "http://www.hanser.ch/web/ganter.aspx?cmd=normen&quickfind={0}&LCID=1031&pageID=14##", "0 Treffer"));
            websearchlist.Add(new Websearch("Igus", "http://www.igus.ch/Search?q={0}", "keine Ergebnisse"));
            websearchlist.Add(new Websearch("Würth", "https://eshop.wuerth-ag.ch/is-bin/INTERSHOP.enfinity/WFS/3126-B1-Site/de_DE/-/CHF/ViewAfterSearch-ExecuteAfterSearch?ufd-SearchCategory=Gesamtkatalog&SearchCategory=3126&SearchResultType=&EffectiveSearchTerm=&VisibleSearchTerm={0}&x=9&y=6", "Anzahl gefundene Produkte: 0"));



            try
            {

                if (!File.Exists(SwAddin.path_to_websearch_file))
                {

                    XmlWriterSettings settings = new XmlWriterSettings();

                    // settings.Encoding = Encoding.GetEncoding("UTF-8");
                    settings.Indent = true;
                    settings.IndentChars = "\t";
                    // settings.Indent = true;
                    // settings.NewLineHandling = NewLineHandling.Replace;
                    // settings.IndentChars = " ";
                    // settings.NewLineOnAttributes = true;
                    //  settings.OmitXmlDeclaration = true;



                    using (XmlWriter writer = XmlWriter.Create(SwAddin.path_to_websearch_file, settings))//
                    {
                        writer.WriteStartDocument();

                        writer.WriteStartElement("Websearch_By_Company");

                        foreach (Websearch websearch in websearchlist)
                        {
                            writer.WriteStartElement("Websearch");

                            writer.WriteElementString("ID", websearch.Id);
                            writer.WriteElementString("URL", websearch.Url);
                            writer.WriteElementString("NoMatch", websearch.Nomatchkeyword);


                            writer.WriteEndElement();
                        }

                        writer.WriteEndElement();
                        writer.WriteEndDocument();

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
            }
        }


        /// <summary>
        /// Read the Websearch file 
        /// </summary>
        /// <param name="websearch_list"></param>
        public static void Read_XML_Websearch_File(List<Websearch> websearch_list)
        {

            try
            {
                if (File.Exists(SwAddin.path_to_websearch_file))
                {

                    XmlDocument xdoc = new XmlDocument();
                    xdoc.Load(SwAddin.path_to_websearch_file);

                    foreach (XmlNode websearch in xdoc.SelectNodes("/Websearch_By_Company/*"))
                    {
                        if (websearch != null)
                        {
                            websearch_list.Add(new Websearch(websearch["ID"].InnerText, websearch["URL"].InnerText, websearch["NoMatch"].InnerText));
                            Debug.Print(websearch["ID"].InnerText + websearch["URL"].InnerText + websearch["NoMatch"].InnerText);
                        }

                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
            }
        }

        /// <summary>
        /// Check if URL exists
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static bool page_exists(string url)
        { 
           
            bool pageExists;
            try
            {
                Uri requesturi;
                Uri responseuri;
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = WebRequestMethods.Http.Head;
                request.Timeout = 4000;

                if (request != null)
                {
                    if (request.RequestUri.Scheme == Uri.UriSchemeHttp || request.RequestUri.Scheme == Uri.UriSchemeHttps)
                    {
                        requesturi = request.RequestUri;
                        System.Threading.Thread.Sleep(500); 
                        //ServicePointManager .ServerCertificateValidationCallback +=(sender, cert, chain, sslPolicyErrors) => true;
                        using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                        {
                            responseuri = response.ResponseUri;
                            pageExists = (responseuri == requesturi);
                            //   MessageBox.Show(response.StatusCode.ToString());
                            return pageExists;
                        }
                    }
                    return false;
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
                    return false;
                /*
                if (ex.Status == WebExceptionStatus.ProtocolError)
                {
                    if (((HttpWebResponse)ex.Response).StatusCode == HttpStatusCode.NotFound)
                    {
                        MessageBox.Show("No Ethernet Connection");
                    }
                }
                else if (ex.Status == WebExceptionStatus.NameResolutionFailure)
                {
                    // handle name resolution failure
                }
                return false;*/


            }
            
        }

        /// <summary>
        /// Check all Parts about correct order number
        /// </summary>
        /// <param name="Standard_Parts"></param>
        /// <param name="websearch_list"></param>
        public static void Process_Order_Number(List<BOM_Part_Informations> Standard_Parts, List<Websearch> websearch_list)
        {
            try
            {
                foreach(BOM_Part_Informations part in Standard_Parts)
                {
                    part.valid_order_number = "na";
                    foreach (Websearch company in websearch_list)
                    {
                        if (String.Equals(part.manufacturer, company.Id, StringComparison.OrdinalIgnoreCase))
                        {
                            if (Check_if_item_number_exists(company.Url, part.order_number, company.Nomatchkeyword))
                            {
                                part.valid_order_number = "True";
                            }
                            else
                            {
                                part.valid_order_number = "False";
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
            }
        }

        /// <summary>
        /// Check if Product Number exists
        /// </summary>
        /// <param name="searchurl"></param>
        /// <param name="item_number"></param>
        /// <param name="no_matches"></param>
        public static bool Check_if_item_number_exists(string searchurl,string item_number,string no_matches)
        {

            
            try
            {

                string returnvalue = "";
                if (page_exists(string.Format(searchurl, item_number)))
                {
                    WebRequest req = WebRequest.Create(string.Format(searchurl, item_number));

                    using (WebResponse res = req.GetResponse())
                    {
                        StreamReader sr = new StreamReader(res.GetResponseStream());

                        returnvalue = sr.ReadToEnd();
                    }
                    if (returnvalue.Contains(string.Format(no_matches)))
                    {
                        MessageBox.Show(string.Format("Item Number {0} doesnt exist", item_number));
                        return false;
                    }
                    else
                    {
                     //   MessageBox.Show(string.Format("Teil mit der Nummer {0} der Firma existiert", item_number));
                        return true;

                    }
                    /*
                    if (!File.Exists(@"C:\test.txt"))
                    {
                        File.Create(@"C:\test.txt").Close();
                    }
                    string delimter = ";";
                    string capsulate = "\"";




                    using (System.IO.TextWriter writer = File.CreateText(@"C:\test.txt"))
                    {

                        writer.WriteLine(returnvalue);

                    }*/
                }
                else
                {
                    MessageBox.Show("Url doesnt exist");
                    return false;
                }
                
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.StackTrace);
                return false;
                
            }
        }



        public static void Igus_Order()
        {

            string textbox_id = "ART_NR";
            string value = "\"" +  "234532" + "\"" ;
            string button_id = "234532";
           /* */
            /*
            HtmlDocument document = browser.Document;
            HtmlElement inputValue = document.GetEementById("ctl00_ContentPlaceHolder1_txtNAICS");
            element.SetAttribute("value", "334511");
            HtmlElement submitButton = document.GetElementById("ctl00_ContentPlaceHolder1_btnSearch2");
            submitButton.InvokeMember("click");*/
            try
            {

                       string query = "544234";
                        WebRequest req = WebRequest.Create("https://www.festo.com/net/de_de/SupportPortal/InternetSearch.aspx?q=" + query);
                        /*string postData = "ART_NR=125265";

                        byte[] send = Encoding.Default.GetBytes(postData);
                        req.Method = "POST";
                        req.ContentType = "application/x-www-form-urlencoded";
                        req.ContentLength = send.Length;

                        Stream sout = req.GetRequestStream();
                        sout.Write(send, 0, send.Length);
                        sout.Flush();
                        sout.Close();
                */
                        WebResponse res = req.GetResponse();
                        StreamReader sr = new StreamReader(res.GetResponseStream());
                        string returnvalue = sr.ReadToEnd();
                        if (returnvalue.Contains(string.Format("Ihre Suche nach „{0}“ ergab kein Ergebnis", query)))
                        {
                            MessageBox.Show("Teile Nummer der Firma existiert nicht");
                        }
                        else
                        {
                            MessageBox.Show("Teile Nummer der Firma existiert");

                        }
                        if (!File.Exists(@"C:\test.txt"))
                        {
                            File.Create(@"C:\test.txt").Close();
                        }
                        string delimter = ";";
                        string capsulate = "\"";




                        using (System.IO.TextWriter writer = File.CreateText(@"C:\test.txt"))
                        {

                            writer.WriteLine(returnvalue);
                            
                        }
                        //MessageBox.Show();

                /*
                Process.Start("http://www.igus.de/Quickorder");
                Thread.Sleep(5000);
                SendKeys.SendWait("%D");
                Thread.Sleep(100);
                SendKeys.SendWait(EncodeForSendKey(string.Format(" javascript:function x(){document.getElementById({0}).value={1};} x();", textbox_id, value)));
                SendKeys.SendWait("{ENTER}");*/
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static string EncodeForSendKey(string value)
        {
            StringBuilder sb = new StringBuilder(value);
            sb.Replace("{", "{{}");
            sb.Replace("}", "{}}");
            sb.Replace("{{{}}", "{{}");
            sb.Replace("[", "{[}");
            sb.Replace("]", "{]}");
            sb.Replace("(", "{(}");
            sb.Replace(")", "{)}");
            sb.Replace("+", "{+}");
            sb.Replace("^", "{^}");
            sb.Replace("%", "{%}");
            sb.Replace("~", "{~}");
            return sb.ToString();
        }
    
    }
}
