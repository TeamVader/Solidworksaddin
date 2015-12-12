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



        /// <summary>
        /// Create a List to search items for the order number by company
        /// </summary>
        public static void Create_XML_Websearch_File()
        {

            List<Websearch> websearchlist = new List<Websearch>();
            websearchlist.Add(new Websearch("Festo","https://www.festo.com/net/de_de/SupportPortal/InternetSearch.aspx?q={0}","WarningMessage"));
            websearchlist.Add(new Websearch("Hanser", "http://www.hanser.ch/web/ganter.aspx?cmd=normen&quickfind={0}&LCID=1031&pageID=14##", "0 Treffer"));
            websearchlist.Add(new Websearch("Igus", "http://www.igus.ch/Search?q={0}", "Artikel_SucheResultText" + "\"" + ">keine Ergebnisse"));
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
                request.Timeout = 6000;

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
               // MessageBox.Show(ex.StackTrace);
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

                        if (!File.Exists(string.Format(@"C:\{0}.txt", item_number)))
                        {
                            File.Create(string.Format(@"C:\{0}.txt", item_number)).Close();
                        }
                        string delimter = ";";
                        string capsulate = "\"";




                        using (System.IO.TextWriter writer = File.CreateText(string.Format(@"C:\{0}.txt", item_number)))
                        {

                            writer.WriteLine(returnvalue);

                        }

                        return false;
                    }
                    else
                    {
                     //   MessageBox.Show(string.Format("Teil mit der Nummer {0} der Firma existiert", item_number));
                        return true;

                    }

                    
                }
                else
                {
                    MessageBox.Show(string.Format("Url : {0} doesnt exist", searchurl));
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
