using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
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
    class Color_Functions
    {

       public enum swBodyType_e
       {
          swAllBodies = -1,
          swSolidBody = 0,
          swSheetBody = 1,
          swWireBody = 2,
          swMinimumBody = 3,
          swGeneralBody = 4,
          swEmptyBody = 5
       }

       public static void Set_Standard_part_Color(ModelDoc2 swModel,List<BOM.BOM_Part_Informations> Standard_Parts, List<BOM.BOM_Part_Informations> Custom_Parts)
       {
           string swDocname;
           string[] docname;
           try
           {
                 swDocname = swModel.GetTitle();
                 docname = swDocname.Split('.');

                 foreach(BOM.BOM_Part_Informations part in Standard_Parts)
                 {

                    Set_transparency(swModel,part.part_number,docname[0]);
                 }
             

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.StackTrace);
           }
       }

       public static void Set_Custom_part_Transparency(ModelDoc2 swModel, List<BOM.BOM_Part_Informations> Standard_Parts, List<BOM.BOM_Part_Informations> Custom_Parts)
       {
           string swDocname;
           string[] docname;
           try
           {
               swDocname = swModel.GetTitle();
               docname = swDocname.Split('.');

               foreach (BOM.BOM_Part_Informations part in Custom_Parts)
               {
                   for(int i = 1; i <= int.Parse(part.quantity) ; i++)
                   {
                     Set_transparency(swModel, part.part_number + "-" + i.ToString(), docname[0]);
                   }
               }
               //    MessageBox.Show(swDocname);

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.StackTrace);
           }
       }

       public static void Set_transparency(ModelDoc2 swModel, string name,string docname)
       {
           AssemblyDoc swAssembly = null;
           bool boolstatus = false;
           try
           {
               if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
               {
                  // MessageBox.Show(name + "@" + docname);
                   swModel.ClearSelection2(true);
                   
                   boolstatus = swModel.Extension.SelectByID2(name + "@" + docname, "COMPONENT", 0, 0, 0, true, 0, null, 0);
                   swAssembly = ((AssemblyDoc)(swModel));
                  // MessageBox.Show(boolstatus.ToString());
                   swAssembly.SetComponentTransparent(true);
                   swModel.ClearSelection2(true);

               }
           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.StackTrace);
           }

           //  boolstatus = swDoc.Extension.SelectByID2("195599_GRLA_F_1_8_QS_8_D-3@toolbox-tutorial", "COMPONENT", 0, 0, 0, false, 0, null, 0);


       }

        public static void Change_Color(ModelDoc2 swModel,string name)
        {
            AssemblyDoc swAssembly = null;
            try
            {
                if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    swAssembly = ((AssemblyDoc)(swModel));
                    swModel.ClearSelection2(true);
                    swModel.Extension.SelectByID2(name, "COMPONENT", 0, 0, 0, true, 0, null, 0);
                    swAssembly.SetComponentTransparent(true);
                    swModel.ClearSelection2(true);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
            }

          //  boolstatus = swDoc.Extension.SelectByID2("195599_GRLA_F_1_8_QS_8_D-3@toolbox-tutorial", "COMPONENT", 0, 0, 0, false, 0, null, 0);


        }
    }
}
