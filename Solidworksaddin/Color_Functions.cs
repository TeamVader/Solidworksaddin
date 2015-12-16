﻿using Microsoft.Office.Interop.Excel;
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
                    
                         Change_Color(swModel, part.part_number, docname[0]);
                     
                    
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
                   
                     Set_transparency(swModel, part.part_number, docname[0]);
                  
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

           Component2 swComp = null;
           ModelDoc2 swCompDoc = null;
           bool boolstatus = false;

           try
           {
               if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
               {

                   swModel.ClearSelection2(true);
                   //  boolstatus = swModel.Extension.SelectByID2(name + "@" + docname, "COMPONENT", 0, 0, 0, true, 0, null, 0);
                   //  SelectionMgr SwSelMgr = swModel.SelectionManager;
                   swAssembly = (AssemblyDoc)swModel;
                   var Components = swAssembly.GetComponents(false);

                   for (int i = 0; i < Components.Length; i++)
                   {

                       //swComp = swAssembly.GetComponentByName(name);
                       swComp = Components[i];
                       // MessageBox.Show(name);
                       if (swComp != null)
                       {
                           if (swComp.Name2.Contains(name))
                           {
                               var vMatProps = swComp.MaterialPropertyValues;
                               if (vMatProps == null)
                               {
                                   swCompDoc = swComp.GetModelDoc();
                                   if (swCompDoc == null)
                                   {
                                       return;
                                   }
                                   vMatProps = swCompDoc.MaterialPropertyValues;
                               }
                               vMatProps[7] = 1; //Transparency
                               
                               swComp.MaterialPropertyValues = vMatProps;
                               swModel.ClearSelection2(true);
                           }


                       }
                   }
               }
           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message + ex.StackTrace);
           }


       }

       public static void Change_Color(ModelDoc2 swModel, string name, string docname)
        {
            AssemblyDoc swAssembly = null;
           
            Component2 swComp = null;
            ModelDoc2 swCompDoc = null;
            bool boolstatus = false;

            try
            {
                if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {

                    swModel.ClearSelection2(true);
                    //  boolstatus = swModel.Extension.SelectByID2(name + "@" + docname, "COMPONENT", 0, 0, 0, true, 0, null, 0);
                    //  SelectionMgr SwSelMgr = swModel.SelectionManager;
                    swAssembly = (AssemblyDoc)swModel;
                    var Components = swAssembly.GetComponents(false);

                    for (int i = 0; i < Components.Length; i++)
                    {

                        //swComp = swAssembly.GetComponentByName(name);
                        swComp = Components[i];
                        // MessageBox.Show(name);
                        if (swComp != null)
                        {
                            if (swComp.Name2.Contains(name))
                            {
                                var vMatProps = swComp.MaterialPropertyValues;
                                if (vMatProps == null)
                                {
                                    swCompDoc = swComp.GetModelDoc();
                                    if (swCompDoc == null)
                                    {
                                        return;
                                    }
                                    vMatProps = swCompDoc.MaterialPropertyValues;
                                }
                                vMatProps[0] = 1; //Red
                                vMatProps[1] = 0; //Green
                                vMatProps[2] = 0; //Blue
                                swComp.MaterialPropertyValues = vMatProps;
                                swModel.ClearSelection2(true);
                            }


                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }

          //  boolstatus = swDoc.Extension.SelectByID2("195599_GRLA_F_1_8_QS_8_D-3@toolbox-tutorial", "COMPONENT", 0, 0, 0, false, 0, null, 0);


        }
    }
}
