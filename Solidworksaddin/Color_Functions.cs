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

           try
           {


               foreach (BOM.BOM_Part_Informations part in Standard_Parts)
               {
                   if (!part.IsAssembly)
                   {
                       if (part.IsStandard)
                       {
                       //  MessageBox.Show("part number" + part.part_number);

                           Set_transparency(swModel,part.part_number);

                       }

                       else
                       {
                           Change_Color(swModel, part.part_number, "Red");

                       }

                       if (part.valid_order_number == "False")
                       {
                           Change_Color(swModel, part.part_number, "Green");

                       }


                   }


               }
           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.StackTrace);
           }
       }

       public static void Set_Custom_part_Transparency(ModelDoc2 swModel, List<BOM.BOM_Part_Informations> Standard_Parts, List<BOM.BOM_Part_Informations> Custom_Parts)
       {
           
           try
           {
               

               foreach (BOM.BOM_Part_Informations part in Custom_Parts)
               {
                   if (!part.IsAssembly)
                   {

                       Set_transparency(swModel, part.part_number);
                   }
               }
             

           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.StackTrace);
           }
       }

       public static void Set_transparency(ModelDoc2 swModel, string name)
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
                       if (swComp != null)
                       {
                           if (swComp.Name2.Contains(name))
                           {
                            //   MessageBox.Show(name);

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

       public static void Change_Color(ModelDoc2 swModel, string name,string color)
        {
            AssemblyDoc swAssembly = null;
           
            Component2 swComp = null;
            ModelDoc2 swCompDoc = null;
            string compare_name = "";
            string[] Componentsubstring;
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
                   // MessageBox.Show(Components.Length.ToString());
                    for (int i = 0; i < Components.Length; i++)
                    {

                        //swComp = swAssembly.GetComponentByName(name);
                        swComp = Components[i];
                      //  MessageBox.Show(swComp.Name2);
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
                                if (color != "")
                                {
                                    switch (color)
                                    {
                                        case "Green":
                                             vMatProps[0] = 0; //Red
                                             vMatProps[1] = 1; //Green
                                             vMatProps[2] = 0; //Blue
                                            break;
                                        case "Blue":
                                             vMatProps[0] = 0; //Red
                                             vMatProps[1] = 0; //Green
                                             vMatProps[2] = 1; //Blue
                                            break;
                                        case "Red":
                                             vMatProps[0] = 1; //Red
                                             vMatProps[1] = 0; //Green
                                             vMatProps[2] = 0; //Blue
                                            break;
                                        default:
                                            vMatProps[0] = 1; //Red
                                            vMatProps[1] = 0; //Green
                                            vMatProps[2] = 0; //Blue
                                            break;
                                    }
                                }
                               
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
