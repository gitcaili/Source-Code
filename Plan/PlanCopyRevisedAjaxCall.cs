using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;
using System.Data;
using System.IO;
using Sage.CRM.UI;
using System.Diagnostics;
using System.Globalization;
using Sage.CRM.Wrapper;
using Sage.CRM.Utils;

namespace NZPACRM.Plan
{
    class PlanCopyRevisedAjaxCall : Web
    {
        CRMHelper objCRM = new CRMHelper();
        int intNextBookingID;
        int strEntityId = 0;
        int strEntityIdNew = 0;

        string strEntityNamePrimary = "Booking";
        string strEntityNameSeconday = "Planbuilder";
        string strBookName;
        string strEntityWorkflowName = "";
        string strNewCasesNextWorkflowState = "";

        int intHiddenRowCount = 0;
        int intNextID = 0;
        int intGridID = 0;
        string sURL = "";
        string strMessage = "";
        string sPlanIdString = "";
        string sPlanId = "";

        public PlanCopyRevisedAjaxCall()
        {

        }

        public override void BuildContents()
        {
            try
            {
                if (!String.IsNullOrEmpty(Dispatch.EitherField("book_bookingid")))
                    strEntityId = Convert.ToInt32(Dispatch.EitherField("book_bookingid"));
                else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key58")))
                    strEntityId = Convert.ToInt32(Dispatch.EitherField("Key58"));

                if (!String.IsNullOrEmpty(Dispatch.EitherField("HiddenRowCount")))
                    intHiddenRowCount = Convert.ToInt32(Dispatch.EitherField("HiddenRowCount")) + 1;

                if (!String.IsNullOrEmpty(Dispatch.EitherField("NextID")))
                    intNextID = Convert.ToInt32(Dispatch.EitherField("NextID"));

                if (!String.IsNullOrEmpty(Dispatch.EitherField("GridID")))
                    intGridID = Convert.ToInt32(Dispatch.EitherField("GridID"));

                if (strEntityId > 0)
                {
                    #region Copy Record
                    CopyPlanRecord(strEntityNamePrimary, strEntityId);
                    #endregion

                    #region Copy Plan Builder Record
                    string PlanBuilderID = "";
                    Record RecPlanBuilder = FindRecord(strEntityNameSeconday, "pnbr_plan=" + strEntityId);

                    if (!RecPlanBuilder.Eof())
                    {
                        while (!RecPlanBuilder.Eof())
                        {
                            if (!String.IsNullOrEmpty(RecPlanBuilder.GetFieldAsString("pnbr_Pnbr_planbuilderid")))
                                PlanBuilderID = RecPlanBuilder.GetFieldAsString("pnbr_Pnbr_planbuilderid");

                            #region selected Plan Id
                            //AddContent("<BR>PlanID=" + PlanID);
                            Record RecEachPlanBuilder = FindRecord(strEntityNameSeconday, "pnbr_Pnbr_planbuilderid=" + PlanBuilderID);

                            if (!RecEachPlanBuilder.Eof())
                            {
                                //AddContent("<BR>PlanID 1=" + PlanID);
                                Record objPlanBuilder = new Record(strEntityNameSeconday);
                                objPlanBuilder.SetField("pnbr_plan", intNextBookingID);
                                objPlanBuilder.SetField("pnbr_publications", RecEachPlanBuilder.GetFieldAsString("pnbr_publications"));
                                objPlanBuilder.SetField("pnbr_ratecard", RecEachPlanBuilder.GetFieldAsString("pnbr_ratecard"));
                                objPlanBuilder.SetField("pnbr_sections", RecEachPlanBuilder.GetFieldAsString("pnbr_sections"));
                                objPlanBuilder.SetField("pnbr_other", RecEachPlanBuilder.GetFieldAsString("pnbr_other"));
                                objPlanBuilder.SetField("pnbr_subsection", RecEachPlanBuilder.GetFieldAsString("pnbr_subsection"));
                                objPlanBuilder.SetField("pnbr_days", RecEachPlanBuilder.GetFieldAsString("pnbr_days"));
                                objPlanBuilder.SetField("pnbr_date", RecEachPlanBuilder.GetFieldAsString("pnbr_date"));
                                objPlanBuilder.SetField("pnbr_size", RecEachPlanBuilder.GetFieldAsString("pnbr_size"));
                                objPlanBuilder.SetField("pnbr_height", RecEachPlanBuilder.GetFieldAsString("pnbr_height"));
                                objPlanBuilder.SetField("pnbr_width", RecEachPlanBuilder.GetFieldAsString("pnbr_width"));
                                objPlanBuilder.SetField("pnbr_custom", RecEachPlanBuilder.GetFieldAsString("pnbr_custom"));
                                objPlanBuilder.SetField("pnbr_color", RecEachPlanBuilder.GetFieldAsString("pnbr_color"));
                                objPlanBuilder.SetField("pnbr_loading", RecEachPlanBuilder.GetFieldAsString("pnbr_loading"));
                                objPlanBuilder.SetField("pnbr_standardrate", RecEachPlanBuilder.GetFieldAsString("pnbr_standardrate"));
                                objPlanBuilder.SetField("pnbr_loadingvalue", RecEachPlanBuilder.GetFieldAsString("pnbr_loadingvalue"));
                                objPlanBuilder.SetField("pnbr_discount", RecEachPlanBuilder.GetFieldAsString("pnbr_discount"));
                                objPlanBuilder.SetField("pnbr_total", RecEachPlanBuilder.GetFieldAsString("pnbr_total"));


                                //objPlanBuilder.SetField("pnbr_plan", intNextBookingID);
                                objPlanBuilder.SaveChanges();
                            }
                            RecPlanBuilder.GoToNext();
                        }
                            #endregion

                        //CopyPlanBuliderRecord(strEntityNameSeconday, Convert.ToInt32(PlanBuilderID));
                        //RecPlanBuilder.SetField("book_docversion", strEntityId);
                        //RecPlanBuilder.SaveChanges();

                    }
                    #endregion
                }
                AddContent("<returnmsg>" + intNextBookingID + "</returnmsg>");
                //NavigateToSummary(strEntityNamePrimary, strEntityId, strEntityId);
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
        }

        private void CopyPlanRecord(string entityName, int entityId)
        {
            try
            {
                TableInfo tableInfo = Metadata.GetTableInfo(entityName);
                Record existingRecord = FindRecord(entityName, String.Format("{0}={1}", tableInfo.IdField, entityId));
                existingRecord.GoToFirst();
                Record newRecord = new Record(entityName);

                IeWareRecord enumerableExistingRecord = GetEnumerableRecord(existingRecord);
                IeWareRecord enumerableNewRecord = GetEnumerableRecord(newRecord);

                int j = 0;
                string refID = "";

                foreach (object i in enumerableExistingRecord)
                {
                    string p = i.ToString();
                    if (p != tableInfo.IdField)
                    {
                        try
                        {
                            j++;
                            enumerableNewRecord[p] = enumerableExistingRecord[p];

                            if (entityName == "Booking")
                            {
                                if (j == 1)
                                {
                                    refID = GenerateSequecenumber(entityId.ToString());
                                }

                                if (String.IsNullOrEmpty(refID) || refID == "")
                                    strBookName = enumerableExistingRecord["book_Name"] + "-Cloned";
                                else
                                    strBookName = enumerableExistingRecord["book_reference"] + "-Cloned";
                                                                 
                                enumerableNewRecord["book_docversion"] = "";
                                enumerableNewRecord["book_revisedBookId"] = entityId;
                                enumerableNewRecord["book_Agency"] = enumerableExistingRecord["book_Agency"];
                                enumerableNewRecord["book_Contact"] = enumerableExistingRecord["book_Contact"];
                                enumerableNewRecord["book_Client"] = enumerableExistingRecord["book_Client"];
                                enumerableNewRecord["book_agencycode"] = enumerableExistingRecord["book_agencycode"];
                                enumerableNewRecord["book_reference"] = refID;
                                enumerableNewRecord["book_CreatedBy"] = this.CurrentUser.UserId;
                                enumerableNewRecord["book_Name"] = enumerableExistingRecord["book_Name"];
                                enumerableNewRecord["book_billedby"] = enumerableExistingRecord["book_billedby"];
                                enumerableNewRecord["book_description"] = enumerableExistingRecord["book_description"];
                                enumerableNewRecord["book_costingversion"] = enumerableExistingRecord["book_costingversion"];

                                enumerableNewRecord.SaveChanges();
                                intNextBookingID = enumerableNewRecord.RecordID;

                                
                            }
                        }
                        catch (Exception ex)
                        {
                            //handle exception
                            this.AddError(ex.Message);
                            /*sURL = Url("281");
                            strMessage = "Error" + ex.Message;
                            objCRM.GetStatusBlock(entityName, strMessage, "false", entityId.ToString());

                            AddUrlButton("Continue", "continue.gif", sURL);*/
                        }
                    }
                }
                int bookrev = 0;
                Record objPlanRecord = FindRecord("Booking", "book_BookingID='" + entityId + "' and book_Deleted is null");
                Record objPlanRevRecord = FindRecord("Booking", "book_revisedBookId='" + entityId + "' and book_Deleted is null");                    
                //if (!String.IsNullOrEmpty(objPlanRecord.GetFieldAsInt("book_docversion").ToString()))
                //    bookrev = objPlanRecord.GetFieldAsInt("book_docversion");
                bookrev = objPlanRevRecord.RecordCount;
                //bookrev++;                               
                objPlanRecord.SetField("book_docversion", bookrev);
                objPlanRecord.SetField("book_Status", "Inactive");
                objPlanRecord.SaveChanges();
            }
            catch (Exception ex)
            {
                /*string strMessage = "Plan unable to clone. " + ex.Message;
                objCRM.GetStatusBlock(entityName, strMessage, "false", entityId.ToString());

                sURL = Url("432");
                strMessage = "Error" + ex.Message;
                objCRM.GetStatusBlock(entityName, strMessage, "false", entityId.ToString());

                AddUrlButton("Continue", "continue.gif", sURL);*/
                this.AddError(ex.Message);
            }
        }

        private void NavigateToSummary(string entityName, int entityId, int intNextBookingID)
        {
            if (intNextBookingID != entityId)
            {
                strMessage = "Plan is successfully cloned. Click Continue to navigate to New Plan.";
                objCRM.GetStatusBlock(entityName, strMessage, "", intNextBookingID.ToString());
                sURL = UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage&Key58=" + intNextBookingID + "&Key37=" + intNextBookingID + "&book_bookingid=" + intNextBookingID);

                AddUrlButton("Continue", "continue.gif", sURL);
            }
            else
            {
                strMessage = "Unable to clone Plan. Click Continue to navigate to previous Plan.";
                objCRM.GetStatusBlock(entityName, strMessage, "false", entityId.ToString());
                sURL = UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage&Key58=" + entityId + "&Key37=" + entityId + "&book_bookingid=" + entityId);
                AddUrlButton("Continue", "continue.gif", sURL);
            }
        }

        private IeWareRecord GetEnumerableRecord(Sage.CRM.Data.Record rec)
        {
            System.Type t = rec.GetType();
            System.Reflection.MethodInfo mi = t.GetMethod("GetInternalEntity", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            return (Sage.CRM.Wrapper.IeWareRecord)mi.Invoke(rec, new object[] { });
        }

        public string GetStatusBlock(string EnityName, string StatusMsg, string isValidColumn, string sRowCount)
        {
            string InstructionText = "";
            string sHTML = "";

            sHTML += HTML.StartTable();
            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData("<span style='float:left;'>Status Type</span><br />", "GRIDHEAD", "width=150px height=25px");
            sHTML += HTML.TableData("<span style='float:left;'>Application Message</span>", "GRIDHEAD");
            sHTML += HTML.TableRow("");
            if (isValidColumn == "")
            {
                if (sRowCount != "N")
                {
                    sHTML += HTML.TableData("&nbsp;&nbsp;<span font-Size:3px;'>Success</span>", "VIEWBOX");
                }
                else
                    sHTML += HTML.TableData("&nbsp;&nbsp;<span font-Size:3px;'>Error</span>", "VIEWBOX");
            }
            else if (isValidColumn == "false")
                sHTML += HTML.TableData("&nbsp;&nbsp;<span font-Size:3px;'>Error</span>", "VIEWBOX");
            sHTML += HTML.TableData("<span font-Size:10px;'>" + StatusMsg + "</span>", "VIEWBOX");
            sHTML += "<BR>";
            sHTML += HTML.TableRow("");
            sHTML += HTML.EndTable();
            sHTML += "<BR>";
            AddContent(HTML.Box("<span style='color:#2B547E;'>Sage CRM Application Status</span>", sHTML));
            return "";
        }
        private string GenerateSequecenumber(string entityId)
        {
            string strUniqueID = "";
            try
            {
                Record recCustomSysParams = FindRecord("custom_sysparams", "Parm_Name = 'Bookingbook_BookingID'");
                if (!recCustomSysParams.Eof())
                {
                    int intUniqueID = recCustomSysParams.GetFieldAsInt("Parm_Value");

                    if (intUniqueID > 0)
                    {
                        intUniqueID = intUniqueID + 1;
                        recCustomSysParams.SetField("Parm_Value", intUniqueID);
                        recCustomSysParams.SaveChanges();
                        strUniqueID = intUniqueID.ToString();
                        DateTime today = DateTime.Now;
                        int YY = today.Year;
                        int MM = today.Month;
                        int DD = today.Day;
                        StringBuilder builder = new StringBuilder();
                        string year = YY.ToString();
                        strUniqueID = this.CurrentUser.UserId.ToString() + "-" + Convert.ToInt32(strUniqueID.ToString()).ToString("D5");

                    }
                }
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
            return strUniqueID;
        }
    }
}
