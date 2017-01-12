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

namespace NZPACRM.Plan
{
    public class PlanSummaryPage : Web
    {
        CRMHelper objCrmHelp = new CRMHelper();

        int iPlanID = 0;
        int iEntityId = 0;
        string sHostName = "";
        string sInstallName = "";
        string sPlanID = "";
        int iPlanBuilderId = 0;
        public PlanSummaryPage()
            : base()
        {
            AddContent(HTML.InputHidden("HiddenRecentPlanid", ""));
            string sAct = "";

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Act")))
                sAct = Dispatch.EitherField("Act");
            else
                sAct = "";

            #region Get current Plan id

            if (!String.IsNullOrEmpty(Dispatch.EitherField("book_bookingid")))
            {
                iPlanID = Convert.ToInt32(Dispatch.EitherField("book_bookingid"));
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key58")))
            {
                iPlanID = Convert.ToInt32(Dispatch.EitherField("Key58"));
            }
            else
            {
                iPlanID = Convert.ToInt32(GetContextInfo("Booking", "book_bookingid"));
            }
            #endregion

            #region Redirect to if key37 is not present

            if (!String.IsNullOrEmpty(Dispatch.EitherField("RecentValue")))
            {
                //AddContent("bookid" + Dispatch.EitherField("RecentValue"));
                var sRecPlanId = Dispatch.EitherField("RecentValue");
                sRecPlanId = sRecPlanId.Replace("432X", "");
                string sUrl = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/" + "eware.dll/Do?SID=" + Dispatch.EitherField("SID") + "&Act=432&Mode=1&CLk=T&Key0=58&Key58=" + sRecPlanId + "&Key37=" + sRecPlanId + "&book_BookingID=" + sRecPlanId + "&dotnetdll=NZPACRM&dotnetfunc=RunPlanSummaryPage&J=Booking%20Summary&J=Booking%20Summary&T=Booking";
                //AddContent("<BR>"+sUrl);
                Dispatch.Redirect(sUrl);
            }
            #endregion

            string sPlanStatus = "Select * from Booking where book_bookingid=" + iPlanID;

            QuerySelect objplanRec = GetQuery();
            objplanRec.SQLCommand = sPlanStatus;
            objplanRec.ExecuteReader();

            if (!objplanRec.Eof())
            {
                if (!String.IsNullOrEmpty(objplanRec.FieldValue("book_deleted")))
                {
                    Dispatch.Redirect(UrlDotNet(ThisDotNetDll, "RunRateCardRedirectorPage"));
                }
            }

            #region Set WorkFlow

            #endregion

            try
            {
                Record objEntityIdRec = FindRecord("custom_tables", "bord_name='Booking' and bord_deleted is null ");

                if (!objEntityIdRec.Eof())
                    iEntityId = objEntityIdRec.GetFieldAsInt("Bord_tableid");

                if (!String.IsNullOrEmpty(Dispatch.Host))
                {
                    CurrentUser.SessionWrite("HostName", Dispatch.Host);
                }

                if (!String.IsNullOrEmpty(Dispatch.InstallName))
                {
                    CurrentUser.SessionWrite("InstallName", Dispatch.InstallName);
                }
            }
            catch (Exception ex)
            {

            }

            if (CurrentUser.SessionRead("HostName") != null)
            {
                sHostName = CurrentUser.SessionRead("HostName").ToString();
            }

            if (CurrentUser.SessionRead("InstallName") != null)
            {
                sInstallName = CurrentUser.SessionRead("InstallName").ToString();
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Act")))
            {
                if (Dispatch.EitherField("Act").ToLower() == "523")
                {
                    Dispatch.Redirect(UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&book_bookingid=" + iPlanID + "&T=Booking");
                }
            }

        }

        public override void BuildContents()
        {

            base.OnLoad = "javascript:SetPlanTopContext();";
            AddContent(HTML.Form());
            EntryGroup objPlanSummaryBox = new EntryGroup("BookingNewEntry", "Booking");
            EntryGroup objBuilderBox = new EntryGroup("PlanBuilderScreen", "Planbuilder");
            AddTopContent(GetCustomEntityTopFrame("Booking"));
            string hMode = "";

            if (!String.IsNullOrEmpty(Dispatch.EitherField("HiddenMode")))
            {
                hMode = Dispatch.EitherField("HiddenMode");
            }

            if (hMode == "edit")
            {
                Record objPlanRecRec = FindRecord("Booking", "book_bookingid='" + iPlanID + "' and book_Deleted is null");
                Record objBuilderRec = FindRecord("Planbuilder", "pnbr_plan=" + iPlanID + " and pnbr_Deleted is null");
                AddContent(HTML.InputHidden("HiddenMode", ""));
                objPlanSummaryBox.Fill(objPlanRecRec);
                objBuilderBox.Fill(objBuilderRec);
                AddContent(objPlanSummaryBox.GetHtmlInEditMode(objPlanRecRec));
                AddContent("<BR>");
                if (!objBuilderRec.Eof())
                {
                    //AddContent("IF");
                    //AddContent(objBuilderBox.GetHtmlInEditMode(objBuilderRec));
                }
                else
                {
                    //AddContent("ELSE");
                    //AddContent(objBuilderBox.GetHtmlInEditMode());
                }

                AddContent(HTML.InputHidden("HiddenPlanMode", ""));
                AddContent(HTML.InputHidden("HiddenSelectedDays", ""));
                //AddSubmitButton("Save", "Save.gif", "javascript:document.EntryForm.HiddenMode.value='save';document.EntryForm.submit();");
                AddSubmitButton("Save", "Save.gif", "javascript:document.EntryForm.HiddenMode.value='save';GetPlanBuilderSelectedDays();document.EntryForm.submit();");
                AddUrlButton("Cancel", "cancel.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&book_bookingid=" + iPlanID + "&From=edit");
                //'AddUrlButton("Delete", "delete.gif", UrlDotNet(ThisDotNetDll, "RunBuildingDeletePage") + "&J=Summary&book_bookingid=" + iPlanID);
                if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_pnbr_planbuilderid")))
                {
                    AddUrlButton("Close Plan", "close.gif", UrlDotNet(ThisDotNetDll, "RunPlanClosePage"));
                }
            }
            else if (hMode == "save")
            {
                if (objPlanSummaryBox.Validate() == true && objBuilderBox.Validate() == true)
                {
                    string sBuilderId = "";
                    #region Save Plan Record
                    Record objPlanRec = FindRecord("Booking", "book_bookingid='" + iPlanID + "' and book_Deleted is null");
                    objPlanSummaryBox.Fill(objPlanRec);
                    objPlanRec.SaveChanges();
                    #endregion

                    #region PlanBuilder code

                    #region Planbuilder Item based on days

                    string sSelectedDays = "";
                    string sPublicationID = "";
                    string sStandardSize = "";
                    string scommissiontype = "";
                    string sSections = "";
                    double sStandardRate = 0;
                    //AddContent("My DLL:-" + Dispatch.ContentField("pnbr_date"));
                    if (objCrmHelp.CheckDateIsExists(Dispatch.ContentField("pnbr_date")))
                    {
                        //AddContent("Days" + Dispatch.ContentField("HiddenSelectedDays"));
                        if (!String.IsNullOrEmpty(Dispatch.ContentField("HiddenSelectedDays")))
                        {
                            sSelectedDays = Dispatch.ContentField("HiddenSelectedDays");
                        }
                        if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_publications")))
                        {
                            sPublicationID = Dispatch.ContentField("pnbr_publications");
                        }
                        if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_standardsize")))
                        {
                            sStandardSize = Dispatch.ContentField("pnbr_standardsize");
                        }
                        if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_commissiontype")))
                        {
                            scommissiontype = Dispatch.ContentField("pnbr_commissiontype");
                        }
                        if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_sections")))
                        {
                            sSections = Dispatch.ContentField("pnbr_sections");
                        }
                        if (sSelectedDays != "")
                        {
                            string sDaysQuery = "";
                            string[] sDays = sSelectedDays.Split(',');
                            //AddContent("sDays.Length = " + sDays.Length);
                            if (sDays.Length > 1)
                            {
                                double sFinalTotal = 0;
                                double sDiscount = 0;
                                double sLoadingValue = 0;
                                double sPercentLoadingValue = 0;
                                //Record objPlanRec1 = FindRecord("Booking", "book_bookingid='" + iPlanID + "' and book_Deleted is null");
                                //objPlanSummaryBox.Fill(objPlanRec1);
                                //AddContent(objPlanSummaryBox);
                                //AddContent("<BR>");
                                #region Code to add record based on multiple selected Days
                                foreach (string sDay in sDays)
                                {
                                    if (sDay == "Mon")
                                        sDaysQuery = "rate_Monday";

                                    if (sDay == "Tues")
                                        sDaysQuery = "rate_tuesday";

                                    if (sDay == "Wed")
                                        sDaysQuery = "rate_wednesday";

                                    if (sDay == "Thur")
                                        sDaysQuery = "rate_thrusday";

                                    if (sDay == "Fri")
                                        sDaysQuery = "rate_friday";

                                    if (sDay == "Sat")
                                        sDaysQuery = "rate_saturday";

                                    if (sDay == "Sun")
                                        sDaysQuery = "rate_sunday";

                                    string strSQL = "select " + sDaysQuery + " from RatesCard ";
                                    strSQL += " where rate_deleted is null ";
                                    strSQL += " and rate_PublicationsID='" + sPublicationID + "' ";
                                    strSQL += " and rate_standardsize='" + sStandardSize + "' ";
                                    strSQL += " and rate_commissiontype='" + scommissiontype + "' ";
                                    strSQL += " and rate_section='" + sSections + "' ";

                                    QuerySelect objPblcnRec = GetQuery();
                                    objPblcnRec.SQLCommand = strSQL;
                                    objPblcnRec.ExecuteReader();

                                    if (!objPblcnRec.Eof())
                                    {
                                        sStandardRate = Convert.ToDouble(objPblcnRec.FieldValue(sDaysQuery));
                                    }
                                    //AddContent("<BR>StandardRate = " + sStandardRate + "<BR>SQL =" + strSQL + "<BR> sDay=" + sDay);
                                    if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_loading")))
                                    {
                                        #region Discount Value
                                        if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_discount")))
                                        {
                                            sDiscount = Convert.ToDouble(Dispatch.ContentField("pnbr_discount"));
                                        }
                                        #endregion
                                        #region Calculation
                                        if (Dispatch.ContentField("pnbr_loading").ToString() == "Dollar")
                                        {
                                            if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_loadingvalue")))
                                            {
                                                sLoadingValue = Convert.ToDouble(Dispatch.ContentField("pnbr_loadingvalue"));
                                            }
                                            if (sStandardRate > 0 && sLoadingValue > 0 && sDiscount == 0)
                                            {
                                                sFinalTotal = Convert.ToDouble(sStandardRate) + Convert.ToDouble(sLoadingValue);
                                                //AddContent("<br>sFinalTotal = " + sFinalTotal);
                                            }
                                            else if (sStandardRate > 0 && sLoadingValue == 0 && sDiscount == 0)
                                            {
                                                sFinalTotal = sStandardRate;
                                            }
                                            else
                                            {
                                                if (sStandardRate == 0 && sLoadingValue > 0 && sDiscount == 0)
                                                    sFinalTotal = Convert.ToDouble(sLoadingValue);
                                                else if (sStandardRate == 0 && sLoadingValue > 0 && sDiscount > 0)
                                                    sFinalTotal = sLoadingValue - sDiscount;
                                                else
                                                    sFinalTotal = Convert.ToDouble(sStandardRate) + Convert.ToDouble(sLoadingValue) - Convert.ToDouble(sDiscount);

                                                //AddContent("<br>sFinalTotal Discount = " + sFinalTotal + " = sLoadingValue =" + sLoadingValue + " = sDiscount =" + sDiscount);
                                            }
                                        }
                                        else if (Dispatch.ContentField("pnbr_loading").ToString() == "Percentage")
                                        {
                                            if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_percentvalue")))
                                            {
                                                sPercentLoadingValue = Convert.ToDouble(Dispatch.ContentField("pnbr_percentvalue"));
                                            }
                                            if (sStandardRate > 0 && sPercentLoadingValue > 0 && sDiscount == 0)
                                            {
                                                sFinalTotal = Convert.ToDouble(sStandardRate / 100) * Convert.ToDouble(sPercentLoadingValue);
                                                sFinalTotal = sFinalTotal + sStandardRate;
                                            }
                                            else if (sStandardRate > 0 && sPercentLoadingValue == 0 && sDiscount == 0)
                                            {
                                                sFinalTotal = sStandardRate;
                                            }
                                            else
                                            {
                                                if (sStandardRate == 0 && sPercentLoadingValue > 0 && sDiscount == 0)
                                                    sFinalTotal = 0;
                                                else if (sStandardRate == 0 && sPercentLoadingValue > 0 && sDiscount > 0)
                                                    sFinalTotal = 0;
                                                else
                                                    sFinalTotal = Convert.ToDouble(sStandardRate / 100) * Convert.ToDouble(sDiscount);
                                                sFinalTotal = Convert.ToDouble(sStandardRate) - Convert.ToDouble(sFinalTotal);
                                                double sLoadingVal = Convert.ToDouble(sStandardRate / 100) * Convert.ToDouble(sPercentLoadingValue);
                                                sFinalTotal = Convert.ToDouble(sFinalTotal) + Convert.ToDouble(sPercentLoadingValue);
                                            }
                                            //AddContent("<br>sFinalTotal Discount = " + sFinalTotal + " = sPercentLoadingValue =" + sPercentLoadingValue + " = sDiscount =" + sDiscount);
                                        }
                                        #endregion
                                    }
                                    if (String.IsNullOrEmpty(Dispatch.EitherField("pnbr_pnbr_planbuilderid")))
                                    {                                         
                                        #region Add new Plan Builder
                                        //AddContent("HERE pnbr_publications" + Dispatch.EitherField("pnbr_publications") + "<BR> pnbr_publisher" + Dispatch.EitherField("pnbr_publisher"));

                                        if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_publications")))
                                        {
                                            Record objBuilderRec = new Record("Planbuilder");
                                            objBuilderRec.SetField("pnbr_plan", iPlanID);
                                            objBuilderBox.Fill(objBuilderRec);
                                            objBuilderRec.SaveChanges();

                                            Record newobjBuilderRec = FindRecord("Planbuilder", "pnbr_Pnbr_planbuilderid =" + objBuilderRec.RecordId.ToString());
                                            newobjBuilderRec.SetField("pnbr_total", sFinalTotal);
                                            newobjBuilderRec.SetField("pnbr_days", "," + sDay + ",");
                                            newobjBuilderRec.SetField("pnbr_standardrate", sStandardRate);
                                            if (Dispatch.ContentField("pnbr_loading").ToString() == "Percentage" && sStandardRate == 0)
                                                newobjBuilderRec.SetField("pnbr_percentvalue", 0);
                                            newobjBuilderRec.SaveChanges();
                                        }
                                        else if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_publisher")))
                                        {
                                            Record recPublications = FindRecord("publications", "pblc_PublishersID =" + Dispatch.EitherField("pnbr_publisher"));
                                            if (!recPublications.Eof())
                                            {
                                                while (!recPublications.Eof())
                                                {
                                                    string sPublicationsID = recPublications.GetFieldAsString("pblc_PublicationsID").ToString();
                                                    //AddContent("<BR> pblc_publicationid = " + recPublications.GetFieldAsString("pblc_PublicationsID") + " <BR> sPublicationID =" + sPublicationID);
                                                    Record objBuilderRec = new Record("Planbuilder");
                                                    //objBuilderBox.Fill(objBuilderRec);
                                                    objBuilderRec.SetField("pnbr_plan", iPlanID);
                                                    objBuilderRec.SetField("pnbr_publications", sPublicationsID);
                                                    objBuilderRec.SaveChanges();
                                                    recPublications.GoToNext();

                                                    Record newobjBuilderRec = FindRecord("Planbuilder", "pnbr_Pnbr_planbuilderid =" + objBuilderRec.RecordId.ToString());
                                                    newobjBuilderRec.SetField("pnbr_total", sFinalTotal);
                                                    newobjBuilderRec.SetField("pnbr_standardrate", sStandardRate);
                                                    newobjBuilderRec.SetField("pnbr_days", "," + sDay + ",");
                                                    newobjBuilderRec.SaveChanges();
                                                }
                                            }
                                        }
                                        //AddContent("<BR> dayssss" + "," + sDay + ",");
                                        #endregion
                                    }
                                    
                                    sStandardRate = 0;
                                }
                                #endregion
                            }
                            else
                            {
                                if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_pnbr_planbuilderid")))
                                {
                                    Record recPlanBuilUpdate = FindRecord("planbuilder", "pnbr_pnbr_planbuilderid =" + Dispatch.EitherField("pnbr_pnbr_planbuilderid"));
                                    objBuilderBox.Fill(recPlanBuilUpdate);
                                    recPlanBuilUpdate.SaveChanges();
                                }
                                else
                                {
                                    #region Add Plan Builder
                                    //AddContent("HERE pnbr_publications" + Dispatch.EitherField("pnbr_publications") + "<BR> pnbr_publisher" + Dispatch.EitherField("pnbr_publisher"));
                                    Record objPlanRec2 = FindRecord("Booking", "book_bookingid='" + iPlanID + "' and book_Deleted is null");
                                    objPlanSummaryBox.Fill(objPlanRec2);
                                    AddContent(objPlanSummaryBox);
                                    AddContent("<BR>");
                                    if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_publications")))
                                    {
                                        Record objBuilderRec = new Record("Planbuilder");
                                        objBuilderRec.SetField("pnbr_plan", iPlanID);
                                        objBuilderBox.Fill(objBuilderRec);
                                        objBuilderRec.SaveChanges();
                                    }
                                    else if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_publisher")))
                                    {
                                        Record recPublications = FindRecord("publications", "pblc_PublishersID =" + Dispatch.EitherField("pnbr_publisher"));
                                        if (!recPublications.Eof())
                                        {
                                            while (!recPublications.Eof())
                                            {
                                                string sPublicationsID = recPublications.GetFieldAsString("pblc_PublicationsID").ToString();
                                                //AddContent("<BR> pblc_publicationid = " + recPublications.GetFieldAsString("pblc_PublicationsID") + " <BR> sPublicationID =" + sPublicationID);
                                                Record objBuilderRec = new Record("Planbuilder");
                                                objBuilderBox.Fill(objBuilderRec);
                                                objBuilderRec.SetField("pnbr_plan", iPlanID);
                                                objBuilderRec.SetField("pnbr_publications", sPublicationsID);
                                                objBuilderRec.SaveChanges();
                                                recPublications.GoToNext();
                                            }
                                        }
                                    }
                                    #endregion
                                    
                                }
                            }
                        }
                    }
                    else
                    {
                        AddError("Selected date is from Public Holiday Set");
                    }
                    #endregion

                    #endregion
                    AddContent(HTML.InputHidden("HiddenMode", ""));
                    string sURL = UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&book_bookingid=" + iPlanID + "&OnEdit=Edit";

                    Dispatch.Redirect(sURL);

                }
                else
                {
                    #region Show Validation
                    AddError("Validation Errors - Please correct the highlighted entries");
                    AddContent(HTML.InputHidden("HiddenMode", ""));
                    Record objPlanRec = FindRecord("Booking", "book_bookingid='" + iPlanID + "' and book_Deleted is null");
                    //Record objBuilderRec = FindRecord("Planbuilder", "pnbr_plan='" + iPlanID + "' and pnbr_Deleted is null");
                    Record objBuilderRec = FindRecord("Planbuilder", "pnbr_plan is null and pnbr_Deleted is null");
                    objPlanSummaryBox.Fill(objBuilderRec);
                    objBuilderBox.Fill(objPlanRec);
                    objPlanSummaryBox.GetHtmlInEditMode();
                    objBuilderBox.GetHtmlInEditMode();
                    AddContent(objPlanSummaryBox);
                    AddContent(objBuilderBox);
                    AddSubmitButton("Save", "Save.gif", "javascript:document.EntryForm.HiddenMode.value='save';document.EntryForm.submit();");
                    AddUrlButton("Cancel", "cancel.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&book_bookingid=" + iPlanID);
                    #endregion
                }

            }
            else
            {
                string sAct = "";

                if (!String.IsNullOrEmpty(Dispatch.EitherField("Act")))
                    sAct = Dispatch.EitherField("Act");
                else
                    sAct = "";

                #region Set WorkFlow
                this.AddWorkflowButtons("BOOKING");
                //'this.UseWorkflow = true;
                #endregion

                if (iPlanID > 0)
                {
                    //AddContent("iPlanID =" + iPlanID);
                    if (!String.IsNullOrEmpty(Dispatch.EitherField("Key0")))
                    {
                        string sFlag = "";

                        if (!String.IsNullOrEmpty(Dispatch.EitherField("From")))
                            sFlag = Dispatch.EitherField("From");

                        this.AddWorkflowButtons("BOOKING");

                        if (sFlag == "")
                        {
                            if (Dispatch.EitherField("Key0") == "58")
                                SetContext("Booking", iPlanID);
                        }


                        #region Add Blocks to Booking summary
                        Record objPlanRec = FindRecord("Booking", "book_bookingid='" + iPlanID + "' and book_Deleted is null");

                        Record objBuilderRec = FindRecord("Planbuilder", "pnbr_plan='" + iPlanID + "' and pnbr_Deleted is null");

                        AddContent(HTML.InputHidden("HiddenMode", ""));
                        objPlanSummaryBox.Fill(objPlanRec);
                        objBuilderBox.Fill(objBuilderRec);

                        string sPlanURL = UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&hiddenMode=edit";
                        if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_pnbr_planbuilderid")))
                        {
                            AddContent(objPlanSummaryBox.GetHtmlInEditMode(objPlanRec));
                            AddContent(HTML.InputHidden("HiddenPlanMode", ""));
                            AddContent(HTML.InputHidden("HiddenSelectedDays", ""));
                            AddSubmitButton("Save", "Save.gif", "javascript:document.EntryForm.HiddenMode.value='save';GetPlanBuilderSelectedDays();document.EntryForm.submit();");
                            AddUrlButton("Cancel", "cancel.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&book_bookingid=" + iPlanID + "&From=edit");
                        }
                        else
                        {
                            AddContent(objPlanSummaryBox);
                            AddContent("<BR />");
                            //AddContent(objBuilderBox);                        
                            AddUrlButton("Change", "edit.gif", "javascript:document.EntryForm.HiddenMode.value='edit';document.EntryForm.submit();");
                            AddUrlButton("Delete", "delete.gif", UrlDotNet(ThisDotNetDll, "RunPlanDeletePage") + "&J=Booking Summary&book_bookingid=" + iPlanID);
                        }

                        #endregion
                    }
                }

                if (sAct == "520")
                {
                    AddTopContent(GetCustomEntityTopFrame("Booking"));
                    this.AddWorkflowButtons("BOOKING");

                    #region Add Blocks to Booking summary
                    Record objPlanRec = FindRecord("Booking", "book_bookingid='" + iPlanID + "' and book_Deleted is null");
                    Record objBuilderRec = FindRecord("Planbuilder", "pnbr_plan='" + iPlanID + "' and pnbr_Deleted is null");
                    AddContent(HTML.InputHidden("HiddenMode", ""));
                    objPlanSummaryBox.Fill(objPlanRec);
                    objBuilderBox.Fill(objBuilderRec);
                    string sPlanURL = UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&hiddenMode=edit";
                    AddContent(objPlanSummaryBox);
                    AddContent("<BR />");
                    //AddContent(objBuilderBox);
                    AddUrlButton("Change", "edit.gif", "javascript:document.EntryForm.HiddenMode.value='edit';document.EntryForm.submit();");
                    AddUrlButton("Delete", "delete.gif", UrlDotNet(ThisDotNetDll, "RunPlanDeletePage") + "&J=Booking Summary&book_bookingid=" + iPlanID);
                    #endregion
                }

            }
            #region Code to add Delete and Update Plandetails

            string hpMode = "";
            if (!String.IsNullOrEmpty(Dispatch.EitherField("HiddenPlanMode")))
            {
                hpMode = Dispatch.EitherField("HiddenPlanMode");
            }
            if (hpMode == "UpdateItem")
            {
                if (objCrmHelp.CheckDateIsExists(Dispatch.ContentField("pnbr_date")))
                {
                    if (iPlanID == 0)
                    {
                        iPlanID = Convert.ToInt32(Dispatch.EitherField("pnbr_plan"));
                    }

                    #region Update Plan Builder
                    Record objPlanRec = FindRecord("Booking", "book_bookingid='" + iPlanID + "' and book_Deleted is null");
                    objPlanSummaryBox.Fill(objPlanRec);
                    AddContent(objPlanSummaryBox);
                    AddContent("<BR />");
                    //Dispatch.ContentField("pnbr_publisher");
                    if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_pnbr_planbuilderid")))
                    {
                        Record recPlanBuilder = FindRecord("planbuilder", "pnbr_pnbr_planbuilderid =" + Dispatch.EitherField("pnbr_pnbr_planbuilderid"));
                        objBuilderBox.Fill(recPlanBuilder);
                        recPlanBuilder.SaveChanges();

                        string sURL = UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&book_bookingid=" + iPlanID;
                        Dispatch.Redirect(sURL);
                    }
                    #endregion
                }
                else
                {
                    AddError("Selected date is from Public Holiday Set");
                }
            }
            else if (hpMode == "DeleteItem")
            {
                if (iPlanID == 0)
                {
                    iPlanID = Convert.ToInt32(Dispatch.EitherField("pnbr_plan"));
                }
                if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_pnbr_planbuilderid")))
                {
                    Record recPlanBuilder = FindRecord("planbuilder", "pnbr_pnbr_planbuilderid =" + Dispatch.EitherField("pnbr_pnbr_planbuilderid"));
                    objBuilderBox.Fill(recPlanBuilder);
                    recPlanBuilder.SetField("pnbr_Deleted", "1");
                    recPlanBuilder.SaveChanges();

                    string sURL = UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&book_bookingid=" + iPlanID;
                    Dispatch.Redirect(sURL);
                }
            }
            else if (hpMode == "AddItem")
            {
                #region Planbuilder Item based on days
                //string sSelectedDays = "";
                //string sPublicationID = "";
                //string sStandardSize = "";
                //string scommissiontype = "";
                //string sSections = "";
                //double sStandardRate = 0;
                ////AddContent("My DLL:-" + Dispatch.ContentField("pnbr_date"));
                //if (objCrmHelp.CheckDateIsExists(Dispatch.ContentField("pnbr_date")))
                //{
                //    //AddContent("Days" + Dispatch.ContentField("HiddenSelectedDays"));
                //    if (!String.IsNullOrEmpty(Dispatch.ContentField("HiddenSelectedDays")))
                //    {
                //        sSelectedDays = Dispatch.ContentField("HiddenSelectedDays");
                //    }
                //    if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_publications")))
                //    {
                //        sPublicationID = Dispatch.ContentField("pnbr_publications");
                //    }
                //    if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_standardsize")))
                //    {
                //        sStandardSize = Dispatch.ContentField("pnbr_standardsize");
                //    }
                //    if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_commissiontype")))
                //    {
                //        scommissiontype = Dispatch.ContentField("pnbr_commissiontype");
                //    }
                //    if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_sections")))
                //    {
                //        sSections = Dispatch.ContentField("pnbr_sections");
                //    }
                //    if (sSelectedDays != "")
                //    {
                //        string sDaysQuery = "";
                //        string[] sDays = sSelectedDays.Split(',');
                //        //AddContent("sDays.Length = " + sDays.Length);
                //        if (sDays.Length > 1)
                //        {
                //            double sFinalTotal = 0;
                //            double sDiscount = 0;
                //            double sLoadingValue = 0;
                //            double sPercentLoadingValue = 0;
                //            Record objPlanRec = FindRecord("Booking", "book_bookingid='" + iPlanID + "' and book_Deleted is null");
                //            objPlanSummaryBox.Fill(objPlanRec);
                //            //AddContent(objPlanSummaryBox);
                //            //AddContent("<BR>");
                //            #region Code to add record based on multiple selected Days
                //            foreach (string sDay in sDays)
                //            {
                //                if (sDay == "Mon")
                //                    sDaysQuery = "rate_Monday";

                //                if (sDay == "Tues")
                //                    sDaysQuery = "rate_tuesday";

                //                if (sDay == "Wed")
                //                    sDaysQuery = "rate_wednesday";

                //                if (sDay == "Thur")
                //                    sDaysQuery = "rate_thrusday";

                //                if (sDay == "Fri")
                //                    sDaysQuery = "rate_friday";

                //                if (sDay == "Sat")
                //                    sDaysQuery = "rate_saturday";

                //                if (sDay == "Sun")
                //                    sDaysQuery = "rate_sunday";

                //                string strSQL = "select " + sDaysQuery + " from RatesCard ";
                //                strSQL += " where rate_deleted is null ";
                //                strSQL += " and rate_PublicationsID='" + sPublicationID + "' ";
                //                strSQL += " and rate_standardsize='" + sStandardSize + "' ";
                //                strSQL += " and rate_commissiontype='" + scommissiontype + "' ";
                //                strSQL += " and rate_section='" + sSections + "' ";

                //                QuerySelect objPblcnRec = GetQuery();
                //                objPblcnRec.SQLCommand = strSQL;
                //                objPblcnRec.ExecuteReader();

                //                if (!objPblcnRec.Eof())
                //                {
                //                    sStandardRate = Convert.ToDouble(objPblcnRec.FieldValue(sDaysQuery));
                //                }
                //                //AddContent("<BR>StandardRate = " + sStandardRate + "<BR>SQL =" + strSQL + "<BR> sDay=" + sDay);
                //                if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_loading")))
                //                {
                //                    #region Discount Value
                //                    if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_discount")))
                //                    {
                //                        sDiscount = Convert.ToDouble(Dispatch.ContentField("pnbr_discount"));
                //                    }
                //                    #endregion
                //                    #region Calculation
                //                    if (Dispatch.ContentField("pnbr_loading").ToString() == "Dollar")
                //                    {
                //                        if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_loadingvalue")))
                //                        {
                //                            sLoadingValue = Convert.ToDouble(Dispatch.ContentField("pnbr_loadingvalue"));
                //                        }
                //                        if (sStandardRate > 0 && sLoadingValue > 0 && sDiscount == 0)
                //                        {
                //                            sFinalTotal = Convert.ToDouble(sStandardRate) + Convert.ToDouble(sLoadingValue);
                //                            //AddContent("<br>sFinalTotal = " + sFinalTotal);
                //                        }
                //                        else if (sStandardRate > 0 && sLoadingValue == 0 && sDiscount == 0)
                //                        {
                //                            sFinalTotal = sStandardRate;
                //                        }
                //                        else
                //                        {
                //                            if (sStandardRate == 0 && sLoadingValue > 0 && sDiscount == 0)
                //                                sFinalTotal = Convert.ToDouble(sLoadingValue);
                //                            else if (sStandardRate == 0 && sLoadingValue > 0 && sDiscount > 0)
                //                                sFinalTotal = sLoadingValue - sDiscount;
                //                            else
                //                                sFinalTotal = Convert.ToDouble(sStandardRate) + Convert.ToDouble(sLoadingValue) - Convert.ToDouble(sDiscount);

                //                            //AddContent("<br>sFinalTotal Discount = " + sFinalTotal + " = sLoadingValue =" + sLoadingValue + " = sDiscount =" + sDiscount);
                //                        }
                //                    }
                //                    else if (Dispatch.ContentField("pnbr_loading").ToString() == "Percentage")
                //                    {
                //                        if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_percentvalue")))
                //                        {
                //                            sPercentLoadingValue = Convert.ToDouble(Dispatch.ContentField("pnbr_percentvalue"));
                //                        }
                //                        if (sStandardRate > 0 && sPercentLoadingValue > 0 && sDiscount == 0)
                //                        {
                //                            sFinalTotal = Convert.ToDouble(sStandardRate / 100) * Convert.ToDouble(sPercentLoadingValue);
                //                            sFinalTotal = sFinalTotal + sStandardRate;
                //                        }
                //                        else if (sStandardRate > 0 && sPercentLoadingValue == 0 && sDiscount == 0)
                //                        {
                //                            sFinalTotal = sStandardRate;
                //                        }
                //                        else
                //                        {
                //                            if (sStandardRate == 0 && sPercentLoadingValue > 0 && sDiscount == 0)
                //                                sFinalTotal = 0;
                //                            else if (sStandardRate == 0 && sPercentLoadingValue > 0 && sDiscount > 0)
                //                                sFinalTotal = 0;
                //                            else
                //                                sFinalTotal = Convert.ToDouble(sStandardRate / 100) * Convert.ToDouble(sDiscount);
                //                            sFinalTotal = Convert.ToDouble(sStandardRate) - Convert.ToDouble(sFinalTotal);
                //                            double sLoadingVal = Convert.ToDouble(sStandardRate / 100) * Convert.ToDouble(sPercentLoadingValue);
                //                            sFinalTotal = Convert.ToDouble(sFinalTotal) + Convert.ToDouble(sPercentLoadingValue);
                //                        }
                //                        //AddContent("<br>sFinalTotal Discount = " + sFinalTotal + " = sPercentLoadingValue =" + sPercentLoadingValue + " = sDiscount =" + sDiscount);
                //                    }
                //                    #endregion
                //                }
                //                #region Add Plan Builder
                //                //AddContent("HERE pnbr_publications" + Dispatch.EitherField("pnbr_publications") + "<BR> pnbr_publisher" + Dispatch.EitherField("pnbr_publisher"));

                //                if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_publications")))
                //                {
                //                    Record objBuilderRec = new Record("Planbuilder");
                //                    objBuilderRec.SetField("pnbr_plan", iPlanID);
                //                    objBuilderBox.Fill(objBuilderRec);
                //                    objBuilderRec.SaveChanges();

                //                    Record newobjBuilderRec = FindRecord("Planbuilder", "pnbr_Pnbr_planbuilderid =" + objBuilderRec.RecordId.ToString());
                //                    newobjBuilderRec.SetField("pnbr_total", sFinalTotal);
                //                    newobjBuilderRec.SetField("pnbr_days", "," + sDay + ",");
                //                    newobjBuilderRec.SetField("pnbr_standardrate", sStandardRate); 
                //                    if (Dispatch.ContentField("pnbr_loading").ToString() == "Percentage" && sStandardRate == 0)
                //                        newobjBuilderRec.SetField("pnbr_percentvalue", 0);
                //                    newobjBuilderRec.SaveChanges();
                //                }
                //                else if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_publisher")))
                //                {
                //                    Record recPublications = FindRecord("publications", "pblc_PublishersID =" + Dispatch.EitherField("pnbr_publisher"));
                //                    if (!recPublications.Eof())
                //                    {
                //                        while (!recPublications.Eof())
                //                        {
                //                            string sPublicationsID = recPublications.GetFieldAsString("pblc_PublicationsID").ToString();
                //                            //AddContent("<BR> pblc_publicationid = " + recPublications.GetFieldAsString("pblc_PublicationsID") + " <BR> sPublicationID =" + sPublicationID);
                //                            Record objBuilderRec = new Record("Planbuilder");
                //                            //objBuilderBox.Fill(objBuilderRec);
                //                            objBuilderRec.SetField("pnbr_plan", iPlanID);
                //                            objBuilderRec.SetField("pnbr_publications", sPublicationsID);                                           
                //                            objBuilderRec.SaveChanges();
                //                            recPublications.GoToNext();

                //                            Record newobjBuilderRec = FindRecord("Planbuilder", "pnbr_Pnbr_planbuilderid =" + objBuilderRec.RecordId.ToString());
                //                            newobjBuilderRec.SetField("pnbr_total", sFinalTotal);
                //                            newobjBuilderRec.SetField("pnbr_standardrate", sStandardRate);                                            
                //                            newobjBuilderRec.SetField("pnbr_days", "," + sDay + ",");
                //                            newobjBuilderRec.SaveChanges();
                //                        }
                //                    }
                //                }
                //                //AddContent("<BR> dayssss" + "," + sDay + ",");
                //                #endregion
                //                sStandardRate = 0;
                //            }
                //            base.OnLoad = "javascript:ClearPlanBuilderScreen();";
                //            AddContent(HTML.InputHidden("HiddenMode", ""));
                //            //sPlanID = iPlanID + "," + objBuilderRec.RecordId.ToString();
                //            AddUrlButton("Change", "edit.gif", "javascript:document.EntryForm.HiddenMode.value='edit';document.EntryForm.submit();");
                //            AddUrlButton("Delete", "delete.gif", UrlDotNet(ThisDotNetDll, "RunPlanDeletePage") + "&J=Booking Summary&book_bookingid=" + iPlanID);
                //            string sURL = UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&book_bookingid=" + iPlanID;
                //            Dispatch.Redirect(sURL);
                //            #endregion
                //        }
                //        else
                //        {
                //            //'AddContent("<BR> HERE");
                //            #region Add Plan Builder
                //            //AddContent("HERE pnbr_publications" + Dispatch.EitherField("pnbr_publications") + "<BR> pnbr_publisher" + Dispatch.EitherField("pnbr_publisher"));
                //            Record objPlanRec = FindRecord("Booking", "book_bookingid='" + iPlanID + "' and book_Deleted is null");
                //            objPlanSummaryBox.Fill(objPlanRec);
                //            AddContent(objPlanSummaryBox);
                //            AddContent("<BR>");
                //            if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_publications")))
                //            {
                //                Record objBuilderRec = new Record("Planbuilder");
                //                objBuilderRec.SetField("pnbr_plan", iPlanID);
                //                objBuilderBox.Fill(objBuilderRec);
                //                objBuilderRec.SaveChanges();
                //            }
                //            else if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_publisher")))
                //            {
                //                Record recPublications = FindRecord("publications", "pblc_PublishersID =" + Dispatch.EitherField("pnbr_publisher"));
                //                if (!recPublications.Eof())
                //                {
                //                    while (!recPublications.Eof())
                //                    {
                //                        string sPublicationsID = recPublications.GetFieldAsString("pblc_PublicationsID").ToString();
                //                        //AddContent("<BR> pblc_publicationid = " + recPublications.GetFieldAsString("pblc_PublicationsID") + " <BR> sPublicationID =" + sPublicationID);
                //                        Record objBuilderRec = new Record("Planbuilder");
                //                        objBuilderBox.Fill(objBuilderRec);
                //                        objBuilderRec.SetField("pnbr_plan", iPlanID);
                //                        objBuilderRec.SetField("pnbr_publications", sPublicationsID);
                //                        objBuilderRec.SaveChanges();
                //                        recPublications.GoToNext();
                //                    }
                //                }
                //            }
                //            base.OnLoad = "javascript:ClearPlanBuilderScreen();";
                //            AddContent(HTML.InputHidden("HiddenMode", ""));
                //            //sPlanID = iPlanID + "," + objBuilderRec.RecordId.ToString();
                //            AddUrlButton("Change", "edit.gif", "javascript:document.EntryForm.HiddenMode.value='edit';document.EntryForm.submit();");
                //            AddUrlButton("Delete", "delete.gif", UrlDotNet(ThisDotNetDll, "RunPlanDeletePage") + "&J=Booking Summary&book_bookingid=" + iPlanID);
                //            string sURL = UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&book_bookingid=" + iPlanID;
                //            Dispatch.Redirect(sURL);
                //            #endregion
                //        }
                //    }
                //}
                //else
                //{
                //    AddError("Selected date is from Public Holiday Set");
                //}
                #endregion
            }
            #endregion
            #region Code for Plan builder list population

            string sPlanId = Dispatch.EitherField("pnbr_plan");
            if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_Pnbr_planbuilderid")))
            {
                //AddContent("iPlanID =" + iPlanID) ;
                if (iPlanID == 0)
                {

                    Record objPlanRec = FindRecord("Booking", "book_bookingid='" + sPlanId + "' and book_Deleted is null");
                    objPlanSummaryBox.Fill(objPlanRec);
                    AddContent(objPlanSummaryBox);
                    AddContent("<BR />");
                }
                #region Edit Plan Builder Screen
                objBuilderBox.GetEntry("pnbr_publisher").ReadOnly = true;
                objBuilderBox.GetEntry("pnbr_publications").ReadOnly = true;
                Record recPlanBuilder = FindRecord("Planbuilder", "pnbr_Pnbr_planbuilderid = " + Dispatch.EitherField("pnbr_Pnbr_planbuilderid"));
                objBuilderBox.Fill(recPlanBuilder);
                objBuilderBox.GetHtmlInEditMode(recPlanBuilder);
                //objBuilderBox.GetEntry("pnbr_publisher").ReadOnly = true;     
                AddContent(objBuilderBox);
                //AddContent(HTML.InputHidden("HiddenMode", ""));
                AddContent(HTML.InputHidden("HiddenPlanMode", ""));
                if (iPlanID == 0)
                {
                    iPlanID = Convert.ToInt32(sPlanId); ;
                }
                if (recPlanBuilder.GetFieldAsString("pnbr_loading") == "Percentage")
                {
                    string sValue = recPlanBuilder.GetFieldAsString("pnbr_percentvalue");
                    AddContent(HTML.InputHidden("hdnShowPercentValue", "Show"));
                    AddContent(HTML.InputHidden("hdnShowPercentValues", sValue));
                    //AddContent("iPlanID =" + recPlanBuilder.GetFieldAsString("pnbr_loading"));
                    //objBuilderBox.GetEntry("pnbr_percentvalue").
                }
                else
                {
                    AddContent(HTML.InputHidden("hdnShowPercentValue", ""));
                }
                //string sURLUpdateItem = "javascript:document.EntryForm.HiddenPlanMode.value='UpdateItem';document.EntryForm.submit();";
                //AddUrlButton("Update Plan Details", "save.gif", sURLUpdateItem);
                //string sURLDeleteItem = "javascript:document.EntryForm.HiddenPlanMode.value='DeleteItem';document.EntryForm.submit();";
                string sURLDeleteItem = "javascript:DeletePlanBuilderRecord()";
                AddUrlButton("Delete Plan Item", "delete.gif", sURLDeleteItem);
                //string sCancelURL = UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&book_bookingid=" + iPlanID;
                //Dispatch.Redirect(sCancelURL);
                //AddUrlButton("Cancel", "cancel.gif", sCancelURL);
                #endregion
            }
            else
            {
                if (hMode == "edit")
                {
                    //AddContent("iPlanID " + iPlanID);
                    Record objNewPlanBuilder = new Record("planbuilder");
                    objBuilderBox.Fill(objNewPlanBuilder);
                    objBuilderBox.GetHtmlInEditMode();
                    AddContent(objBuilderBox);                    
                }
            }

            if (iPlanID == 0)
            {

                sPlanID = Dispatch.EitherField("pnbr_plan");
                //AddContent("<BR> sPlanIdsssss =" + sPlanId);
            }
            //AddContent("<BR> iPlanID" + iPlanID + "sPlanID =" + sPlanID);
            List objList = new List("PlanBuilderList");
            if (iPlanID == 0)
            {
                //AddContent("HERE");
                objList.Filter = "pnbr_plan =" + sPlanID + " and pnbr_Deleted is null";
            }
            else
            {
                //AddContent("else planid" + iPlanID);
                objList.Filter = "pnbr_plan = " + iPlanID + " and pnbr_Deleted is null";
            }
            AddContent(objList);
            if (String.IsNullOrEmpty(Dispatch.EitherField("pnbr_Pnbr_planbuilderid")))
            {
                AddContent(HTML.InputHidden("HiddenPlanMode", ""));
                AddContent(HTML.InputHidden("HiddenSelectedDays", ""));
                //string sURLAddItem = "javascript:document.EntryForm.HiddenPlanMode.value='AddItem';GetPlanBuilderSelectedDays();document.EntryForm.submit();";
                //AddUrlButton("Add Plan Details", "save.gif", sURLAddItem);
            }
            GetTabs("Booking", "Booking Summary");
            //AddUrlButton("Copy Plan", "clone.gif", UrlDotNet(ThisDotNetDll, "RunPlanCopyPage"));
            if (hMode != "edit")
            {
                AddUrlButton("Copy Plan Line Items", "clone.gif", UrlDotNet(ThisDotNetDll, "RunPlanNewCopyPage") + "&book_BookingID=" + Dispatch.EitherField("book_BookingID"));

                AddUrlButton("Reactivate Plan", "PrevCircle.gif", UrlDotNet(ThisDotNetDll, "RunReactivatePlanPage"));
            }
            #endregion
        }


    }
}
