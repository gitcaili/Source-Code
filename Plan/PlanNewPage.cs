using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;
using System.Data;
using System.IO;

namespace NZPACRM.Plan
{
    public class PlanNewPage : Web
    {
        CRMHelper objCrmHelp = new CRMHelper();
        string iPlanId = "";
        string sPlanBuilderId = "";
        public PlanNewPage()
            : base()
        {

        }

        public override void BuildContents()
        {
            EntryGroup objPlanSummaryBox = new EntryGroup("BookingNewEntry", "Booking");
            EntryGroup objBuilderBox = new EntryGroup("PlanBuilderScreen", "Planbuilder");

            #region Add HTML Form so that Navigation will work as expected
            AddContent(HTML.Form());
            #endregion
            if (CurrentUser.HasRights(Sage.PermissionType.Insert, "Booking") == false)
            {
                AddError("Security violation - You must have Permission to view or insert records for this table.");
                return;
            }

            base.OnLoad = "javascript:SetDefaultValue();";
            string hMode = "";
            if (!String.IsNullOrEmpty(Dispatch.EitherField("HiddenMode")))
            {
                hMode = Dispatch.EitherField("HiddenMode");
            }
            if (hMode == "save")
            {
                if (objPlanSummaryBox.Validate() == true)
                {
                    if (objCrmHelp.CheckDateIsExists(Dispatch.ContentField("pnbr_date")))
                    {
                        if (CurrentUser.SessionRead("PlanName") != null)
                        {
                            CurrentUser.SessionWrite("PlanName", "");
                        }

                        #region CreateNewPlan
                        Record objNewPlan = new Record("Booking");
                        objPlanSummaryBox.Fill(objNewPlan);
                        objNewPlan.SetWorkflowInfo("Booking Workflow", "Logged");
                        objNewPlan.SaveChanges();
                        iPlanId = objNewPlan.RecordId.ToString();


                        #endregion

                        #region Set Plan Fields
                        Record objPlanRec = FindRecord("Booking", "Book_bookingid=" + iPlanId);
                        while (!objPlanRec.Eof())
                        {
                            string sAgencyId = "";
                            string sAgencyCode = "";
                            if (!String.IsNullOrEmpty(objPlanRec.GetFieldAsString("book_agency")))
                                sAgencyId = objPlanRec.GetFieldAsString("book_agency");
                            else
                                sAgencyId = "";
                            if (sAgencyId != "")
                            {
                                Record objCompany = FindRecord("Company", "comp_companyid=" + sAgencyId);
                                if (!objCompany.Eof())
                                {
                                    if (!String.IsNullOrEmpty(objCompany.GetFieldAsString("comp_agencycode")))
                                        sAgencyCode = objCompany.GetFieldAsString("comp_agencycode");
                                    else
                                        sAgencyCode = "";
                                }
                            }

                            //'objPlanRec.SetField("book_documentversion", "DOC-001");
                            objPlanRec.SetField("book_agencycode", sAgencyCode);
                            //'objPlanRec.SetField("book_reference", "REF-001");
                            objPlanRec.SetField("book_opened", System.DateTime.Now);
                            objPlanRec.GoToNext();
                        }
                        objPlanRec.SaveChanges();
                        #endregion

                        if (!String.IsNullOrEmpty(Dispatch.ContentField("pnbr_publications")))
                        {
                            #region PlanBuilder
                            /*Record objNewPlanBuilder = new Record("planbuilder");
                            objBuilderBox.Fill(objNewPlanBuilder);
                            objNewPlanBuilder.SaveChanges();
                            sPlanBuilderId = objNewPlanBuilder.RecordId.ToString();
                            string sSQL = " update planbuilder set pnbr_plan=" + iPlanId + " where pnbr_Pnbr_planbuilderid=" + sPlanBuilderId;

                            QuerySelect objPlanBuilderRec = GetQuery();
                            objPlanBuilderRec.SQLCommand = sSQL;
                            objPlanBuilderRec.ExecuteReader();*/
                            #endregion
                            #region PlanBuilder with Days
                            string sSelectedDays = "";
                            string sPublicationID = "";
                            string sStandardSize = "";
                            string scommissiontype = "";
                            string sSections = "";
                            double sStandardRate = 0;
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
                                    //Record objPlanRec1 = FindRecord("Booking", "book_bookingid='" + objPlanRec.RecordId + "' and book_Deleted is null");
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
                                        #region Add Plan Builder
                                        //AddContent("HERE pnbr_publications" + Dispatch.EitherField("pnbr_publications") + "<BR> pnbr_publisher" + Dispatch.EitherField("pnbr_publisher"));

                                        if (!String.IsNullOrEmpty(Dispatch.EitherField("pnbr_publications")))
                                        {
                                            Record objBuilderRec = new Record("Planbuilder");
                                            objBuilderRec.SetField("pnbr_plan", iPlanId);
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
                                                    objBuilderRec.SetField("pnbr_plan", iPlanId);
                                                    objBuilderRec.SetField("pnbr_publications", sPublicationsID);
                                                    objBuilderRec.SaveChanges();
                                                    recPublications.GoToNext();

                                                    Record newobjBuilderRec = FindRecord("Planbuilder", "pnbr_Pnbr_planbuilderid =" + recPublications.RecordId.ToString());
                                                    newobjBuilderRec.SetField("pnbr_total", sFinalTotal);
                                                    newobjBuilderRec.SetField("pnbr_standardrate", sStandardRate);
                                                    newobjBuilderRec.SetField("pnbr_days", "," + sDay + ",");
                                                    newobjBuilderRec.SaveChanges();
                                                }
                                            }
                                        }
                                        //AddContent("<BR> dayssss" + "," + sDay + ",");
                                        #endregion
                                        sStandardRate = 0;
                                    }
                                    #endregion
                                }

                            }
                            #endregion

                        }
                        string sURL = UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&Key37=" + iPlanId + "&J=Booking Summary&book_bookingid=" + iPlanId + "&From=New";
                        Dispatch.Redirect(sURL);
                    }
                    else
                    {
                        AddError("Selected date is from Public Holiday Set");

                        AddContent(HTML.InputHidden("HiddenMode", ""));
                        Record objNewPlan = new Record("Booking");
                        Record objNewPlanBuilder = new Record("planbuilder");
                        objPlanSummaryBox.Fill(objNewPlan);
                        objBuilderBox.Fill(objNewPlanBuilder);
                        objPlanSummaryBox.GetHtmlInEditMode();
                        objBuilderBox.GetHtmlInEditMode();
                        AddContent(HTML.InputHidden("HiddenPlanMode", ""));
                        AddContent(HTML.InputHidden("HiddenSelectedDays", ""));
                        AddUrlButton("Save", "save.gif", "javascript:document.EntryForm.HiddenMode.value='save';GetPlanBuilderSelectedDays();document.EntryForm.submit();");
                        AddUrlButton("Cancel", "cancel.gif", UrlDotNet(ThisDotNetDll, "RunPlanDedupePage"));
                    }
                }
                else
                {
                    #region Show Validation
                    AddError("Validation Errors - Please correct the highlighted entries");

                    AddContent(HTML.InputHidden("HiddenMode", ""));
                    Record objNewPlan = new Record("Booking");
                    Record objNewPlanBuilder = new Record("planbuilder");
                    objPlanSummaryBox.Fill(objNewPlan);
                    objBuilderBox.Fill(objNewPlanBuilder);
                    objPlanSummaryBox.GetHtmlInEditMode();
                    objBuilderBox.GetHtmlInEditMode();
                    AddContent(HTML.InputHidden("HiddenPlanMode", ""));
                    AddContent(HTML.InputHidden("HiddenSelectedDays", ""));
                    AddUrlButton("Save", "save.gif", "javascript:document.EntryForm.HiddenMode.value='save';GetPlanBuilderSelectedDays();document.EntryForm.submit();");
                    AddUrlButton("Cancel", "cancel.gif", UrlDotNet(ThisDotNetDll, "RunPlanDedupePage"));
                    #endregion
                }
            }
            else
            {
                string sCancelURL = "";

                if (Dispatch.EitherField("From") == "ignore")
                {
                    if (!String.IsNullOrEmpty(Dispatch.EitherField("PlanName")))
                        sCancelURL = UrlDotNet(ThisDotNetDll, "RunPlanConflictPage") + "&J=Booking Summary&PlanName=dedupe";
                }

                else
                {
                    sCancelURL = UrlDotNet(ThisDotNetDll, "RunPlanDedupePage");
                }
                AddContent(HTML.InputHidden("HiddenMode", ""));

                if (!String.IsNullOrEmpty(Dispatch.EitherField("PlanName")))
                {
                    if (CurrentUser.SessionRead("PlanName") != null)
                    {
                        AddContent(HTML.InputHidden("HiddenName", CurrentUser.SessionRead("PlanName").ToString()));
                    }
                    else
                    {
                        AddContent(HTML.InputHidden("HiddenName", ""));
                    }
                }
                else
                {
                    AddContent(HTML.InputHidden("HiddenName", ""));
                }

                #region Building Block
                Record objNewBuilding = new Record("Booking");
                Record objNewPlanBuilder = new Record("planbuilder");
                objPlanSummaryBox.Fill(objNewBuilding);
                objBuilderBox.Fill(objNewPlanBuilder);
                objPlanSummaryBox.GetHtmlInEditMode();
                objBuilderBox.GetHtmlInEditMode();
                #endregion

                #region Add Buttons
                AddContent(HTML.InputHidden("HiddenPlanMode", ""));
                AddContent(HTML.InputHidden("HiddenSelectedDays", ""));

                AddUrlButton("Save", "save.gif", "javascript:document.EntryForm.HiddenMode.value='save';GetPlanBuilderSelectedDays();document.EntryForm.submit();");
                AddUrlButton("Cancel", "cancel.gif", sCancelURL);
                #endregion
            }
            AddContent(objPlanSummaryBox);
            AddContent("<BR>");
            AddContent(objBuilderBox);
        }
    }
}
