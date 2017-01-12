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
    public class AjaxCallPage : Web
    {
        string sCompanyID = "";
        string sFrom = "";
        string sPlanName = "";
        string sOutputPlan = "";
        string sPublicationID = "";
        string sRateCardName = "";
        string sOutputRateCard = "";
        string sRateCardID = "";
        string sSectionName = "";
        string sOutputSectionName = "";
        string sSectionId = "";
        string sSubSectionName = "";
        string sOutputSubSectionName = "";
        string sColorRateCardId = "";
        string sColorCode = "";
        string sColorCaption = "";
        string sAgencyId = "";
        string sClientName = "";
        string sOutputClientName = "";

        string scommissiontype = "";
        string sStandardSize = "";
        string sDays = "";
        string sDaysQuery = "";
        string sSections = "";
        public AjaxCallPage()
            :base()
        {

        }

        public override void  BuildContents()
        {
            if (!String.IsNullOrEmpty(Dispatch.EitherField("sFrom")))
                sFrom = Dispatch.EitherField("sFrom");

            if (!String.IsNullOrEmpty(Dispatch.EitherField("CompanyId")))
                sCompanyID = Dispatch.EitherField("CompanyId");

            if (!String.IsNullOrEmpty(Dispatch.EitherField("PublicationID")))
                sPublicationID = Dispatch.EitherField("PublicationID");

            if (!String.IsNullOrEmpty(Dispatch.EitherField("RateCardId")))
                sRateCardID = Dispatch.EitherField("RateCardId");

            if (!String.IsNullOrEmpty(Dispatch.EitherField("SectionId")))
                sSectionId = Dispatch.EitherField("SectionId");

            if (!String.IsNullOrEmpty(Dispatch.EitherField("RateCard")))
                sColorRateCardId = Dispatch.EitherField("RateCard");

            if (!String.IsNullOrEmpty(Dispatch.EitherField("AgencyID")))
                sAgencyId = Dispatch.EitherField("AgencyID");

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Days")))
                sDays = Dispatch.EitherField("Days");
            
            if (!String.IsNullOrEmpty(Dispatch.EitherField("commissiontype")))
                scommissiontype = Dispatch.EitherField("commissiontype");

            if (!String.IsNullOrEmpty(Dispatch.EitherField("StandardSize")))
                sStandardSize = Dispatch.EitherField("StandardSize");

            if (!String.IsNullOrEmpty(Dispatch.EitherField("sections")))
                sSections = Dispatch.EitherField("sections");

            if (sFrom == "communication")
            {
                if (sCompanyID != "")
                {
                    Record objPlanRec = FindRecord("Booking", "book_Agency = " + sCompanyID + " and book_deleted is null ");

                    if (!objPlanRec.Eof())
                    {
                        while (!objPlanRec.Eof())
                        {
                            sPlanName = objPlanRec.GetFieldAsString("book_name");
                            sOutputPlan += "'" + sPlanName + "'" + ",";
                            objPlanRec.GoToNext();
                        }
                        AddContent("<returnmsg>" + sOutputPlan + "</returnmsg>");
                    }
                    else
                        AddContent("<returnmsg> </returnmsg>");
                }
            }
            else if (sFrom == "RateCard")
            {
                if (sPublicationID != "")
                {
                    Record objPublicationRec = FindRecord("RatesCard", "rate_PublicationsID = " + sPublicationID);
                    if (!objPublicationRec.Eof())
                    {
                        while (!objPublicationRec.Eof())
                        {
                            sRateCardName = objPublicationRec.GetFieldAsString("rate_name");
                            sOutputRateCard += "'" + sRateCardName + "'" + ",";
                            objPublicationRec.GoToNext();
                        }
                        AddContent("<returnmsg>" + sOutputRateCard + "</returnmsg>");
                    }
                    else
                        AddContent("<returnmsg> </returnmsg>");

                }
            }

            //else if (sFrom == "Section")
            //{
            //    string strSQL = "select sctn_name from Sections inner join RatesCard on sctn_Sctn_sectionid=rate_section where rate_PublicationsID = '" + sPublicationID + "'";
            //    QuerySelect objPblcnRec = GetQuery();
            //    objPblcnRec.SQLCommand = strSQL;
            //    objPblcnRec.ExecuteReader();

            //    if (!objPblcnRec.Eof())
            //    {
            //        while (!objPblcnRec.Eof())
            //        {
            //            sSectionName = objPblcnRec.FieldValue("sctn_name").ToString();
            //            sOutputSectionName += "'" + sSectionName + "'" + ",";

            //            objPblcnRec.Next();
            //        }
            //        AddContent("<returnmsg>" + sOutputSectionName + "</returnmsg>");
            //    }
            //    else
            //        AddContent("<returnmsg> </returnmsg>");

            //}

            else if (sFrom == "Section")
            {
                if (sPublicationID != "")
                {
                    Record objSectionRec = FindRecord("sections", "sctn_publicationid = " + sPublicationID);

                    if (!objSectionRec.Eof())
                    {
                        while (!objSectionRec.Eof())
                        {
                            sSectionName = objSectionRec.GetFieldAsString("sctn_name");
                            sOutputSectionName += "'" + sSectionName + "'" + ",";
                            objSectionRec.GoToNext();
                        }
                        AddContent("<returnmsg>" + sOutputSectionName + "</returnmsg>");
                    }
                    else
                        AddContent("<returnmsg> </returnmsg>");

                }
            }

            else if (sFrom == "SS")
            {
                if (sSectionId != "")
                {
                    Record objSubSectionRec = FindRecord("subsection", "suse_section = " + sSectionId);
                    if (!objSubSectionRec.Eof())
                    {
                        while (!objSubSectionRec.Eof())
                        {
                            sSubSectionName = objSubSectionRec.GetFieldAsString("suse_name");
                            sOutputSubSectionName += "'" + sSubSectionName + "'" + ",";
                            objSubSectionRec.GoToNext();
                        }
                        AddContent("<returnmsg>" + sOutputSubSectionName + "</returnmsg>");
                    }
                    else
                        AddContent("<returnmsg> </returnmsg>");
                }
            }

            else if (sFrom == "Color")
            {
                if (sColorRateCardId != "")
                {
                    Record objRatesCardRec = FindRecord("RatesCard", "rate_RatesCardID = " + sColorRateCardId);
                    if (!objRatesCardRec.Eof())
                    {
                        sColorCode = objRatesCardRec.GetFieldAsString("rate_color");
                        if (sColorCode == "" || sColorCode == "null" || sColorCode == "undefined") sColorCode = "";
                        if (sColorCode != "")
                        {
                            AddContent("<returnmsg>" + sColorCode + "</returnmsg>");
                        }
                        else
                            AddContent("<returnmsg> </returnmsg>");
                    }
                }
            }

            else if (sFrom == "Client")
            {
                Record objClientRec = FindRecord("Client", "client_CompanyId=" + sAgencyId);
                if (!objClientRec.Eof())
                {
                    while (!objClientRec.Eof())
                    {
                        sClientName = objClientRec.GetFieldAsString("client_name");
                        sOutputClientName = "'" + sClientName + "'" + ",";
                        objClientRec.GoToNext();
                    }
                    AddContent("<returnmsg>" + sOutputClientName + "</returnmsg>");
                }
                else
                    AddContent("<returnmsg> </returnmsg>");
            }
            else if (sFrom == "DefaultRate")
            {
                if (sDays == "Mon")
                {
                    sDaysQuery = "rate_Monday";
                }
                if (sDays == "Tues")
                {
                    sDaysQuery = "rate_tuesday";
                }
                if (sDays == "Wed")
                {
                    sDaysQuery = "rate_wednesday";
                }
                if (sDays == "Thur")
                {
                    sDaysQuery = "rate_thrusday";
                }
                if (sDays == "Fri")
                {
                    sDaysQuery = "rate_friday";
                }
                if (sDays == "Sat")
                {
                    sDaysQuery = "rate_saturday";
                }
                if (sDays == "Sun")
                {
                    sDaysQuery = "rate_sunday";
                }

                string strSQL = "select " + sDaysQuery + " from RatesCard where rate_PublicationsID='" + sPublicationID + "' and rate_standardsize='" + sStandardSize + "' and rate_commissiontype='" + scommissiontype + "' and rate_section='" + sSections + "'";
                
                QuerySelect objPblcnRec = GetQuery();
                objPblcnRec.SQLCommand = strSQL;
                objPblcnRec.ExecuteReader();

                if (!objPblcnRec.Eof())
                {
                    while (!objPblcnRec.Eof())
                    {
                        sSectionName = objPblcnRec.FieldValue(sDaysQuery).ToString();
                        sOutputSectionName = sSectionName;

                        objPblcnRec.Next();
                    }
                    AddContent("<returnmsg>" + sOutputSectionName + "</returnmsg>");
                }
                else
                    AddContent("<returnmsg> </returnmsg>");

            }
        }
    }
}
