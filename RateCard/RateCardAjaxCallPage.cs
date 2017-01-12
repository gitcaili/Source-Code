using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;
using System.Data;
using System.IO;


namespace NZPACRM.RateCard
{    
    class RateCardAjaxCallPage : Web
    {
        string sPublicationID = "";
        string sDays = "";
        string sStandardsize = "";
        string sSize = "";
        string sCommissionType = "";
        string sSections = "";
        string sSubsection = "";
        string sRateCardId = "";
        string sRetRateCardId = "";
        public RateCardAjaxCallPage()
        { 
        
        }
        public override void BuildContents()
        {
            #region Ajaxcall to filter recotrd 
            if (!String.IsNullOrEmpty(Dispatch.EitherField("PublicationID")))
            {
                sPublicationID = Dispatch.EitherField("PublicationID");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_sections")))
            {
                sSections = Dispatch.EitherField("Pnbr_sections");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_subsection")))
            {
                sSubsection = Dispatch.EitherField("Pnbr_subsection");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_days")))
            {
                sDays = Dispatch.EitherField("Pnbr_days");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_size")))
            {
                sSize = Dispatch.EitherField("Pnbr_size");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_standardsize")))
            {
                sStandardsize = Dispatch.EitherField("Pnbr_standardsize");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_commissiontype")))
            {
                sCommissionType = Dispatch.EitherField("Pnbr_commissiontype");
            }           
            #endregion
            #region Filter Methods
            string sSQl = "";
            sSQl = "select rate_Name,rate_RatesCardID from RatesCard where rate_Deleted is null";
            if (sPublicationID != "")
            {
                sSQl += " and rate_PublicationsID = '" + sPublicationID + "' ";
            }
            if (sSections != "")
            {
                sSQl += " and rate_section = '" + sSections + "' ";
            }
            //if (sSubsection != "")
            //{
            //    sSQl += " and rate_subsectionid = " + sSubsection;
            //}
            if (sDays != "")
            {
                sSQl += " and rate_Day = " + sDays;
            }
            if (sSize != "")
            {
                sSQl += " and rate_Size = '" + sSize + "' ";
            }
            if (sStandardsize != "")
            {
                sSQl += " and rate_standardsize = '" + sStandardsize + "' ";
            }
            if (sCommissionType != "")
            {
                sSQl += " and rate_commissiontype = '" + sCommissionType + "' ";
            }
            //AddContent(" sSQl <BR>" + sSQl);
            //return;
            QuerySelect objRateCardRec = GetQuery();
            objRateCardRec.SQLCommand = sSQl;
            objRateCardRec.ExecuteReader();
            if (!objRateCardRec.Eof())
            {
                while (!objRateCardRec.Eof())
                {
                    sRateCardId = objRateCardRec.FieldValue("rate_RatesCardID");
                    sRetRateCardId += sRateCardId + ",";
                    objRateCardRec.Next();
                }
                AddContent("<returnmsg>" + sRetRateCardId + "</returnmsg>");
            }
            #endregion
        }
    }
}
