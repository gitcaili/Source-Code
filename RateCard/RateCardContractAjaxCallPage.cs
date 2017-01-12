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
    class RateCardContractAjaxCallPage:Web
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
        string sCustomClass = "";
        string sCustomType = "";
        string sColor = "";
        string sClientID = "";

        public RateCardContractAjaxCallPage()
        { }
        public override void BuildContents() { 
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

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_CustomType")))
            {
                sCustomClass = Dispatch.EitherField("Pnbr_CustomType");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_classtype")))
            {
                sCustomType = Dispatch.EitherField("Pnbr_classtype");
            }
            if (!String.IsNullOrEmpty(Dispatch.EitherField("client_Clientid")))
            {
                sClientID = Dispatch.EitherField("client_Clientid");
            }



            #endregion
            #region Filter Methods
            string sSQl = "";
            bool mod = false;

            if (sSize == "Custom" && sCustomClass == "Display" ) mod = true;

            sSQl = "select rate_Name,rate_RatesCardID from RatesCard where rate_Deleted is null ";
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
            //if (sDays != "")
            //{
            //    sSQl += " and rate_Day = " + sDays;
            ////}
         
            //{
                sSQl += " and rate_Size = '" + sSize + "' ";
            //}
            if (mod){
                sSQl += " and (rate_standardsize = 'Column cm rate' or rate_standardsize = 'Module rate') ";
            }
           else if (sStandardsize != "")
            {
                sSQl += " and rate_standardsize = '" + sStandardsize + "' ";
            }
            if (sCommissionType != "")
            {
                sSQl += " and rate_commissiontype = '" + sCommissionType + "' ";
            }
            if (sCustomClass != "")
            {
                sSQl += " and rate_customtype = '" + sCustomClass + "' ";
            }
            if (sCustomType != "")
            {
                sSQl += " and rate_customsheet = '" + sCustomType + "' ";
            }
            if (sClientID != ""){
                sSQl += "and rate_client = '" + sClientID + "'";
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
                    break; /// this is rubberband for Franco 
                }
                AddContent("<returnmsg>" + sRetRateCardId + "</returnmsg>");
            }
            else
            {
                sSQl = "";
                sSQl = "select rate_Name,rate_RatesCardID from RatesCard where rate_Deleted is null ";
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
                //if (sDays != "")
                //{
                //    sSQl += " and rate_Day = " + sDays;
                //}
                //if (sSize != "")
                //{
                //    sSQl += " and rate_Size = '" + sSize + "' ";
                //}
                if (mod)
                {
                    sSQl += " and (rate_standardsize = 'Column cm rate' or rate_standardsize = 'Module rate') ";
                }
                else if (sStandardsize != "")
                {
                    sSQl += " and rate_standardsize = '" + sStandardsize + "' ";
                }
                if (sCommissionType != "")
                {
                    sSQl += " and rate_commissiontype = '" + sCommissionType + "' ";
                }
                if (sCustomClass != "")
                {
                    sSQl += " and rate_customtype = '" + sCustomClass + "' ";
                }
                if (sCustomType != "")
                {
                    sSQl += " and rate_customsheet = '" + sCustomType + "' ";
                }
               
                //AddContent(" sSQl <BR>" + sSQl);
                //return;
                objRateCardRec = GetQuery();
                objRateCardRec.SQLCommand = sSQl;
                objRateCardRec.ExecuteReader();
                if (!objRateCardRec.Eof())
                {
                    while (!objRateCardRec.Eof())
                    {
                        sRateCardId = objRateCardRec.FieldValue("rate_RatesCardID");
                        sRetRateCardId += sRateCardId + ",";
                        objRateCardRec.Next();
                        break; /// this is rubberband for Franco 
                    }
                    AddContent("<returnmsg>" + sRetRateCardId + "</returnmsg>");
                }
                else
                {
                    AddContent("<returnmsg>D.O.D</returnmsg>"); 
                }
            }
              
            #endregion
        
    }
    }
}
