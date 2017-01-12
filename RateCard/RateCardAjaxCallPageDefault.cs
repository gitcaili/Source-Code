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
    class RateCardAjaxCallPageDefault : Web
    {
        string sPublicationID = "";
        //string sDay = "";

        string sDays = "";
        string sStandardsize = "";
        string sSize = "";
        string sCommissionType = "";
        string sSections = "";
        string sSubsection = "";
        string sRateCardId = "";
        string sRetRateCardId = "";
        string sCustomClass="";
        string sCustomType ="";
        string sColor = "";
        string sClinet = "";
        public RateCardAjaxCallPageDefault()
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
            if(!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_color")))
            {
                sColor = Dispatch.EitherField("Pnbr_color");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_size")))
            {
                sSize = Dispatch.EitherField("Pnbr_size");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_standardsize")))
            {
                sStandardsize = Dispatch.EitherField("Pnbr_standardsize");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_days")))
            {
                sDays = Dispatch.EitherField("Pnbr_days");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_commissiontype")))
            {
                sCommissionType = Dispatch.EitherField("Pnbr_commissiontype");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_custom")))
            {
                sCustomClass = Dispatch.EitherField("Pnbr_custom");
            }

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_classtype")))
            {
                sCustomType = Dispatch.EitherField("Pnbr_classtype");
            }
            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_client")))
            {
                sClinet = Dispatch.EitherField("Pnbr_client");
            }


            #endregion
            bool mod = false;
            if (sSize == "Custom" && sCustomClass == "Display") mod = true;


            string sSQl = "";
            sSQl = "select rate_Name,rate_RatesCardID,rate_" +dayswap(sDays)+" from RatesCard where rate_Deleted is null";
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
           
            if (sSize != "" && !mod)
            {
                sSQl += " and rate_Size = '" + sSize + "' ";
            }
            if (sClinet != "")
            {
                sSQl += "and  rate_client = '" + sClinet + "'";
            }
            if (sDays != "")
            {
                sSQl += " and rate_Day like '%" + sDays + "%'";
            }
            if (sStandardsize != "")
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
            if (sColor != "")
            {
                sSQl += " and rate_color = '" + sColor + "'";
            }
            if (mod)
            {
                sSQl += "and rate_size = 'Standard' and (rate_standardsize = 'Columncmrate' or rate_standardsize = 'ColumnCmRate' or rate_standardsize ='Modulerate' or rate_standardsize ='ModuleRate')";
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
                    sRateCardId = objRateCardRec.FieldValue("rate_" + dayswap(sDays));
                    sRetRateCardId += sRateCardId + ",";
                    objRateCardRec.Next();
                }
                AddContent("<returnmsg>" + sRetRateCardId + "</returnmsg>");
            }
            else
            {
                #region Filter Methods
                sSQl = "";
                sSQl = "select rate_Name,rate_RatesCardID,rate_" + dayswap(sDays) + " from RatesCard where rate_Deleted is null";
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

                if (sSize != "" && !mod)
                {
                    sSQl += " and rate_Size = '" + sSize + "' ";
                }
                
                if (sDays != "")
                {
                    sSQl += " and rate_Day like '%" + sDays + "%'";
                }
                if (sStandardsize != "")
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
                if (sColor != "")
                {
                    sSQl += " and rate_color = '" + sColor + "'";
                }
                if (mod)
                {
                    sSQl += "and rate_size = 'Standard' and (rate_standardsize = 'Columncmrate' or rate_standardsize = 'ColumnCmRate' or rate_standardsize ='Modulerate' or rate_standardsize ='ModuleRate')";
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
                        sRateCardId = objRateCardRec.FieldValue("rate_" + dayswap(sDays));
                        sRetRateCardId += sRateCardId + ",";
                        objRateCardRec.Next();
                    }
                    AddContent("<returnmsg>" + sRetRateCardId + "</returnmsg>");
                }
                else AddContent(sSQl);
            }
               
            #endregion
        }

        public string dayswap(string day)
        {
            if (day.Equals("Mon")) return "Monday";
            if (day.Equals("Tues")) return "tuesday";
            if (day.Equals("Wed")) return "wednesday";
            if (day.Equals("Thur")) return "thrusday";
            if (day.Equals("Fri")) return "friday";
            if (day.Equals("Sat")) return "saturday";
            return "sunday";
        }
    }

}
