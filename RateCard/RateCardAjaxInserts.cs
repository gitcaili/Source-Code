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
    class RateCardAjaxInserts:Web
    {
        string sPublicationID = "";
        string sDays = "";
        string sStandardsize = "";
        //string sSize = "";
        string sCommissionType = "";
        string sSections = "";
        //string sSubsection = "";
        string sRateCardId = "";
        string sRetRateCardId = "";
        string sInserts = "";
        //string sCustomClass = "";
        //string sCustomType = "";
      //  string sColor = "";
      //  string sClientID = "";

        public RateCardAjaxInserts() { }
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

           
            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_days")))
            {
                sDays = Dispatch.EitherField("Pnbr_days");
            }

           

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_standardsize")))
            {
                sStandardsize = Dispatch.EitherField("Pnbr_standardsize");
            }

          
            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_commissiontype")))
            {
                sCommissionType = Dispatch.EitherField("Pnbr_commissiontype");
            }



            string sectionsql = "Select * from Sections where sctn_sctn_sectionid ='" + sSections + "'";
            QuerySelect sectionquery = GetQuery();
            sectionquery.SQLCommand = sectionsql;
            sectionquery.ExecuteReader();
            string sectionname = sectionquery.FieldValue("sctn_name");

            if (sectionname.Contains("Inserts"))
            {


                #endregion
                #region Filter Methods
                string sSQl = "";
                sSQl = "select rate_Name,rate_RatesCardID from RatesCard where rate_Deleted is null ";
                if (sPublicationID != "")
                {
                    sSQl += " and rate_PublicationsID = '" + sPublicationID + "' ";
                }
                if (sSections != "")
                {
                    sSQl += " and rate_section = '" + sSections + "' ";
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
            }
            else
            {
                AddContent("<returnmsg>D.O.D</returnmsg>");
            }
            #endregion

        }

    }
}
