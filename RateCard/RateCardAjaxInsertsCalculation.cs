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
    class RateCardAjaxInsertsCalculation : Web
    {

        public RateCardAjaxInsertsCalculation() { }
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

            if (!String.IsNullOrEmpty(Dispatch.EitherField("Pnbr_inserts")))
            {
                sInserts = Dispatch.EitherField("Pnbr_inserts");
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
                sSQl = "select rate_Name,rate_RatesCardID,rate_height,rate_width,rate_" + dayswap(sDays) + " from RatesCard where rate_Deleted is null ";
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
                        sRateCardId = objRateCardRec.FieldValue("rate_" + dayswap(sDays));
                        string min = objRateCardRec.FieldValue("rate_height").ToString();
                        string max = objRateCardRec.FieldValue("rate_width").ToString();

                        double imin = Double.Parse(min);
                        double imax = Double.Parse(max);
                        double ivalue = Double.Parse(sInserts);

                        if (imin <= ivalue && ivalue <= imax)
                        {
                            double rate = double.Parse(sRateCardId);
                            double total = rate * ivalue;

                           sRetRateCardId += sRateCardId + "," + total;
                       }
                      //  objRateCardRec.Next();
                        break; /// this is rubberband for Franco 
                    }
                    AddContent("<returnmsg>" + sRetRateCardId + "</returnmsg>");
                }


                else
                {
                    AddContent("<returnmsg>" +"D.O.D" + "</returnmsg>");

                }
            }
            else
            {
                AddContent("<returnmsg>" + "D.O.D" + "</returnmsg>");

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
