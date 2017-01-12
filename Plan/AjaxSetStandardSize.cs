using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;
using System.Data;
using System.IO;
using Sage.CRM.Utils;

namespace NZPACRM.Plan
{
    public class AjaxSetStandardSize:Web
    {
        string standardSize = "";
        string sPublicationID = "";
        public AjaxSetStandardSize()
        {

        }
        public override void BuildContents()
        {
            if (!String.IsNullOrEmpty(Dispatch.EitherField("publicationID")))
            {
                sPublicationID = Dispatch.EitherField("publicationID");
            }
            else
            {
                sPublicationID = "";
            }
            if (sPublicationID == "")
            {
                //   string strSQLstring = "select Capt_US from Custom_Captions where Capt_Family = 'rate_standardsize' and Capt_Deleted is null";
                string strSQLstring = "select Capt_US from Custom_Captions where Capt_Family = 'rate_standardsize' and Capt_Deleted is null";
                QuerySelect objClientRec = GetQuery();
                objClientRec.SQLCommand = strSQLstring;
                objClientRec.ExecuteReader();
                if (!objClientRec.Eof())
                {
                    while (!objClientRec.Eof())
                    {
                        standardSize += objClientRec.FieldValue("Capt_US") + ",";
                        //AddContent(HTML.InputHidden("hdnSetStandardSize", standardSize));
                        objClientRec.Next();
                    }

                    AddContent("<returnmsg>" + standardSize + "</returnmsg>");
                }
            }
            else if (sPublicationID != "")
            {
                string strSQLstring = "select distinct rate_standardsize from RatesCard where rate_PublicationsID=" + sPublicationID + " and rate_Deleted is null";
                //string strSQLstring  = "select top 5 Capt_US from Custom_Captions where Capt_Family = 'rate_standardsize' and Capt_Deleted is null";
                QuerySelect objClientRec = GetQuery();
                objClientRec.SQLCommand = strSQLstring;
                objClientRec.ExecuteReader();
                if (!objClientRec.Eof())
                {
                    while (!objClientRec.Eof())
                    {
                        standardSize += objClientRec.FieldValue("rate_standardsize") + ",";
                        //AddContent(HTML.InputHidden("hdnSetStandardSize", standardSize));
                        objClientRec.Next();
                    }

                    AddContent("<returnmsg>" + standardSize + "</returnmsg>");
                }
            }
        }
    }
}
