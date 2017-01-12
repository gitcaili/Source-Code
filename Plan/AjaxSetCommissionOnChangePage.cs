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
    class AjaxSetCommissionOnChangePage : Web
    {
        string sBookingId = "";
        string sClienType = "";
        string sClientId = "";
        string strCommType = "";
        public AjaxSetCommissionOnChangePage()
        {
        
        }

        public override void BuildContents()
        {
            try
            {               
                #region Adding Html Form
                AddContent(HTML.Form());
                #endregion

                #region Booking Id from QueryString     
                if (!String.IsNullOrEmpty(Dispatch.EitherField("client_ClientID")))
                {
                    sClientId = Dispatch.EitherField("client_ClientID");
                }
              
                if (sClientId != "")
                {
                    string strSQL = "Select REPLACE(client_type,',','') as client_type from Client where client_ClientID='" + sClientId + "' and client_Deleted is null";
                    QuerySelect objBookRec = GetQuery();
                    objBookRec.SQLCommand = strSQL;
                    objBookRec.ExecuteReader();
                    if (!objBookRec.Eof())
                    {
                        sClienType = objBookRec.FieldValue("client_type");
                        //AddContent("sClienType =" + sClienType + "a " + strSQL);
                        //return;
                        if (sClienType != "")
                        {
                            string strSQLstring = "select Capt_US from Custom_Captions where Capt_Code in ('" + sClienType + "') and Capt_Family = 'client_type' and Capt_Deleted is null";
                            QuerySelect objClientRec = GetQuery();
                            objClientRec.SQLCommand = strSQLstring;
                            objClientRec.ExecuteReader();
                            if (!objClientRec.Eof())
                            {
                                strCommType = objClientRec.FieldValue("Capt_US");
                                AddContent(HTML.InputHidden("hdnClientCommissionType", strCommType));
                                AddContent("<returnmsg>" + strCommType + "</returnmsg>");
                            }
                        }
                    }
                }
                #endregion
            }
            catch (Exception ex)
            { 
            
            }
        }
    }
}
