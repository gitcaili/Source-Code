using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using System.Data;
using NZPACRM.Common;
using System.IO;

namespace NZPACRM
{
    public class PublishersImportCompletePage : Web
    {
        CRMHelper objCRM = new CRMHelper();
        public PublishersImportCompletePage()
        {
           
        }
        public override void BuildContents()
        {
            try
            {
                string LogFile = "";
                string StatusMsg = "";
                int ImportedRecord = 0;

                string Fail = "";
                if (!String.IsNullOrEmpty(Dispatch.EitherField("LogFileName")))
                {
                    LogFile = Dispatch.EitherField("LogFileName");
                }
                
                if (!String.IsNullOrEmpty(Dispatch.EitherField("imported")))
                {
                    ImportedRecord = Convert.ToInt32(Dispatch.EitherField("imported"));
                    StatusMsg += "Total " + ImportedRecord + " Records are imported in Sage CRM." + "<br />";
                }
                if (!String.IsNullOrEmpty(Dispatch.EitherField("Fail")))
                {
                    Fail = Dispatch.EitherField("Fail");
                }
                if (Fail != "Fail" || Fail == "")
                {
                    #region Get Client template block
                    objCRM.GetStatusBlock("Publishers", StatusMsg, "Success");
                    #endregion
                    string backURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                    backURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=1650&Mode=1&CLk=T&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data Management";
                    
                    AddUrlButton("Continue", "continue.gif", backURL);
                    
                    string LogUrl = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/NZPAImport/ImportLogs/" + LogFile; 
                    AddUrlButton("Show Logs", "CustMaint.gif", "javascript:OpenLogFile('" + LogUrl + "')");
                }
                else
                {
                    StatusMsg = "Error while importing Publishers data. Kindly make sure that you have selected proper Import template.";
                    #region Get Client template block
                    objCRM.GetStatusBlock("Publishers", StatusMsg, "Error");
                    #endregion
                    
                    string backURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                    backURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=1650&Mode=1&CLk=T&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data Management";
                    AddUrlButton("Continue", "continue.gif", backURL);
                }
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
        }
    }
}
