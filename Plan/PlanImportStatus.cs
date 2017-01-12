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
    public class PlanImportStatus:Web
    {
         CRMHelper objCRM = new CRMHelper();
        string shttpURL = "";

        public PlanImportStatus()
        {
            #region get Http from url
            try
            {
                string s = Dispatch.ServerVariable("HTTP_REFERER");
                char[] cSplit = { '/' };
                string[] sHTTP = s.Split(cSplit);

                if (!String.IsNullOrEmpty(sHTTP[0]))
                    shttpURL = sHTTP[0];

                if (CurrentUser.SessionRead("HTTP_REFERER") == null)
                {
                    CurrentUser.SessionWrite("HTTP_REFERER", shttpURL);
                }
            }
            catch (Exception ex)
            {
                if (CurrentUser.SessionRead("HTTP_REFERER") != null)
                {
                    shttpURL = CurrentUser.SessionRead("HTTP_REFERER").ToString();
                }
            }
        }

        public override void BuildContents()
        {   
            try
            {
                int iSucessCount = 0;
                int iFailedCount = 0;
                string sStatusMsg = "";
                string sLogPath = "";
                string isValidColumn = "";
                string sRowCount = "";
                string sAllDuplicate = "";
                
                if (!String.IsNullOrEmpty(Dispatch.EitherField("ValidColumn")))
                {
                    AddContent(Dispatch.EitherField("ValidColumn"));
                        
                    isValidColumn = "ROW";
                    sStatusMsg = "An Error Occured while reading the Column Names of the excel sheet.Please verify the column name before importing";
                }
                
                if (!String.IsNullOrEmpty(Dispatch.EitherField("hasrows")))
                {
                    sRowCount = Dispatch.EitherField("hasrows");
                }

                if (!String.IsNullOrEmpty(Dispatch.EitherField("AllDup")))
                {
                    sAllDuplicate = Dispatch.EitherField("AllDup");
                }
                
                if (sRowCount == "Y")
                {
                    if (sAllDuplicate == "E")
                    {
                        sStatusMsg += " All Record of excel sheet exist in Sage CRM." + "<br />";
                    }
                    else
                    {
                        if (iSucessCount != 0)
                        {
                            sStatusMsg += " Publication doesn't exist in Sage CRM." + "<br />";
                        }
                    }
                    if (isValidColumn == "")
                    {
                        if (!String.IsNullOrEmpty(Dispatch.EitherField("inserted")))
                        {
                            iSucessCount = Convert.ToInt32(Dispatch.EitherField("inserted"));
                        }

                        if (!String.IsNullOrEmpty(Dispatch.EitherField("Failed")))
                        {
                            iFailedCount = Convert.ToInt32(Dispatch.EitherField("Failed"));

                        }

                        if (iSucessCount != 0 || iFailedCount != 0)
                        {
                            sStatusMsg += iSucessCount + " Plan record(s) imported in Sage CRM. For more details, click on Show Logs button." + "<br />";
                        }

                        if (!String.IsNullOrEmpty(Dispatch.EitherField("LogPath")))
                        {
                            sLogPath = Dispatch.EitherField("LogPath");
                        }
                    }
                }
                else if (sRowCount == "N")
                {
                    sStatusMsg = "Excel sheet has no records.";
                }
                
                #region Get Client template block
                objCRM.GetRateCardStatusBlock("Plan Import", sStatusMsg, isValidColumn, sRowCount);
                #endregion

                string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=1650&Mode=1&CLk=T&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data Management";
                AddUrlButton("Continue", "continue.gif", sURL);

                if (isValidColumn == "" && sRowCount == "Y")
                {
                    string LogUrl = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/NZPAImport/LogFiles/" + sLogPath; //"http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/" + LogFileName;
                    //AddContent(LogUrl);
                    AddUrlButton("Show Logs", "CustMaint.gif", "javascript:OpenLogFile('" + LogUrl + "')");
                   // AddUrlButton("Show Log", "CustMaint.gif", "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/do/" + LogUrl + "?SID=" + Dispatch.EitherField("SID") + "&Act=1282&Mode=1&FileName=\\" + LogUrl);
                }

            }
            catch (Exception Ex)
            {
                this.AddError(Ex.Message);
            }
            #endregion
        }
    }
}
