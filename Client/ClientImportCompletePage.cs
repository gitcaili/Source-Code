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
    public class ClientImportCompletePage : Web
    {
        CRMHelper objCRM = new CRMHelper();
        public ClientImportCompletePage()
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

                #region Get Client template block
                if (Fail != "Fail" || Fail == "")
                {
                    
                    objCRM.GetStatusBlock("Client", StatusMsg, "Success");

                    string backURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                    backURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=1650&Mode=1&CLk=T&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data Management";
                    AddUrlButton("Continue", "continue.gif", backURL);

                    string LogUrl = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/NZPAImport/ImportLogs/" + LogFile; 
                    AddUrlButton("Show Logs", "CustMaint.gif", "javascript:OpenLogFile('" + LogUrl + "');");  
                }
                else
                {
                    StatusMsg = "Error while importing Client data. Kindly make sure that you have selected proper Import template.";
                    objCRM.GetStatusBlock("Client", StatusMsg, "Error");

                    string backURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                    backURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=1650&Mode=1&CLk=T&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data Management";
                    AddUrlButton("Continue", "continue.gif", backURL);
                }
                #endregion
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
        }

        #region Get Library Path
        public string GetLibraryPath()
        {
            string Path = "";
            Record RecPath = FindRecord("Custom_SysParams", "parm_name = 'DocStore'");
            Path = RecPath.GetFieldAsString("Parm_Value");
            return Path;
        }
        #endregion
        #region Create Log file

        public string GetlogFile()
        {
            string LibPath = GetLibraryPath();

            string NewPath = LibPath.Replace("\\Library", "");
            NewPath += "WWWRoot\\CustomPages\\NZPAImport\\";
            string sInstallDirName = new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).DirectoryName;
            string Logspath = null;
            string ymd = "";
            try
            {
                string currentPath = NewPath;
                if (!Directory.Exists(Path.Combine(currentPath, "ImportLogs")))
                    Directory.CreateDirectory(Path.Combine(currentPath, "ImportLogs"));

                DateTime theDate = DateTime.Now;
                ymd = theDate.ToString("yyyyMMdd") + "ClientLog.txt";
                Logspath = NewPath + "\\ImportLogs\\" + ymd;
                
            }

            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
            return ymd;
        }


        #endregion
    }
}
