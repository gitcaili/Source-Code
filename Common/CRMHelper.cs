using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using System.Data;
using System.Data.OleDb;

namespace NZPACRM.Common
{
    public class CRMHelper : Web
    {        
        public string smtpusername = "";
        public string smtppwd = "";
        public string servername = "";
        public string smtpport = "";        

        public void SetTabs(string EntityName)
        {
            GetTabs(EntityName);
        }
        public void SetCustomEntityTopFrame(string EntityName)
        {
            AddTopContent(GetCustomEntityTopFrame(EntityName));
        }
        public override void BuildContents()
        {
            //throw new NotImplementedException();
        }
        public void AddressBox(string EntityName)
        {
            //base.OnLoad = "javascript:SetTabIndex();";
            string sHTML = "";

            sHTML += HTML.Form();

            sHTML += HTML.StartTable();

            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData("Address 1:", "VIEWBOXCAPTION");            
            sHTML += HTML.TableData("Address 2:", "VIEWBOXCAPTION");
            sHTML += HTML.TableData("") + HTML.TableData("Type", "VIEWBOXCAPTION");

            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData(HTML.InputText("addr_address1", "", 40, 20,"","",false,"tabindex=1")+"<font style='color:blue;'>*</font>");
            sHTML += HTML.TableData(HTML.InputText("addr_address2", "", 40, 40, "", "", false, "tabindex=2"));
            sHTML += HTML.TableData("Business", "VIEWBOXCAPTION") + HTML.TableData(HTML.InputCheckBox("addr_business", false, "", "", false, "", "tabindex=10"), "VIEWBOXCAPTION");

            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData("Address 3:", "VIEWBOXCAPTION");
            sHTML += HTML.TableData("Address 4:", "VIEWBOXCAPTION");
            sHTML += HTML.TableData("Billing", "VIEWBOXCAPTION") + HTML.TableData(HTML.InputCheckBox("add_billing", false, "", "", false, "", "tabindex=11"), "VIEWBOXCAPTION");

            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData(HTML.InputText("addr_address3", "", 40, 40, "", "", false, "tabindex=3"));
            sHTML += HTML.TableData(HTML.InputText("addr_address4", "", 40, 40, "", "", false, "tabindex=4"));
            sHTML += HTML.TableData("Shipping", "VIEWBOXCAPTION") + HTML.TableData(HTML.InputCheckBox("addr_shipping", false, "", "", false, "", "tabindex=12"), "VIEWBOXCAPTION");

            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData("City:", "VIEWBOXCAPTION");
            sHTML += HTML.TableData("State:", "VIEWBOXCAPTION");

            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData(HTML.InputText("addr_city", "", 30, 20, "", "", false, "tabindex=5"));
            sHTML += HTML.TableData(HTML.InputText("addr_state", "", 30, 10, "", "", false, "tabindex=6"));

            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData("Zip Code:", "VIEWBOXCAPTION");
            sHTML += HTML.TableData("Country:", "VIEWBOXCAPTION");

            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData(HTML.InputText("addr_postcode", "", 10, 15, "", "", false, "tabindex=7"));

            string sSQL = "";
            sSQL = " select LTRIM(RTRIM(cast(Capt_code as nvarchar))) as code, LTRIM(RTRIM(cast(Capt_US as nvarchar))) as Caption from custom_Captions where capt_Deleted is null and capt_family='addr_country'";
            QuerySelect AddressObj = GetQuery();
            AddressObj.SQLCommand = sSQL;
            AddressObj.ExecuteReader();
            string sHTMLCountry = "";
            sHTMLCountry = "<style type=text/css> select {  font-family: Tahoma,Arial;font-size:11px; width:150px;color=#4d4f53 }</style>";
            sHTMLCountry += "";
            sHTMLCountry += "<select  name=addr_country id=addr_country tabindex=8> <option value=''>--None--</option>";
            while (!AddressObj.Eof())
            {
                sHTMLCountry += "<option value=" + AddressObj.FieldValue("code") + ">" + AddressObj.FieldValue("Caption") + "</option>";
                AddressObj.Next();
            }

            sHTMLCountry += "</select>";
            sHTML += HTML.TableData(HTML.Span("addr_country", sHTMLCountry));
            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData("Set as default address for " + EntityName + HTML.InputCheckBox("addr_default", false, "", "", false, "", "tabindex=9"), "VIEWBOXCAPTION");
            sHTML += HTML.EndTable();
            
            AddContent(HTML.Box("Address", sHTML));
        }
        public int CreateNewAddress(string objEntity, string[] objParaName, string[] objParaValue)
        {
            int iAddressId = 0;
            Record newRecord = new Record(objEntity);
            for (int i = 0; i < objParaName.Length; i++)
            {
                newRecord.SetField(objParaName[i].ToString(), objParaValue[i].ToString());
            }
            newRecord.SaveChanges();
            iAddressId = newRecord.RecordId;
            return iAddressId;
        }
        public string BuildExcelConnectionString(string Filename, bool FirstRowContainsHeaders)
        {
            return string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='{0}';Extended Properties=\"Excel 8.0;HDR={1};IMEX=1;\"",
              Filename.Replace("'", "''"),FirstRowContainsHeaders ? "Yes" : "No");
        }
        public string BuildExcel2007ConnectionString(string Filename, bool FirstRowContainsHeaders){
            return string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR={1};IMEX=1;\";",
              Filename.Replace("'", "''"),FirstRowContainsHeaders ? "Yes" : "No");

        }
        #region New Enity Insert Update
        public int InsertUpdateEnity(string objEntity, string[] objParamName, string[] objParamValue)
        {
            int Result = 0;
            try
            {                
                Record newRecord = new Record(objEntity);
                for (int i = 0; i < objParamName.Length - 1; i++)
                {
                    newRecord.SetField(objParamName[i].ToString(), objParamValue[i].ToString().Replace("''","'"));
                }
                newRecord.SaveChanges();
                Result = newRecord.RecordId;                
            }
            catch(Exception ex)
            {
                AddError(ex.Message);
                LogMessage(ex.Message);
            }
            return Result;
        }
        #endregion
        public string GetUploadRecord(string FilePath, string Extension, string isHDR)
        //public DataSet GetUploadRecord(string FilePath, string Extension, string isHDR)
        {            
         //   string error = "";
            string SheetName = "";
            string conStr = "";
            //DataSet ds = null;
            try
            {
                //conStr = BuildExcel2007ConnectionString(FilePath, true);
                switch (Extension)
                {
                    case ".xls": //Excel 97-03
                        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ FilePath +";Extended Properties='Excel 8.0;HDR={1};IMEX=1;";
                        break;
                    case ".xlsx": //Excel 07
                        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties='Excel 8.0;HDR={1};IMEX=1;";
                        break;
                }
                conStr = String.Format(conStr, FilePath, isHDR);
                
                OleDbConnection connExcel = new OleDbConnection(conStr);
                OleDbCommand cmdExcel = new OleDbCommand();
                OleDbDataAdapter oda = new OleDbDataAdapter();
                DataTable dt = new DataTable();
                cmdExcel.Connection = connExcel;

                //Get the name of First Sheet
                connExcel.Open();
                DataTable dtExcelSchema;
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                connExcel.Close();
                                
                //Read Data from First Sheet
                connExcel.Open();
                cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
                oda.SelectCommand = cmdExcel;
                oda.Fill(dt);
                connExcel.Close();
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
                //error = ex.Message;
            }
            return conStr;
        }
        public string GetRateCardStatusBlock(string EnityName, string StatusMsg, string isValidColumn, string sRowCount)
        {
         //   string InstructionText = "";
            string sHTML = "";
            
            sHTML += HTML.StartTable();
            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData("<span style='float:left;'>Status Type</span><br />", "GRIDHEAD", "width=150px height=25px");
            sHTML += HTML.TableData("<span style='float:left;'>Application Message</span>", "GRIDHEAD");
            sHTML += HTML.TableRow("");
            if (isValidColumn == "")
            {
                if (sRowCount != "N")
                {
                    sHTML += HTML.TableData("&nbsp;&nbsp;<span font-Size:3px;'>Message</span>", "VIEWBOX");
                }
                else
                    sHTML += HTML.TableData("&nbsp;&nbsp;<span font-Size:3px;'>Message</span>", "VIEWBOX");
            }
            else if (isValidColumn == "false" || isValidColumn == "ROW")
                sHTML += HTML.TableData("&nbsp;&nbsp;<span font-Size:3px;'>Message</span>", "VIEWBOX");            
            sHTML += HTML.TableData("<span font-Size:10px;'>" + StatusMsg + "</span>", "VIEWBOX");
            sHTML += "<BR>";
            sHTML += HTML.TableRow("");
            sHTML += HTML.EndTable();
            sHTML += "<BR>";
            AddContent(HTML.Box("<span style='color:#2B547E;'>Sage CRM Application Status</span>", sHTML));
            return "";
        }
        public string GetTemplateBlock(string EnityName)
        {
            string InstructionText = "";
            string sHTML = "";
            string sFileURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/";
            InstructionText += "<span style='margin: 1em 3em 10px 5.5em;'><B> 1.</B> Download the template to be used for Import Process by clicking on <B>Download Template</B> link.</span>";
            InstructionText += " <BR> ";
            InstructionText += "<span style='margin: 1em 3em 10px 5.5em;'><B> 2.</B> Copy and Paste all the data to be imported in the downloaded sheet and Save the same on your machine.</span>";
            InstructionText += " <BR> ";
            InstructionText += "<span style='margin: 1em 3em 10px 5.5em;'><B> 3.</B> Click on <B>Browse</B> button to get the excel file to be used for import process. (saved in above step)</span>";
            InstructionText += " <BR> ";
            InstructionText += "<span style='margin: 1em 3em 10px 5.5em;'><B> 4.</B> Once file is selected click on <B>Import button</B> to start import process.</span>";
            InstructionText += " <BR> ";
            InstructionText += "<span style='margin: 1em 3em 10px 5.5em;'><B> 5.</B> This process will import records in Sage CRM.</span>";
            InstructionText += " <BR> ";
            InstructionText += " <BR> ";

            sHTML += HTML.StartTable();
            sHTML += HTML.TableRow("");            
            sHTML += HTML.TableData(InstructionText, "AdminHomeDescription");
            sHTML += "<BR>";            
            sHTML += HTML.TableRow("");
            //sHTML += HTML.TableRow("");
            if (EnityName == "Client")
                sHTML += HTML.TableData("<span style='margin: 1em 3em 10px 55em;font-Size:14px;'><a href='" + sFileURL + "NZPAImport/Templates/ClientTemplate.xls' class='PANEREPEAT'><u>Download Template</u></a></span>");
            if (EnityName == "Publishers")
                sHTML += HTML.TableData("<span style='margin: 1em 3em 10px 55em;font-Size:14px;'><a href='" + sFileURL + "NZPAImport/Templates/PublishersTemplate.xls' class='PANEREPEAT'><u>Download Template</u></a></span>");
            if (EnityName == "Publications")
                sHTML += HTML.TableData("<span style='margin: 1em 3em 10px 55em;font-Size:14px;'><a href='" + sFileURL + "NZPAImport/Templates/PublicationTemplate.xls' class='PANEREPEAT'><u>Download Template</u></a></span>");
            if (EnityName == "Rate Card")
                sHTML += HTML.TableData("<span style='margin: 1em 3em 10px 55em;font-Size:14px;'><a href='" + sFileURL + "NZPAImport/Templates/RateCardTemplate.xlsx' class='PANEREPEAT'><u>Download Template</u></a></span>");
            if (EnityName == "Booking")
                sHTML += HTML.TableData("<span style='margin: 1em 3em 10px 55em;font-Size:14px;'><a href='" + sFileURL + "NZPAImport/Templates/PlanTemplate.xlsx' class='PANEREPEAT'><u>Download Template</u></a></span>");
            sHTML += HTML.EndTable();
            sHTML += "<BR>";
            AddContent(HTML.Box("Import " + EnityName + " Data", sHTML));

            return "";
        }
        public string GetStatusBlock(string EnityName, string StatusMsg, string State)
        {
            string sHTML = "";            

            sHTML += HTML.StartTable();
            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData("<span style='float:left;'>Status Type</span><br />", "GRIDHEAD", "width=150px height=25px");
            sHTML += HTML.TableData("<span style='float:left;'>Application Message</span>", "GRIDHEAD"); 
            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData("&nbsp;&nbsp;<span font-Size:3px;'>" + State + "</span>", "VIEWBOX");
            sHTML += HTML.TableData("<span font-Size:10px;'>" + StatusMsg + "</span>", "VIEWBOX");
            sHTML += "<BR>";
            sHTML += HTML.TableRow("");
            sHTML += HTML.EndTable();
            sHTML += "<BR>";

            AddContent(HTML.Box("<span style='color:#2B547E;'>Sage CRM Application Status</span>", sHTML));

            return "";
        }
        public string GetStatusBlock(string EnityName, string StatusMsg, string isValidColumn, string sRowCount)
        {
           // string InstructionText = "";
            string sHTML = "";

            sHTML += HTML.StartTable();
            sHTML += HTML.TableRow("");
            sHTML += HTML.TableData("<span style='float:left;'>Status Type</span><br />", "GRIDHEAD", "width=150px height=25px");
            sHTML += HTML.TableData("<span style='float:left;'>Application Message</span>", "GRIDHEAD");
            sHTML += HTML.TableRow("");
            if (isValidColumn == "")
            {
                if (sRowCount != "N")
                {
                    sHTML += HTML.TableData("&nbsp;&nbsp;<span font-Size:3px;'>Success</span>", "VIEWBOX");
                }
                else
                    sHTML += HTML.TableData("&nbsp;&nbsp;<span font-Size:3px;'>Error</span>", "VIEWBOX");
            }
            else if (isValidColumn == "false")
                sHTML += HTML.TableData("&nbsp;&nbsp;<span font-Size:3px;'>Error</span>", "VIEWBOX");
            sHTML += HTML.TableData("<span font-Size:10px;'>" + StatusMsg + "</span>", "VIEWBOX");
            sHTML += "<BR>";
            sHTML += HTML.TableRow("");
            sHTML += HTML.EndTable();
            sHTML += "<BR>";
            AddContent(HTML.Box("<span style='color:#2B547E;'>Sage CRM Application Status</span>", sHTML));
            return "";
        }
        public  OleDbCommand oleExcelCommand = default(OleDbCommand);
        public  OleDbDataReader oleExcelReader = default(OleDbDataReader);
        public  OleDbConnection oleExcelConnection = default(OleDbConnection);
        public DataTable ConvertToDataTable(string FileName)
        {
            DataTable res = null;

            try
            {
                string path = System.IO.Path.GetFullPath(FileName);                
                oleExcelConnection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';");                
                oleExcelConnection.Open();
                
                OleDbCommand cmd = new OleDbCommand(); ;
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                DataTable dt = oleExcelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                }
                cmd.Connection = oleExcelConnection;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                res = ds.Tables["excelData"];                
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                oleExcelConnection.Close();
            }

            return res;

        }
        public bool ProgressWorkflow(string EntityPrimaryID, string EntityName, string WorkflowName, string WorkflowState)
        {
            try
            {
                Record objTableID = FindRecord("Custom_Tables", "Bord_Deleted is null and Bord_Name='" + EntityName + "'");
                string sTableID = objTableID.GetFieldAsString("Bord_TableId");

                string QueryStr = "select WkSt_StateId,Work_WorkflowId from workflowstate (nolock)";
                QueryStr += " left outer join Workflow (nolock) on WkSt_WorkflowId = Work_WorkflowId ";
                QueryStr += " where Work_Description='" + WorkflowName + "' and WkSt_Name='" + WorkflowState + "' and wkst_deleted is null and Work_Deleted is null";
                QuerySelect sQueryObj = GetQuery();
                sQueryObj.SelectSql(QueryStr);

                string sStateid = sQueryObj.FieldValue("WkSt_StateId").ToString();
                string sWorkflowID = sQueryObj.FieldValue("Work_WorkflowId").ToString();

                //'Update Workflow State to Contract Issued
                Record recWorkflowInstance = FindRecord("WorkflowInstance", "WkIn_WorkflowId='" + sWorkflowID + "' And WkIn_CurrentEntityId='" + sTableID + "' and WkIn_CurrentRecordId='" + EntityPrimaryID + "'");
                recWorkflowInstance.SetField("WkIn_CurrentStateId", sStateid);
                recWorkflowInstance.SaveChanges();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
            #region Redirect to Visit Service Report page

            string sURL = "/" + Dispatch.InstallName + "/" + "eware.dll/Do?SID=" + Dispatch.EitherField("SID") + "&Act=432&Mode=1&CLk=T&Key0=58&Key37=" + EntityPrimaryID + "&Key58=" + EntityPrimaryID + "&visi_visitid=" + EntityPrimaryID + "&dotnetdll=WhiteWater&dotnetfunc=RunVisitServiceReport&J=Service Report&T=Visit";
            Dispatch.Redirect(sURL);

            #endregion
        }

        public void SetStageStatus(string Entity, string entityid, string stage, string status)
        {
            Record recBooking = FindRecord(Entity, " book_bookingid='" + entityid + "'");

            recBooking.SetField("book_stage", stage);
            recBooking.SetField("book_status", status);

            recBooking.SaveChanges();
        }

        public void getSMTPDetails()
        {
            try
            {
                string sSQL = "select Parm_Name,Parm_Value from Custom_SysParams where Parm_Name IN('SMTPServer','SMTPPort','SMTPPassword','SMTPUserName')";
                QuerySelect sQueryObj = GetQuery();

                sQueryObj.SQLCommand = sSQL;
                sQueryObj.ExecuteReader();
                if (!sQueryObj.Eof())
                {
                    while (!sQueryObj.Eof())
                    {
                        if (sQueryObj.FieldValue("Parm_Name").ToString().ToLower() == "smtpserver")
                        {
                            servername = sQueryObj.FieldValue("Parm_Value").ToString();
                        }
                        if (sQueryObj.FieldValue("Parm_Name").ToString().ToLower() == "smtpport")
                        {
                            smtpport = sQueryObj.FieldValue("Parm_Value").ToString();
                        }
                        if (sQueryObj.FieldValue("Parm_Name").ToString().ToLower() == "smtppassword")
                        {
                            smtppwd = sQueryObj.FieldValue("Parm_Value").ToString();
                        }
                        if (sQueryObj.FieldValue("Parm_Name").ToString().ToLower() == "smtpusername")
                        {
                            smtpusername = sQueryObj.FieldValue("Parm_Value").ToString();
                        }
                        sQueryObj.Next();
                    }
                }
            }
            catch (Exception ex)
            {
                AddContent(ex.Message.ToString());
            }
        }
        public bool CheckDateIsExists(string sDate)
        {
            //string screenDate = Convert.ToDateTime(sDate).ToString("yyyy-MM-dd");
            //string screenDate = Convert.ToDateTime(sDate).ToString("dd/MM/yyyy");
            string screenDate = sDate;

            string strSQL = "select * from HolidaySetItems where convert(varchar,HSIt_HolidayDate,103) ='" + screenDate + "' and HSIt_Deleted is null";
            QuerySelect sQueryObj = GetQuery();

            sQueryObj.SQLCommand = strSQL;
            sQueryObj.ExecuteReader();

            if (!sQueryObj.Eof())
            {
                return false;
            }
            else
            {
                return true;
            }
        }
    }  
}