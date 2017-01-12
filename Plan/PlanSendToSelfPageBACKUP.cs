using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;
using System.Data;
using System.IO;
using System.Net.Mail;
using System.Net;
using System.Reflection;
using System.Diagnostics;

namespace NZPACRM.Plan
{
    public class PlanSendToSelfPageBACKUP : Web
    {
        Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        Microsoft.Office.Interop.Excel.Application oXL;  
    
        string userID = "";
        string smtpusername = "";
        string smtppwd = "";
        string servername = "";
        string smtpport = "";
        string sBookingID = "";

        string companyID = "";
        string bookingID = "";
        string bookRefe = "";
        string sAgency = "";
        string sAdvertiser = "";
        string sAgencyContact = "";
        string sQuoteVersionRef = "";
        string sCreationDate = "";
        string sCreatedBy = "";
        string mergedFilePath = "";
        string orgFilePath = "";
        string body = "";
        string header = "";

        CRMHelper crmHelperObj = new CRMHelper();

        public PlanSendToSelfPageBACKUP()
            : base()
        {
            if (!String.IsNullOrEmpty(Dispatch.EitherField("Key58")))
            {
                sBookingID = Dispatch.EitherField("Key58");
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key37")))
            {
                sBookingID = Dispatch.EitherField("Key37");
            }
            //set agency id & booking reference number 
            Record recBook = FindRecord("booking", "book_bookingid=" + sBookingID);
            companyID = recBook.GetFieldAsString("book_agency");
            bookRefe = recBook.GetFieldAsString("book_reference");
        }
        public override void BuildContents()
        {           
            //'Get To Email Address
            //'Get from Email Address
            //Send Email with attachment
            //Create Communication
            //Progress Workflow
            //AddContent("Start");
            ReadExistingExcel();            
            getSMTPDetails();
            getSMPTPassword();
            SendEmail();
        }

        private void getSMPTPassword()
        {
            string strSQL = "select * from Custom_Captions where Capt_Code='smtppassword' and Capt_Family='smtppassword'";            

            QuerySelect objPblcnRec = GetQuery();
            objPblcnRec.SQLCommand = strSQL;
            objPblcnRec.ExecuteReader();

            if (!objPblcnRec.Eof())
            {
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("capt_US").ToString()))
                {
                    smtppwd = objPblcnRec.FieldValue("capt_US").ToString();                    
                }
            }
        }


        public void writexmldetails()
        {
            string strSQL = "select * from Planbuilder where pbnr_plan = " + sBookingID;
            //string sCode = "";

            QuerySelect objAddressRec = GetQuery();
            objAddressRec.SQLCommand = strSQL;
            objAddressRec.ExecuteReader();

            while (!objAddressRec.Eof())
            {
                body += "\t\t<ad>\n";
                body += "\t\t\t<ad_details>\n";
                body += "\t\t\t\t<section_id>" + objAddressRec.FieldValue("pnbr_sections") + "</section_id>\n"; // PLEASE CHANGE ME!!!!!
                string strSQLsection = "select * from Section where sctn_sctn_sectionid = " + objAddressRec.FieldValue("pnbr_sections");
                //string sCode = "";

                QuerySelect objAddressRecsection = GetQuery();
                objAddressRecsection.SQLCommand = strSQLsection;
                objAddressRecsection.ExecuteReader();

                body += "\t\t\t\t<section_name>" + objAddressRecsection.FieldValue("sctn_name") + "</section_name>\n";
                body += "\t\t\t\t<sub_section_id>" + objAddressRec.FieldValue("pnbr_subsection") + "</sub_section_id>\n";

                strSQLsection = "select * from Subsection where suse_subsectionid = " + objAddressRec.FieldValue("pnbr_subsection");
                //string sCode = "";

                objAddressRecsection = GetQuery();
                objAddressRecsection.SQLCommand = strSQLsection;
                objAddressRecsection.ExecuteReader();
                body += "\t\t\t\t<sub_section_name>" + objAddressRecsection.FieldValue("suse_name") + "</sub_section_name>\n";
                body += "\t\t\t\t<colour>" + objAddressRec.FieldValue("pnbr_color") + "</colour>\n";
                body += "\t\t\t\t<caption>" + "" + "</caption>\n";
                body += "\t\t\t\t<placement_comment>" + "" + "</placement_comment>\n";
                body += "\t\t\t</ad_details>\n";
                body += "\t\t\t<ad_size>\n";
                body += "\t\t\t\t<ad_size_name>" + objAddressRec.FieldValue("pnbr_standardsize") + "</ad_size_name>\n";
                body += "\t\t\t\t<depth>" + "" + "</depth>\n";
                body += "\t\t\t\t<depth_unit>" + "" + "</depth_unit>\n";
                body += "\t\t\t\t<columns>" + "" + "</columns>\n";
                body += "\t\t\t</ad_size>\n";
                body += "\t\t\t<schedule>\n";
                body += "\t\t\t\t<run_dates>\n";
                string day = objAddressRec.FieldValue("pnbr_days");
                string realday = day.Substring(1, 3);
                string rateday = dayswap(realday);


            }
        }

        private string dayswap(string orday)
        {
            switch (orday)
            {
                case "Mon": return "Monday";
                case "Tues": return "tuesday";
                case "Wed": return "wednesday";
                case "Thur": return "thrusday";
                case "Fri": return "friday";
                case "Sat": return "saturday";
                default: return "sunday";
            }

        }

        public void ReadExistingExcel()
        {
            //string path = @"E:\Projects\ETL914-C\Docs\SendToSelf.xls";
            orgFilePath = GetLibraryPath() + @"\Plan\AgencyDispatchNotice\Template\SendToSelf.xls";

            string checkPathGenTemplate = GetLibraryPath() + @"\Plan\AgencyDispatchNotice\GeneratedTemplate";

            if (!Directory.Exists(checkPathGenTemplate))
            {
                Directory.CreateDirectory(checkPathGenTemplate);
            }

            if (orgFilePath == "")
            {
                //MessageBox.Show("CPS work order sheet path is not found");
            }
            else
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                //'oXL.Visible = true;
                oXL.DisplayAlerts = false;
                try
                {
                    mWorkBook = oXL.Workbooks.Open(orgFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    mWorkSheets = mWorkBook.Worksheets;
                    //Get all the sheets in the workbook
                    //Get the already exists sheet
                    mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item(1);
                    Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
                    //'Set Agency Name
                    SetAgency();
                    //set detail grid value
                    SetDetailGrid();
                }
                catch (Exception ex)
                {
                    AddContent(ex.Message.ToString());
                }
                string date = "";
                date = DateTime.Now.Date.ToShortDateString().Replace(@"/", "-");
                mergedFilePath = GetLibraryPath() + @"\Plan\AgencyDispatchNotice\GeneratedTemplate\SendToSelf_" + date + DateTime.Now.Millisecond.ToString() + ".xls";
                string sSavedCPSpath = mergedFilePath;

                try
                {
                    mWorkBook.SaveAs(sSavedCPSpath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                    Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value);
                    mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);

                    Process[] processes = Process.GetProcessesByName("EXCEL");

                    foreach (var process in processes)
                    {
                        if (process.MainWindowTitle == "")
                            process.Kill();
                    }
                    mWSheet1 = null;
                    mWorkBook = null;

                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    //' mWorkBook.Close();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);

                    //MessageBox.Show("CPS work order sheet generated sucessfully. Click Ok to continue.");
                }
                catch (Exception ex)
                {
                    //AddContent("In-ReadExistingExcel - Catch");
                    AddContent(ex.Message.ToString());
                    //MessageBox.Show("Failed to generte CPS work order sheet. File is open or may be path is not defined to save the file.");
                }
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(mWorkBook);
            }
        }
        private void SetAgency()
        {
            string sCode = "";
            string strSQL = "select * from vCompanyPE where Comp_CompanyId=" + companyID + "";
            QuerySelect objAddressRec = GetQuery();
            objAddressRec.SQLCommand = strSQL;
            objAddressRec.ExecuteReader();

            //'Declare variables
            //AddContent(objBookingList.GetFieldAsString("comp_name"));

            sAgency = objAddressRec.FieldValue("Comp_Name");
            sAdvertiser = "";
            //sAgencyContact = objBookingList.GetFieldAsString("Comp_PhoneNumber");
            //sQuoteVersionRef = objBookingList.GetFieldAsString("");
            sCreationDate = objAddressRec.FieldValue("Comp_CreatedDate");
            sCreatedBy = objAddressRec.FieldValue("comp_createdBy");

            mWSheet1.Cells[8, 3] = sAgency;
            mWSheet1.Cells[9, 3] = sAgencyContact;
            mWSheet1.Cells[10, 3] = sAdvertiser;
            mWSheet1.Cells[11, 3] = sQuoteVersionRef;
            mWSheet1.Cells[8, 10] = sCreationDate;
            mWSheet1.Cells[9, 10] = sCreatedBy;
        }
        private void SetDetailGrid()
        {
            int j = 14;
            string sCode = "";
            string strSQL = "select Comp_Name,Comp_PhoneNumber,Addr_Address1,comp_Category,pblc_Name,pnbr_sections,pnbr_size,pnbr_color,*";
            strSQL += " from vCompanype inner join vAddressCompany on Comp_CompanyId=AdLi_CompanyID inner join Booking on Comp_CompanyId=book_agency ";
            strSQL += " left join Planbuilder on book_BookingID=pnbr_plan left join Publications on pnbr_publications = pblc_PublicationsID left JOIN Sections on sctn_Sctn_sectionid=pnbr_sections ";
            strSQL += " where book_BookingID=" + sBookingID + " and Comp_CompanyId = " + companyID + "";

            QuerySelect objPblcnRec = GetQuery();
            objPblcnRec.SQLCommand = strSQL;
            objPblcnRec.ExecuteReader();

            while (!objPblcnRec.Eof())
            {
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 1] = "";
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 2] = objPblcnRec.FieldValue("pblc_Name").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("sctn_name").ToString()))
                {
                    mWSheet1.Cells[j, 3] = objPblcnRec.FieldValue("sctn_name").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 4] = "";
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_size").ToString()))
                {
                    mWSheet1.Cells[j, 5] = objPblcnRec.FieldValue("pnbr_size").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 6] = "";
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 7] = "";
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_color").ToString()))
                {
                    mWSheet1.Cells[j, 8] = objPblcnRec.FieldValue("pnbr_color").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 9] = "";
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 10] = "";
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 11] = "";
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 12] = "";
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 13] = "";
                }
                j++;
                objPblcnRec.Next();
            }
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
                            servername = sQueryObj.FieldValue("Parm_Value").ToString();  //127.0.0.1
                            //gmail smtp details - servername = "smtp.gmail.com";
                            //servername = "smtp.gmail.com";
                        }
                        if (sQueryObj.FieldValue("Parm_Name").ToString().ToLower() == "smtpport")
                        {
                            smtpport = sQueryObj.FieldValue("Parm_Value").ToString(); //11026
                            //gmail port number - smtpport = "25";
                            //smtpport = "25";
                        }
                        //if (sQueryObj.FieldValue("Parm_Name").ToString().ToLower() == "smtppassword")
                        //{
                        //    smtppwd = sQueryObj.FieldValue("Parm_Value").ToString();
                        //    //smtppwd = "giplinc"; 
                        //}
                        if (sQueryObj.FieldValue("Parm_Name").ToString().ToLower() == "smtpusername")
                        {
                            smtpusername = sQueryObj.FieldValue("Parm_Value").ToString();
                            //smtpusername = "greytrix@gmail.com";
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
        public void SendEmail()
        {
            userID = this.CurrentUser.UserId.ToString();
            string sToEmailAddress = ToEmailID(userID);
            string sFromEmailAddress = GetFromEmailAddress();
            string sBookRefe = "";
            
            if (sToEmailAddress.ToString() == "")
            {
                AddError("Unable To Send Email: To Email address is not available.");
            }

            if (sFromEmailAddress.ToString() == "")
            {
                AddError("Unable To Send Email: From Email address is not available.");
            }

            StringBuilder sb = new StringBuilder();
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress(sFromEmailAddress);
            msg.To.Add(sToEmailAddress);
            msg.Subject = bookRefe + " News Works Plan";
            sb.Append("<br>");
            Record ObjEmailTemplate = FindRecord("EmailTemplates", "EmTe_Name='PlanToSelf' and EmTe_Entity='Booking'");
            if (!ObjEmailTemplate.Eof())
                msg.Body = ObjEmailTemplate.GetFieldAsString("EmTe_Comm_Email");

            msg.IsBodyHtml = true;            
            System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(servername);

            //string sFileName = GetLibraryPath() + "VisitServiceReports\\" + objFileRec.GetFieldAsString("libr_FileName");
            string sFileName = mergedFilePath;
            try
            {
                if (sFileName != null)
                    msg.Attachments.Add(new Attachment(sFileName));
            }
            catch (Exception ex)
            {                
                AddError("Unable To Send Attachment");
                AddContent(ex.Message.ToString());
                AddUrlButton("Continue", "Continue.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage" + "&book_bookingid=" + sBookingID + ""));
            }

            var _with1 = smtpClient;
            smtpClient.Port = Convert.ToInt32(smtpport);

            smtpClient.Credentials = new NetworkCredential(smtpusername, smtppwd);
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;

            try
            {
                smtpClient.EnableSsl = true;                
                
                _with1.Send(msg);
                AddInfo("Email Sent Successfully");
                bool wrkflwResult;

                wrkflwResult = crmHelperObj.ProgressWorkflow(sBookingID, "Booking", "Booking Workflow", "Self");
                if (wrkflwResult)
                {
                    AddInfo("Workflow progressed successfully to Self state");

                    crmHelperObj.SetStageStatus("Booking", sBookingID, "sendtoself", "InProgress");
                }
                else
                {
                    AddInfo("Error Occurred during worlflow progress");
                }                
                string sURL = Url("432") + "&book_bookingid=" + sBookingID;
                AddUrlButton("Continue", "Continue.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage" + "&book_bookingid=" + sBookingID + ""));                
            }
            catch (Exception ex)
            {
                AddUrlButton("Continue", "Continue.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage" + "&book_bookingid=" + sBookingID + ""));
                AddError("Error Occured While Sending An Email: " + ex.Message);
            }
        }
        public string ToEmailID(string sPersonID)
        {
            string sToEmail = "";
            string sSQL = " select * from Users where User_UserId='" + sPersonID + "'";
            QuerySelect sQueryObj = GetQuery();

            sQueryObj.SQLCommand = sSQL;
            sQueryObj.ExecuteReader();
            if (!sQueryObj.Eof())
            {
                sToEmail = sQueryObj.FieldValue("User_EmailAddress").ToString();
            }
            return sToEmail;
        }      
        public string GetLibraryPath()
        {            
            string sLibrPath = "";
            Record objLibrRec = FindRecord("custom_sysparams", "parm_name='DocStore'");
            if (!objLibrRec.Eof())
                sLibrPath = objLibrRec.GetFieldAsString("parm_value");

            return sLibrPath;
        }
        public string GetFromEmailAddress()
        {
            string sFromEmailAddress = "";
            string sSQL = " select * from Custom_EmailAddress (nolock) where  emse_displayname='System Administrator' and EmSe_Deleted is null";
            QuerySelect sQueryObj = GetQuery();

            sQueryObj.SQLCommand = sSQL;
            sQueryObj.ExecuteReader();
            if (!sQueryObj.Eof())
            {
                sFromEmailAddress = sQueryObj.FieldValue("EmSe_EmailAddress").ToString();
            }
            return sFromEmailAddress;
        }

        private void SetAgencyXml()
        {

            Record recBook = FindRecord("booking", "book_bookingid=" + sBookingID);
            header += "\t\t<id>" + recBook.GetFieldAsString("book_bookingid") + "</id>\n";
            header += "\t\t<action>" + recBook.GetFieldAsString("book_name") + "</action>\n";
            header += "\t\t<customer>\n";
            header += "\t\t\t<client_id>" + recBook.GetFieldAsString("book_Client") + "</client_id>\n";
            string strSQL = "select * from Client where client_clientid = " + recBook.GetFieldAsString("book_Client");
            //string sCode = "";

            QuerySelect objAddressRec = GetQuery();
            objAddressRec.SQLCommand = strSQL;
            objAddressRec.ExecuteReader();

            header += "\t\t\t<client_name>" + objAddressRec.FieldValue("client_name") + "</client_name>\n";
            header += "\t\t\t<agency_id>" + recBook.GetFieldAsString("book_agency") + "</agency_id>\n";


            strSQL = "select * from Company where company_companyid = " + recBook.GetFieldAsString("book_agency");
            //string sCode = "";

            objAddressRec = GetQuery();
            objAddressRec.SQLCommand = strSQL;
            objAddressRec.ExecuteReader();
            ////'Declare variables
            ////AddContent(objBookingList.GetFieldAsString("comp_name"));

            sAgency = objAddressRec.FieldValue("Comp_Name");
            header += "\t\t\t<agency_name>" + sAgency + "</agency_name>\n";
            header += "\t\t\t<commission>" + recBook.GetFieldAsString("book_standardrate") + "</commission>\n";
            header += "\t\t</customer>\n";
            
        }
    }
}

