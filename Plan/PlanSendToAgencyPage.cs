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
    public class PlanSendToAgencyPage:Web
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
        string sDay = "";
        string companyID = "";
        string bookingID = "";
        string bookRefe = "";
        string sAgency = "";
        string sAdvertiser = "";
        string sAgencyContact = "";
        string sQuoteVersionRef = "";
        string sNewsWorksRef = "";
        string sStatus = "";
        string sCreationDate = "";
        string sCreatedBy = "";
        string sCampaign = "";
        string sBilledBy = "";
        string xmlwrite = "";
        string mergedFilePath = "";
        string orgFilePath = "";
        string bookContact = "";
        string header = "";
        string body = "";
        string xmlpath = "";
        int max = 0;

        CRMHelper crmHelperObj = new CRMHelper();
        public PlanSendToAgencyPage() : base() {
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
            bookContact = recBook.GetFieldAsString("book_contact");   
        }
        public override void BuildContents()
        {            
            ReadExistingExcel();
            string getpubs  = "select * from Planbuilder where pnbr_plan = " + sBookingID;
            List<int>pubids = new List<int>();
              QuerySelect objAddressRec = GetQuery();
    
            objAddressRec.SQLCommand = getpubs;
            objAddressRec.ExecuteReader();
            AddContent("papercut");
            while(!objAddressRec.Eof()){
                int id = Int32.Parse(objAddressRec.FieldValue("pnbr_publications"));
                if (!pubids.Contains(id)){
                     string getrecs  = "select * from Planbuilder where pnbr_plan = " + sBookingID + " and pnbr_publications = " + objAddressRec.FieldValue("pnbr_publications");
                    
                     QuerySelect objrecs = GetQuery();
    
                     objrecs.SQLCommand = getrecs;
                      objrecs.ExecuteReader();
                    writeheader();
                    SetAgencyXml(objAddressRec.FieldValue("pnbr_publications"));
                    while(!objrecs.Eof()){
                        string sections = "";
                        if (!string.IsNullOrEmpty(objrecs.FieldValue("pnbr_sections").ToString())) sections = objrecs.FieldValue("pnbr_sections");
                        string subsections = "";
                        if (!string.IsNullOrEmpty(objrecs.FieldValue("pnbr_subsection").ToString())) subsections = objrecs.FieldValue("pnbr_subsection");
                        string color = "";
                        if (!string.IsNullOrEmpty(objrecs.FieldValue("pnbr_color").ToString())) color = objrecs.FieldValue("pnbr_color");
                        string ss = "";
                        if (!string.IsNullOrEmpty(objrecs.FieldValue("pnbr_standardsize").ToString())) ss = objrecs.FieldValue("pnbr_standardsize");
                        string high = "";
                        if (!string.IsNullOrEmpty(objrecs.FieldValue("pnbr_height").ToString())) high = objrecs.FieldValue("pnbr_height");
                        string days = "";
                        if (!string.IsNullOrEmpty(objrecs.FieldValue("pnbr_days").ToString())) days = objrecs.FieldValue("pnbr_days");
                        string date = "";
                        if (!string.IsNullOrEmpty(objrecs.FieldValue("pnbr_date").ToString())) date = objrecs.FieldValue("pnbr_date");
                        string key = "";
                        if (!string.IsNullOrEmpty(objrecs.FieldValue("pnbr_KeyNumber").ToString())) key = objrecs.FieldValue("pnbr_KeyNumber");
                        string rc = "";
                        if (!string.IsNullOrEmpty(objrecs.FieldValue("pnbr_ratecard").ToString())) rc = objrecs.FieldValue("pnbr_ratecard");
                        string pubs = "";
                        if (!string.IsNullOrEmpty(objrecs.FieldValue("pnbr_publications").ToString())) pubs = objrecs.FieldValue("pnbr_publications");
                        AddContent(pubs);

                        writexmldetails(sections, subsections,color, ss, high, days, date, key, rc, pubs);
                        objrecs.Next();
                    }
                    string end = "\t\t<date>" + DateTime.Now.ToString() + "</date>\n";
                    end += "\t</export>\n";
                    end += "</booking>";
                    xmlpath = GetLibraryPath() + @"\Plan\" +max + ".xml" ;
                    xmlwrite = header + body + end;
                    System.IO.File.WriteAllText(xmlpath, xmlwrite);
                    max++;

                    pubids.Add(id);
                }
                objAddressRec.Next();
            }

            //writeheader();
            //SetAgencyXml();
            
            getSMTPDetails();
            getSMPTPassword();
            SendEmail();            
        }

        public void writeheader()
        {
            xmlwrite += "<booking>";
            xmlwrite += "\t<export>";
        }

        public void writexmldetails(string sections, string subsection, string color, string standardsize, string height, string days, string date, string key, string rate,string pubname  )
        {
           

            
                body += "\t\t<ad>\n";
                body += "\t\t\t<ad_details>\n";
                body += "\t\t\t\t<section_id>" + "" +  "</section_id>\n"; // PLEASE CHANGE ME!!!!!
                string strSQLsection = "select * from Sections where sctn_sctn_sectionid = " + sections;
                //string sCode = "";

                QuerySelect objAddressRecsection = GetQuery();
                objAddressRecsection.SQLCommand = strSQLsection;
                objAddressRecsection.ExecuteReader();
                if (!string.IsNullOrEmpty(objAddressRecsection.FieldValue("sctn_name").ToString()))
                body += "\t\t\t\t<section_name>" + objAddressRecsection.FieldValue("sctn_name") + "</section_name>\n";
                else body += "\t\t\t\t<section_name>" + "" + "</section_name>\n";
                body += "\t\t\t\t<sub_section_id>" +"" + "</sub_section_id>\n";

                strSQLsection = "select * from Subsection where suse_subsectionid = " + subsection;
                //string sCode = "";

                objAddressRecsection = GetQuery();
                objAddressRecsection.SQLCommand = strSQLsection;
                objAddressRecsection.ExecuteReader();
                if (!string.IsNullOrEmpty(objAddressRecsection.FieldValue("suse_name").ToString()))
                body += "\t\t\t\t<sub_section_name>" + objAddressRecsection.FieldValue("suse_name") + "</sub_section_name>\n";
                else body += "\t\t\t\t<sub_section_name>" + "" + "</sub_section_name>\n";
                body += "\t\t\t\t<colour>" + color + "</colour>\n";
                body += "\t\t\t\t<caption>" + "" + "</caption>\n";
                body += "\t\t\t\t<placement_comment>" + "" + "</placement_comment>\n";
                body += "\t\t\t</ad_details>\n";
                body += "\t\t\t<ad_size>\n";
                if (!string.IsNullOrEmpty(standardsize.ToString()))
                {
                    body += "\t\t\t\t<ad_size_name>" + standardsize+ "</ad_size_name>\n";
                    body += "\t\t\t\t<depth>" + "" + "</depth>\n";
                    body += "\t\t\t\t<depth_unit>" + "" + "</depth_unit>\n";
                    body += "\t\t\t\t<columns>" + "" + "</columns>\n";
                }
                else
                {
                    body += "\t\t\t\t<ad_size_name>" +""+ "</ad_size_name>\n";
                    body += "\t\t\t\t<depth>" + height + "</depth>\n";
                    body += "\t\t\t\t<depth_unit>" + "" + "</depth_unit>\n";
                    body += "\t\t\t\t<columns>" + "" + "</columns>\n";
                }
                body += "\t\t\t</ad_size>\n";
                body += "\t\t\t<schedule>\n";
                body += "\t\t\t\t<run_dates>\n";
                
                string realday = days.Substring(1, 3);
                string rateday = dayswap(realday);

                body += "\t\t\t\t\t<run_date>\n";
                body += "\t\t\t\t\t\t<date>" + date.Substring(0,10) + "</date>\n";
                string rateme = "select * from RatesCard where rate_RatesCardID = " + rate;
                QuerySelect objrate = GetQuery();
                objrate.SQLCommand = rateme;
                objrate.ExecuteReader();
                if (!string.IsNullOrEmpty(objrate.FieldValue("rate_" + rateday).ToString()))
                body += "\t\t\t\t\t\t<price>" + objrate.FieldValue("rate_"+rateday) + "</price>\n";
                else body += "\t\t\t\t\t\t<price>" + ""+ "</price>\n";
                body += "\t\t\t\t\t\t<key_number>" + key + "</key_number>\n";
                body += "\t\t\t\t\t\t<publication>\n";
                body += "\t\t\t\t\t\t\t<pub_id>" + "" + "</pub_id>\n";
                string ratesql = "select * from Publications where pblc_publicationsid = " + pubname;
                QuerySelect objgetrate = GetQuery();
                objgetrate.SQLCommand = ratesql;
                objgetrate.ExecuteReader();
                if (!string.IsNullOrEmpty(objgetrate.FieldValue("pblc_Name").ToString()))
                body += "\t\t\t\t\t\t\t<pub_name>" + objgetrate.FieldValue("pblc_Name") + "</pub_name>\n";
                else body += "\t\t\t\t\t\t\t<pub_name>" + ""+ "</pub_name>\n";
                body += "\t\t\t\t\t\t</publication>\n";
                body += "\t\t\t\t\t</run_date>\n";
                body += "\t\t\t\t</run_dates>\n";
                body += "\t\t\t</schedule>\n";
                body += "\t\t</ad>\n";
        }

        private string dayswap(string orday)
        {
            switch (orday){
                case "Mon": return "Monday";
                case "Tues": return "tuesday";
                case "Wed": return "wednesday";
                case "Thur": return "thrusday";
                case "Fri":return "friday";
                case "Sat": return "saturday";
                default: return "sunday";
            }
             
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
        public void ReadExistingExcel()
        {
            //string path = @"E:\Projects\ETL914-C\Docs\SendToSelf.xls";            
            orgFilePath = GetLibraryPath() + @"\Plan\QuoteLayout\Template\SendToAgency.xls";

            string checkPathGenTemplate = GetLibraryPath() + @"\Plan\QuoteLayout\GeneratedTemplate";

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
                mergedFilePath = GetLibraryPath() + @"\Plan\QuoteLayout\GeneratedTemplate\SendToAgency_" + date + DateTime.Now.Millisecond.ToString() + ".xls";
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
                    AddContent(ex.Message.ToString());
                    //MessageBox.Show("Failed to generte CPS work order sheet. File is open or may be path is not defined to save the file.");

                }
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(mWorkBook);
            }
        }
        private void SetAgency()
        {
            string sCode = "";
            string strSQL = "select * from vCompanyPE inner join booking on book_Agency=Comp_CompanyId where Comp_CompanyId=" + companyID + "";
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
            sBilledBy = objAddressRec.FieldValue("book_billedby");
            sStatus = objAddressRec.FieldValue("book_Status");

            mWSheet1.Cells[9, 3] = sAgency;
            mWSheet1.Cells[10, 3] = sAgencyContact;
            mWSheet1.Cells[11, 3] = sAdvertiser;
            mWSheet1.Cells[12, 3] = sCampaign;
            mWSheet1.Cells[13, 3] = sBilledBy;
            mWSheet1.Cells[9, 10] = sCreationDate;
            mWSheet1.Cells[10, 10] = sCreatedBy;
            mWSheet1.Cells[11, 10] = sQuoteVersionRef;
            mWSheet1.Cells[12, 10] = sNewsWorksRef;
            mWSheet1.Cells[13, 10] = sStatus;            
        }
        private void SetDetailGrid()
        {
            int j = 17;
            string sCode = "";
            string strSQL = "select Comp_Name,Comp_PhoneNumber,Addr_Address1,comp_Category,pblc_Name,pblc_CreatedBy,pblc_CreatedDate,pblc_commision";
            strSQL += " pblc_Status,book_reference,pnbr_days,pnbr_discount,pnbr_color,pnbr_ratecard,pnbr_cost,pnbr_size,pnbr_sections,*";
            strSQL += " from vCompanype inner join vAddressCompany on Comp_CompanyId=AdLi_CompanyID inner join Booking on Comp_CompanyId=book_agency";
            strSQL += " left join Planbuilder on book_BookingID=pnbr_plan left join Publications on pnbr_publications = pblc_PublicationsID left JOIN Sections on sctn_Sctn_sectionid=pnbr_sections ";
            strSQL += " where book_BookingID=" + sBookingID + " and Comp_CompanyId = " + companyID + "";

            QuerySelect objPblcnRec = GetQuery();
            objPblcnRec.SQLCommand = strSQL;
            objPblcnRec.ExecuteReader();

            while (!objPblcnRec.Eof())
            {
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 1] = objPblcnRec.FieldValue("pblc_Name").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_commision").ToString()))
                {
                    mWSheet1.Cells[j, 2] = objPblcnRec.FieldValue("pblc_commision").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("sctn_name").ToString()))
                {
                    mWSheet1.Cells[j, 3] = objPblcnRec.FieldValue("sctn_name").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_days").ToString()))
                {
                    mWSheet1.Cells[j, 4] = GetDays(objPblcnRec.FieldValue("pnbr_days").ToString());
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 5] = objPblcnRec.FieldValue("pblc_Name").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_size").ToString()))
                {
                    mWSheet1.Cells[j, 6] = objPblcnRec.FieldValue("pnbr_size").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 7] = "";
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 8] = "";
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_color").ToString()))
                {
                    mWSheet1.Cells[j, 9] = objPblcnRec.FieldValue("pnbr_color").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_ratecard").ToString()))
                {
                    mWSheet1.Cells[j, 10] = objPblcnRec.FieldValue("pnbr_ratecard").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_discount").ToString()))
                {
                    mWSheet1.Cells[j, 11] = objPblcnRec.FieldValue("pnbr_discount").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_total").ToString()))
                {
                    mWSheet1.Cells[j, 12] = objPblcnRec.FieldValue("pnbr_total").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 13] = "";
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 14] = "";
                }
                j++;
                objPblcnRec.Next();
            }
        }

        private string GetDays(string objPnbrDay)
        {
            string[] arrDay = objPnbrDay.Split(',');
            foreach (string day in arrDay)
            {
                switch (day)
                {
                    case "Mon":
                        sDay += "Monday,";
                        break;
                    case "Tues":
                        sDay += "Tuesday,";
                        break;
                    case "Wed":
                        sDay += "Wednesday";
                        break;
                    case "Thur":
                        sDay += "Thursday,";
                        break;
                    case "Fri":
                        sDay += "Friday,";
                        break;
                    case "Sat":
                        sDay += "Saturday,";
                        break;
                    case "Sun":
                        sDay += "Sunday,";
                        break;
                    default:
                        sDay += "";
                        break;
                }
            }
            return sDay;
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
            string sToEmailAddress = ToEmailID();
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
            Record ObjEmailTemplate = FindRecord("EmailTemplates", "EmTe_Name='PlanToAgency' and EmTe_Entity='Booking'");
            if (!ObjEmailTemplate.Eof())
            {
                Record recUsername = FindRecord("User", "user_userid='" + CurrentUser.UserId + "'");
                msg.Body += ObjEmailTemplate.GetFieldAsString("EmTe_Comm_Email").Replace("username", recUsername.GetFieldAsString("User_FirstName") + " " + recUsername.GetFieldAsString("User_LastName"));
            }
            msg.IsBodyHtml = true;
            System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(servername);

            //string sFileName = GetLibraryPath() + "VisitServiceReports\\" + objFileRec.GetFieldAsString("libr_FileName");            
            string sFileName = mergedFilePath;
            
            try
            {
                if (sFileName != null)
                    msg.Attachments.Add(new Attachment(sFileName));
                for (int i = 0; i < max; i++)
                {
                    msg.Attachments.Add(new Attachment(GetLibraryPath() + @"\Plan\" + i.ToString() + ".xml"));
                }
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

                wrkflwResult = crmHelperObj.ProgressWorkflow(sBookingID, "Booking", "Booking Workflow", "Agency");
                if (wrkflwResult)
                {
                    AddInfo("Workflow progressed successfully to Self state");

                    crmHelperObj.SetStageStatus("Booking", sBookingID, "sendtoagency", "InProgress");
                }
                else
                {
                    AddInfo("Error Occurred during worlflow progress");
                }                
                AddUrlButton("Continue", "Continue.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage" + "&book_bookingid="+sBookingID+""));
            }
            catch (Exception ex)
            {
                AddUrlButton("Continue", "Continue.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage" + "&book_bookingid=" + sBookingID + ""));

                AddError("Error Occured While Sending An Email: " + ex.Message);
            }
        }
        public string ToEmailID()
        {            
            string sToEmail = "";
            string sSQL = " select * from Booking inner join vEmailCompanyAndPerson on Pers_PersonId = book_Contact where book_BookingID=" + sBookingID + "";
            QuerySelect sQueryObj = GetQuery();

            sQueryObj.SQLCommand = sSQL;
            sQueryObj.ExecuteReader();
            if (!sQueryObj.Eof())
            {
                sToEmail = sQueryObj.FieldValue("Emai_EmailAddress").ToString();
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

        private void SetAgencyXml(string pubxml)
        {

            header += "<booking>\n";
            header += "\t<export>\n";

            Record recBook = FindRecord("booking", "book_bookingid=" + sBookingID);
            header += "\t\t<id>" + ""+ "</id>\n";
            if (!string.IsNullOrEmpty(recBook.GetFieldAsString("book_name") .ToString()))
            header += "\t\t<action>" + recBook.GetFieldAsString("book_name") + "</action>\n";
            else header += "\t\t<action>" + "" + "</action>\n";
            header += "\t\t<customer>\n";
            if (!string.IsNullOrEmpty(recBook.GetFieldAsString("book_Client").ToString()))
                header += "\t\t\t<client_id>" + recBook.GetFieldAsString("book_Client") + "</client_id>\n";
            else header += "\t\t\t<client_id>" + ""+ "</client_id>\n";
            if (!string.IsNullOrEmpty(recBook.GetFieldAsString("book_Client").ToString()))
            {
                string strSQL = "select * from Client where client_clientid = '" + recBook.GetFieldAsString("book_Client") + "'";
                //string sCode = "";

                QuerySelect objAddressRec = GetQuery();
                objAddressRec.SQLCommand = strSQL;
                objAddressRec.ExecuteReader();

                if (!string.IsNullOrEmpty(objAddressRec.FieldValue("client_name").ToString()))
                    header += "\t\t\t<client_name>" + objAddressRec.FieldValue("client_name") + "</client_name>\n";
                else header += "\t\t\t<client_name>" + "" + "</client_name>\n";
            }
            else header += "\t\t\t<client_name>" + "" + "</client_name>\n";
            if (!string.IsNullOrEmpty(recBook.GetFieldAsString("book_agency").ToString()))
            header += "\t\t\t<agency_id>" + recBook.GetFieldAsString("book_agency") + "</agency_id>\n";
            else header += "\t\t\t<agency_id>" + ""+ "</agency_id>\n";


            string strSQLComp = "select * from Company where comp_companyid = " + recBook.GetFieldAsString("book_agency");
            //string sCode = "";

           QuerySelect  objAddresscomp = GetQuery();
            objAddresscomp.SQLCommand = strSQLComp;
            objAddresscomp.ExecuteReader();
            ////'Declare variables
            ////AddContent(objBookingList.GetFieldAsString("comp_name"));
            if (!string.IsNullOrEmpty(objAddresscomp.FieldValue("Comp_Name").ToString()))
            header += "\t\t\t<agency_name>" + objAddresscomp.FieldValue("Comp_Name") + "</agency_name>\n";
            else header += "\t\t\t<agency_name>" + "" + "</agency_name>\n";
            string sqlpubcom = "select * from Publications where pblc_PublicationsID = " + pubxml; 
              QuerySelect  obj = GetQuery();
            obj.SQLCommand = sqlpubcom;
            obj.ExecuteReader();
            if (!string.IsNullOrEmpty(obj.FieldValue("pblc_Commision").ToString()))
            header += "\t\t\t<commission>0." + obj.FieldValue("pblc_Commision") + "</commission>\n";
            else header += "\t\t\t<commission>" + "" + "</commission>\n";
            header += "\t\t</customer>\n";
    
        }
    }
}
