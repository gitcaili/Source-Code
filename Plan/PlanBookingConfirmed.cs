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
using System.Globalization;

namespace NZPACRM.Plan
{
    public class PlanBookingConfirmed : Web
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
        string sCampaign = "";
        string sStatus = "";
        string companyID = "";
        string bookingID = "";
        string sNewsWorksRef = "";
        string sBilledBy = "";
        string bookRefe = "";
        string sAgency = "";
        string sAdvertiser = "";
        string sAgencyContact = "";
        string sQuoteVersionRef = "";
        string sCreationDate = "";
        string sCreatedBy = "";
        string mergedFilePath = "";
        string orgFilePath = "";
        int max = 0;

        CRMHelper crmHelperObj = new CRMHelper();

        public PlanBookingConfirmed()
            : base()
        {           
            if (!String.IsNullOrEmpty(Dispatch.EitherField("Key37")))
            {
                sBookingID = Dispatch.EitherField("Key37");
            }
            //set agency id & booking reference number 
            Record recBook = FindRecord("booking", "book_bookingid=" + sBookingID);
            if (!recBook.Eof())
            {
                companyID = recBook.GetFieldAsString("book_agency");
                bookRefe = recBook.GetFieldAsString("book_reference");
               // AddContent("companyid=" + companyID);
               // AddContent("book refe=" + bookRefe);
            }            
        }
        public override void BuildContents()
        {   
            ReadExistingExcel();
            getSMTPDetails();
            getSMPTPassword();
            SendEmail();
        }
        public void ReadExistingExcel()
        {
            //string path = @"E:\Projects\ETL914-C\Docs\SendToSelf.xls";
            orgFilePath = GetLibraryPath() + @"Plan\BookingConfirmed\DN.xls";

            string checkPathGenTemplate = GetLibraryPath() + @"Plan\BookingConfirmed\GeneratedTemplate";

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
                    AddContent("Set Agency Name");
                    SetAgency();
                   // AddContent("set detail grid value");
                    SetDetailGrid();
                }
                catch (Exception ex)
                {
                 //   AddContent(ex.Message.ToString());
                }
                string date = "";
                date = DateTime.Now.Date.ToShortDateString().Replace(@"/", "-");
                mergedFilePath = GetLibraryPath() + @"Plan\BookingConfirmed\DispatchNoticeResult_" + max + ".xls";
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
                    //AddContent("In-ReadExistingExcel - Catch");
                    //   AddContent(ex.Message.ToString());
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
            string strSQL2 = "select * from  booking where book_bookingId=" + sBookingID + "";
            QuerySelect objAddressRec2 = GetQuery();
            objAddressRec2.SQLCommand = strSQL2;
            objAddressRec2.ExecuteReader();

            string strSQL3 = "select * from  person  where pers_personid=" + objAddressRec2.FieldValue("book_contact") + "";
            QuerySelect objAddressRec3 = GetQuery();
            objAddressRec3.SQLCommand = strSQL3;
            objAddressRec3.ExecuteReader();

            string strSQL4 = "select * from  users  where user_userid=" + objAddressRec2.FieldValue("book_createdby") + "";
            QuerySelect objAddressRec4 = GetQuery();
            objAddressRec4.SQLCommand = strSQL4;
            objAddressRec4.ExecuteReader();

            string strSQL5 = "Select * from client where client_ClientId= '" + objAddressRec2.FieldValue("book_client") + "'";
            QuerySelect objAddressRec5 = GetQuery();
            objAddressRec5.SQLCommand = strSQL5;
            objAddressRec5.ExecuteReader();

            sAdvertiser = "";
            if (!objAddressRec5.Eof()) sAdvertiser = objAddressRec5.FieldValue("client_name");
            sAgencyContact = objAddressRec3.FieldValue("pers_firstname") + " " + objAddressRec3.FieldValue("pers_lastname");
            //sQuoteVersionRef = objBookingList.GetFieldAsString("");
            sCreationDate = objAddressRec.FieldValue("Comp_CreatedDate");
            if (!string.IsNullOrEmpty(objAddressRec2.FieldValue("book_CampaignSummary"))) sCampaign = (objAddressRec2.FieldValue("book_CampaignSummary"));

            sCreatedBy = objAddressRec4.FieldValue("user_firstname") + " " + objAddressRec4.FieldValue("user_lastname");
            sBilledBy = objAddressRec.FieldValue("book_billedby");
            sNewsWorksRef = objAddressRec2.FieldValue("book_Reference");
            sQuoteVersionRef = objAddressRec2.FieldValue("book_costingversion");

         //   sCampaign = objAddressRec.FieldValue("book_campaignsummary");
            if (sBilledBy == "Works")
            {
                sBilledBy = "NewsWorks";
            }
            sStatus = objAddressRec2.FieldValue("book_Status");
      
            mWSheet1.Cells[20, 10] = sAgency;
            mWSheet1.Cells[21, 10] = sAgencyContact;
            mWSheet1.Cells[22, 10] = sAdvertiser;
            mWSheet1.Cells[23, 10] = sCampaign;
            mWSheet1.Cells[24, 10] = sBilledBy;
            mWSheet1.Cells[20, 63] = sCreationDate;
            mWSheet1.Cells[21, 63] = sCreatedBy;
            mWSheet1.Cells[22, 63] = sQuoteVersionRef;
            mWSheet1.Cells[23, 63] = sNewsWorksRef;
            mWSheet1.Cells[24, 63] = sStatus;
        }
        
        private void SetDetailGrid()
        {
            int end = 53;
            int j = 29;
            string sCode = "";
            AddContent("HG");
            string strSQL = "select pblc_Name,pblc_CreatedBy,pblc_CreatedDate,pblc_commision, pblc_Status book_reference,pnbr_pnbr_planbuilderid,pnbr_action, pnbr_days, pnbr_discount, pnbr_color, pnbr_ratecard, pnbr_cost, pnbr_size, pnbr_sections,*";


            strSQL += " from  Booking left  join Planbuilder on book_BookingID=pnbr_plan left join Publications on pnbr_publications = pblc_PublicationsID left JOIN Sections on sctn_Sctn_sectionid=pnbr_sections ";

            strSQL += " where book_BookingID=" + sBookingID + "and pnbr_deleted is null";
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)mWSheet1.Cells[46, 1];
            Microsoft.Office.Interop.Excel.Range range2 = (Microsoft.Office.Interop.Excel.Range)mWSheet1.Cells[45, 1];
            Microsoft.Office.Interop.Excel.Range RngToCopy = range2.EntireRow;

            QuerySelect objPblcnRec = GetQuery();
            objPblcnRec.SQLCommand = strSQL;
            objPblcnRec.ExecuteReader();

            int count = 0;
            while (!objPblcnRec.Eof())
            {
                count++;
                objPblcnRec.Next();
            }

            if (count > 24)
            {
                int insert = count - 24;
                for (int k = 0; k < insert; k++)
                {
                    Microsoft.Office.Interop.Excel.Range row = range.EntireRow;
                    row.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, RngToCopy.Copy(Type.Missing));
                    end++;
                }
            }

            objPblcnRec.SQLCommand = strSQL;
            objPblcnRec.ExecuteReader();
            AddContent("JO");

            if (!objPblcnRec.Eof())
            {
                AddContent("Help");
                int Count = 0;
                while (!objPblcnRec.Eof())
                {
                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_action").ToString()))
                    {
                  //      mWSheet1.Cells[j, 4] = objPblcnRec.FieldValue("pnbr_action").ToString();
                    }
                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                    {
                        mWSheet1.Cells[j, 4] = objPblcnRec.FieldValue("pblc_Name").ToString();
                    }
                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_MaterialDelivery").ToString()))
                    {
                        mWSheet1.Cells[j, 76] = objPblcnRec.FieldValue("pblc_MaterialDelivery").ToString();
                    }

                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("sctn_name").ToString()))
                    {
                        mWSheet1.Cells[j, 12] = objPblcnRec.FieldValue("sctn_name").ToString();
                    }
                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_Note").ToString()))
                    {
                        mWSheet1.Cells[j, 22] = objPblcnRec.FieldValue("pnbr_Note").ToString();
                    }
                 
                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_size").ToString()))
                    {
                        mWSheet1.Cells[j, 52] = objPblcnRec.FieldValue("pnbr_standardsize").ToString();
                    }
                  
                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_color").ToString()))
                    {
                        string color = objPblcnRec.FieldValue("pnbr_color").ToString();
                        if (color == "Color" || color == "COLOR") color = "Colour";
                        if (color == "NoColor" || color == "NOCOLOR") color = "Mono";
                        mWSheet1.Cells[j, 82] = color;
                    }
                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_keynumber").ToString()))
                    {
                        mWSheet1.Cells[j, 29] = objPblcnRec.FieldValue("pnbr_keynumber").ToString();
                    }
                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_caption").ToString()))
                    {
                        mWSheet1.Cells[j, 34] = objPblcnRec.FieldValue("pnbr_caption").ToString();
                    }
                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_days").ToString()))
                    {
                        mWSheet1.Cells[j, 39] = GetDays(objPblcnRec.FieldValue("pnbr_days").ToString());
                    }
                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_date").ToString()))
                    {
                        DateTimeFormatInfo format = new DateTimeFormatInfo();
                        format.ShortDatePattern = "dd-MM-yy";
                        format.DateSeparator = "-";
                        DateTime date = Convert.ToDateTime(objPblcnRec.FieldValue("pnbr_date"));
                        mWSheet1.Cells[j, 46] = Convert.ToDateTime(date, format);
                    }
                    if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_plan").ToString()))
                    {
                       
                        string ratesql = "Select * from RatesCard where rate_RatesCardid = '" + objPblcnRec.FieldValue("pnbr_ratecard").ToString() + "'";

                        QuerySelect objRec = GetQuery();
                        objRec.SQLCommand = ratesql;
                        objRec.ExecuteReader();
                        if (!objRec.Eof())
                        {
                         //   AddContent(objPblcnRec.FieldValue("pnbr_date").ToString());
                           AddContent("FREAK");
                            DateTime insert = objPblcnRec.FieldValueAsDate("pnbr_date");
                            AddContent("PASS");
                          //  string before = objRec.FieldValue("rate_BookinDeadlineDays").ToString().Substring(0, 1);
                           // int bre = Int32.Parse(before);

                          //  insert.AddDays((bre * -1));
                            //mWSheet1.Cells[j, 81] = insert.ToShortDateString() + " " + objRec.FieldValue("rate_BookingDeadlineTime").ToString();
                            int widthlim = -1;
                            if (objRec.FieldValue("rate_Size").ToString() != "Standard" || objRec.FieldValue("rate_standardsize").Contains("Module") || objRec.FieldValue("rate_standardsize").Contains("cm") || objRec.FieldValue("rate_standardsize").Contains("Cm") || objRec.FieldValue("rate_standardsize").Contains("CM"))
                            {
                                for (int i = 1; i < 13; i++)
                                {
                                    string num = objRec.FieldValue("rate_width" + i.ToString());
                                    //  AddContent("STRING" + i + " " + num);
                                    double lim = Double.Parse(num);
                                    if (lim == 0)
                                    {
                                        widthlim = i - 1;
                                        break;
                                    }
                                }
                                if (widthlim == -1)
                                {
                                    widthlim = 12;
                                }
                            }
                            else
                            {
                                string num = objRec.FieldValue("rate_SetSizesWidth");
                                int lim = Int32.Parse(num);
                                widthlim = lim;
                            }
                              AddContent("HAPPY");
                            float h;
                            if (objRec.FieldValue("rate_Size").ToString() == "Standard" && !objRec.FieldValue("rate_standardsize").Contains("Module") && !objRec.FieldValue("rate_standardsize").Contains("cm") && !objRec.FieldValue("rate_standardsize").Contains("Cm") && !objRec.FieldValue("rate_standardsize").Contains("CM"))
                                h = float.Parse(objRec.FieldValue("rate_SetSizesHeight"));
                            else h = float.Parse(objPblcnRec.FieldValue("pnbr_height"));
                            //  h = h / 10;
                               AddContent(h.ToString());
                            string specs = "";
                            if (objRec.FieldValue("rate_Size").ToString() == "Standard" && !objRec.FieldValue("rate_standardsize").Contains("Module") && !objRec.FieldValue("rate_standardsize").Contains("cm") && !objRec.FieldValue("rate_standardsize").Contains("Cm") && !objRec.FieldValue("rate_standardsize").Contains("CM"))
                            {
                                specs = h.ToString() + "x" + widthlim.ToString();
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_width").ToString()))
                                    specs = h.ToString() + "x" + float.Parse(objPblcnRec.FieldValue("pnbr_width"));
                            }
                               AddContent(specs);
                            mWSheet1.Cells[j, 58] = specs;
                            string specsize;
                            if (objRec.FieldValue("rate_Size").ToString() == "Standard" && !objRec.FieldValue("rate_standardsize").Contains("Module") && !objRec.FieldValue("rate_standardsize").Contains("cm") && !objRec.FieldValue("rate_standardsize").Contains("Cm") && !objRec.FieldValue("rate_standardsize").Contains("CM"))
                            {
                                 specsize = objRec.FieldValue("rate_height") + "x" + objRec.FieldValue("rate_width");
                            }
                            else
                            {
                                specsize = (h * 10).ToString() + "x" + objRec.FieldValue("rate_width" + widthlim.ToString()); 
                            }
                              mWSheet1.Cells[j, 64] = specsize;
                           AddContent("CATSds");
                            insert = objPblcnRec.FieldValueAsDate("pnbr_date");

                            //string before2 = objRec.FieldValue("rate_MaterialDeadlinDays").ToString().Substring(0, 1);
                            int bre2 = MaterialDays(objRec.FieldValue("rate_MaterialDeadlinDays").ToString(), GetDays(objPblcnRec.FieldValue("pnbr_days").ToString()));
                            AddContent("AM" + bre2 + "BM");

                          insert =  insert.AddDays((bre2 * -1));
                            mWSheet1.Cells[j, 70] = insert.ToShortDateString() + " " + objRec.FieldValue("rate_MaterialDeadlineTime").ToString(); 


                        }
                    }

                    j++;
                    Count++;
                    objPblcnRec.Next();
                }
                mWSheet1.Cells[end, 78] = "Total Insertions: " + Count.ToString();
            }
        }
        private string GetDays(string objPnbrDay)
        {
            string sDay = "";
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
                        return objPnbrDay;
                        break;
                }
            }
            sDay = sDay.Replace(",", "");
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
              //  AddContent(ex.Message.ToString());
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
              //  AddContent(ex.Message.ToString());
                AddUrlButton("Continue", "Continue.gif", UrlDotNet("NZPACRM", "RunPlanDataPage" + "&book_bookingid=" + sBookingID + ""));
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

                wrkflwResult = crmHelperObj.ProgressWorkflow(sBookingID, "Booking", "Booking Workflow", "Dispatched Sent");
                
                if (wrkflwResult)
                {
                    AddInfo("Workflow progressed successfully to Booking Attached");

                    crmHelperObj.SetStageStatus("Booking", sBookingID, "BookingConfirmed", "InProgress");
                }
                else
                {
                    AddError("Error Occurred during worlflow progress");
                }
                string sURL = Url("432") + "&book_bookingid=" + sBookingID;
                AddUrlButton("Continue", "Continue.gif", UrlDotNet("NZPACRM", "RunPlanDataPage" + "&book_bookingid=" + sBookingID + ""));
            }
            catch (Exception ex)
            {
                AddUrlButton("Continue", "Continue.gif", UrlDotNet("NZPACRM", "RunPlanDataPage" + "&book_bookingid=" + sBookingID + ""));
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
            string sSQL = " select * from Custom_EmailAddress (nolock) where emse_displayname='System Administrator' and EmSe_Deleted is null";
            QuerySelect sQueryObj = GetQuery();
            
            sQueryObj.SQLCommand = sSQL;
            sQueryObj.ExecuteReader();
            if (!sQueryObj.Eof())
            {
                sFromEmailAddress = sQueryObj.FieldValue("EmSe_EmailAddress").ToString();
            }
            return sFromEmailAddress;
        }


        public int MaterialDays(string MDT, string Day)
        {
            if (MDT.Contains("prior")) MDT = MDT.Replace(" prior", "");
            int MDTValue = 0;
            int NewDay = 0;

            if (MDT.Contains("Monday")) MDTValue = 1;
            else if (MDT.Contains("Tuesday")) MDTValue = 2;
            else if (MDT.Contains("Wednesday")) MDTValue = 3;
            else if (MDT.Contains("Thursday")) MDTValue = 4;
            else if (MDT.Contains("Friday")) MDTValue = 5;
            else if (MDT.Contains("Saturday")) MDTValue = 6;
            else if (MDT.Contains("Sunday")) MDTValue = 7;
            else {
                AddContent("Out of here" + MDT);
                return Int32.Parse(MDT.Substring(0, 1)); }

            if (Day.Contains("Monday")) NewDay = 1;
            if (Day.Contains("Tuesday")) NewDay = 2;
            if (Day.Contains("Wednesday")) NewDay = 3;
            if (Day.Contains("Thursday")) NewDay = 4;
            if (Day.Contains("Friday")) NewDay = 5;
            if (Day.Contains("Saturday")) NewDay = 6;
            if (Day.Contains("Sunday")) NewDay = 7;

            int difference = MDTValue - NewDay;
            if (difference < 0) difference += 7;
            AddContent(difference + "FO" + MDTValue);
            return difference;
        }
    }
}
