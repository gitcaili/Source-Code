using System;

using Sage.CRM.WebObject;

using Sage.CRM.Data;
using NZPACRM.Common;

using System.IO;

using System.Diagnostics;
using System.Text;
using System.Reflection;
using System.Net.Mail;
using System.Net;
using System.Globalization;

namespace NZPACRM.Plan
{
    public class PlanOtherImport : DataPageNew
    {
        Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        Microsoft.Office.Interop.Excel.Application oXL;
        CRMHelper objCRM = new CRMHelper();
        string servername = "";
        string smtpport = "";
        string smtpusername = "";
        string smtppwd = "";
        string userID = "";
        string LogfileName = "";
        string shttpURL = "";
        Boolean revised = false;
        string agentcy = "";
        Boolean newplan = true;
        //   OleDbConnection oleExcelConnection = default(OleDbConnection);
        public PlanOtherImport()
            : base("Booking", "book_BookingID", "")
        {
            string CurrUser = CurrentUser.UserId.ToString(); // the current user

            #region get Http from url
            try
            {
                string s = Dispatch.ServerVariable("HTTP_REFERER");
                char[] cSplit = { '/' };
                string[] sHTTP = s.Split(cSplit);

                if (!String.IsNullOrEmpty(sHTTP[0]))
                    shttpURL = sHTTP[0]; // checking to see if it is security or something?


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

            string sHostName = Dispatch.Host;

            string sInstallName = Dispatch.InstallName;
            #endregion

        }

        //public OleDbConnection oleExcelConnection = default(OleDbConnection);
        public override void BuildContents()
        {
            int bid2 = -1;
            
 
            try
            {
                //OleDbConnection conexel;
                //AddContent("<script type='text/javascript' src='../CustomPages/Booking/ClientFuncs.js'></script>");
                AddContent("<script type='text/javascript' src='../js/custom/ClientFuncs.js'></script>");
                string sSuccessErrorMessage = "";
                int iFailedCount = 0;
                int InsertCount = 0;
                int iDupRecord = 0;
                string isDatavalid = "";

                #region Adding Html Form
                AddContent(HTML.Form());
                #endregion

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

                sHTML += HTML.TableData("<span style='margin: 1em 3em 10px 55em;font-Size:14px;'><a href='" + sFileURL + "NZPAImport/Templates/PlanTemplate.xlsx' class='PANEREPEAT'><u>Download Template</u></a></span>");
                sHTML += HTML.EndTable();
                sHTML += "<BR>";
                AddContent(HTML.Box("Import Plan Data", sHTML));

                AddContent("<BR><BR>" + HTML.Box("File", "<br>&nbsp;&nbsp;<input type='file' id='fileupload' name='pic' size='70'>&nbsp;<input type='BUTTON' class='Edit'value='Import'name='upload' onclick='javascipt:CheckFileNew();'></br></br>"));

                string backURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                backURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=1650&Mode=1&CLk=T&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data Management";
                AddUrlButton("Back", "prevcircle.gif", backURL);

                AddContent(HTML.InputHidden("HIDDEN_LibraryPath", GetLibraryPath().Replace("Library", "")));
                AddContent(HTML.InputHidden("HIDDEN_FilePath", ""));
                AddContent(HTML.InputHidden("HIDDEN_Save", ""));
                AddContent(HTML.InputHidden("HIDDEN_FileName", ""));
                AddContent(HTML.InputHidden("HIDDEN_FilePathChrome", ""));
                AddContent(HTML.InputHidden("HIDDEN_browser", ""));

                if (!String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_Save")))
                {
                    if (Dispatch.ContentField("HIDDEN_Save") == "Save")
                    {
                        AddContent(Dispatch.ContentField("HIDDEN_FilePath"));
                        string SavedFilePath = "";

                        if (String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_FilePathChrome")))
                        {
                            SavedFilePath = SaveRateCardRLocation();
                        }
                        else
                        {
                            SavedFilePath = GetLibraryPath() + @"\RateCard\" + Dispatch.ContentField("HIDDEN_FileName");
                        }

                        //  DataTable dt = new DataTable();
                        oXL = new Microsoft.Office.Interop.Excel.Application();
                        oXL.DisplayAlerts = false;

                        try
                        {
                            mWorkBook = oXL.Workbooks.Open(SavedFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                            var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; data source=" + SavedFilePath + "; Extended Properties='Excel 12.0;HDR=NO;IMEX=1;'", SavedFilePath);
                            mWorkSheets = mWorkBook.Worksheets;
                            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item(1);
                            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;

                            //   conexel = new OleDbConnection(connectionString);
                            //  OleDbCommand cmdExcel = new OleDbCommand();
                            //var adapter = new OleDbDataAdapter();
                            //cmdExcel.Connection = conexel;
                            // conexel.Open();
                            // cmdExcel.CommandText = "SELECT * FROM [sheet1$]";
                            // adapter.SelectCommand = cmdExcel;

                            // var adapter = new OleDbDataAdapter("SELECT * FROM [sheet1$]", connectionString);

                            //var ds = new DataSet();
                            //adapter.Fill(ds);

                            //   adapter.Fill(dt);

                            // conexel.Close();
                    //        AddContent("START");
                            agentcy = "";
                            string contact = "";
                            string advertiser = "";
                            string Camp = "";

                            if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[12, 5] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                agentcy = (string)(mWSheet1.Cells[12, 5] as Microsoft.Office.Interop.Excel.Range).Value;
                        //    AddContent("agent" + agentcy);
                            //copyme = agentcy;
                            if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[13, 5] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                contact = (string)(mWSheet1.Cells[13, 5] as Microsoft.Office.Interop.Excel.Range).Value;

                            if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[14, 5] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                advertiser = (string)(mWSheet1.Cells[14, 5] as Microsoft.Office.Interop.Excel.Range).Value;

                            if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[15, 5] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                Camp = (string)(mWSheet1.Cells[15, 5] as Microsoft.Office.Interop.Excel.Range).Value;

                            string pid = "";
                            string cid = "";
                            string comp = "";
                         //   string comtype = "";
                            DateTime d1 = DateTime.Now;

                            int bid = 0;
                           
                         //   AddContent("CAT" + contact);

                            char[] delimiterChars2 = { ' ', ',', '.', ':', '\t' };
                            // AddContent(");
                           
                            string name = "" + agentcy + " " + d1.Year.ToString();
                          
                          
                            //  DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;

                            //  Calendar cal = dfi.Calendar;
                            // AddContent((cal.GetWeekOfYear(d1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek)).ToString() + "JODM");
                            string strSQL = "select * from Booking where book_Agency =(Select comp_CompanyId from Company where Comp_Name like '" + agentcy + "' and comp_deleted is null) and book_CampaignSummary = '" + Camp + "' and book_Deleted is null and book_status <> 'Closed'";
                            QuerySelect objCurrencyRec = GetQuery();
                            QuerySelect emailLinkRec = GetQuery();
                         //   AddContent(strSQL);
                            objCurrencyRec.SQLCommand = strSQL;
                            objCurrencyRec.ExecuteReader();
                            if (!objCurrencyRec.Eof())
                            {
                                string sbid = objCurrencyRec.FieldValue("book_Bookingid").ToString();
                                bid = Int32.Parse(sbid);
                                bid2 = bid;
                                newplan = false;

                                string sdoc = objCurrencyRec.FieldValue("book_docversion").ToString();
                                int docnew = Int32.Parse(sdoc) + 1;

                                

                                Record version = FindRecord("Booking", "book_Bookingid = '" + sbid + "'");
                                Record newBook = new Record("Booking");
                                string[] parms = new string[] { "book_Name", "book_Agency", "book_Contact", "book_Client", "book_CampaignSummary", "book_costingversion", "book_status", "book_agencycode", "book_docversion", "booK_reference", "book_version" };
                                string[] values = new string[] { objCurrencyRec.FieldValue("book_name").ToString(), objCurrencyRec.FieldValue("book_agency").ToString(), objCurrencyRec.FieldValue("book_contact").ToString(), objCurrencyRec.FieldValue("book_client").ToString(), objCurrencyRec.FieldValue("book_campaignsummary").ToString(), objCurrencyRec.FieldValue("book_costingversion").ToString(), "InProgress", objCurrencyRec.FieldValue("book_agencycode").ToString(), docnew.ToString(), objCurrencyRec.FieldValue("book_reference").ToString(), objCurrencyRec.FieldValue("book_version").ToString() };


                                for (int i = 0; i < parms.Length; i++)
                                {
                                    newBook.SetField(parms[i].ToString(), values[i].ToString());

                                }

                                string[] oldparms = new string[] { "book_status" };
                                string[] oldvalues = new string[] { "Closed" };

                                for (int i = 0; i < oldparms.Length; i++)
                                {
                                    version.SetField(oldparms[i].ToString(), oldvalues[i].ToString());

                                }

                                newBook.SaveChanges();
                                version.SaveChanges();
                                int newbid = newBook.RecordId;
                                copyplanbuilders(bid.ToString(), newbid.ToString());
                                bid = newbid;

                                string wfid = "13";
                                string curren = "10234";
                                string currstate = "55";


                                string WriteEntity = "WorkFlowInstance";
                                parms = new string[] { "WkIn_WorkflowId", "WkIn_CurrentEntityId", "WkIn_CurrentRecordId", "WkIn_CurrentStateId" };
                                values = new string[] { wfid, curren, bid.ToString(), currstate };
                                Record WFRecord = new Record(WriteEntity);
                                for (int i = 0; i < parms.Length; i++)
                                {
                                    WFRecord.SetField(parms[i].ToString(), values[i].ToString());

                                }
                                WFRecord.SaveChanges();
                                int wid = WFRecord.RecordId;
                                Record r = FindRecord("Booking", "book_BookingId = " + bid.ToString());
                                parms = new string[] { "book_WorkflowId" };
                                values = new string[] { wid.ToString() };
                                for (int i = 0; i < parms.Length; i++)
                                {

                                    r.SetField(parms[i].ToString(), values[i].ToString());

                                }
                                r.SaveChanges();




                                revised = true;

                            }

                            String agentcode = "";
                          //  AddContent(bid.ToString());
                            if (newplan)
                            {
                                //AddContent(" P.O.P ");
                                string strSQLco = "Select * from Company where Comp_Name ='" + agentcy + "' and comp_deleted is null";
                                QuerySelect objCurrencyCo = GetQuery();
                                // AddContent(strSQL2);
                                objCurrencyCo.SQLCommand = strSQLco;
                                objCurrencyCo.ExecuteReader();
                           //     AddContent(" P.O.P2 ");


                                if (!objCurrencyCo.Eof())
                                {
                                    comp = objCurrencyCo.FieldValue("comp_CompanyId");
                                    agentcode = objCurrencyCo.FieldValue("comp_AgencyCode");
                                }

                                string strSQL2 = "select * from Email where Emai_EmailAddress = '" + contact + "'";
                                objCurrencyRec = GetQuery();
                            //    AddContent(strSQL2);
                                objCurrencyRec.SQLCommand = strSQL2;
                                objCurrencyRec.ExecuteReader();

                                if (!objCurrencyRec.Eof())
                                {
                                    string SQLel = "select * from EmailLink where ELink_EmailID ='" + objCurrencyRec.FieldValue("Emai_EmailId").ToString() + "'";

                                    emailLinkRec.SQLCommand = SQLel;
                                    emailLinkRec.ExecuteReader();
                                    if (!emailLinkRec.Eof()) { pid = emailLinkRec.FieldValue("ELink_RecordID"); }
                                    //get the current person     
                                }

                                string SQLcl = "select * from Client where client_Name like '%" + advertiser + "%'";
                                QuerySelect Client = GetQuery();
                                Client.SQLCommand = SQLcl;
                                Client.ExecuteReader();
                                if (!Client.Eof())
                                {
                                    cid = Client.FieldValue("client_ClientId");
                                //    comtype = Client.FieldValue("client_Type");
                                }
                                int iVersionNo = 0;
                                string sSQL = "select ISNULL( MAX(book_version),0) as Version from Booking";
                                QuerySelect sQueryObj = GetQuery();
                                sQueryObj.SQLCommand = sSQL;
                                sQueryObj.ExecuteReader();
                                
                                if (!sQueryObj.Eof())
                                {
                                    iVersionNo = Convert.ToInt32(sQueryObj.FieldValue("Version")) + 1;
                                }

                                string WriteEntity = "Booking";

                                string[] parms = new string[] { "book_Name", "book_Agency", "book_Contact", "book_Client", "book_CampaignSummary", "book_costingversion", "book_status", "book_agencycode", "book_docversion", "booK_reference","book_version" };
                                string[] values = new string[] { name, comp, pid, cid, Camp, "", "InProgress", agentcode, "1",GenerateSequecenumber(""),iVersionNo.ToString() };

                                Record bookRecord = new Record(WriteEntity);
                                for (int i = 0; i < parms.Length; i++)
                                {
                                    bookRecord.SetField(parms[i].ToString(), values[i].ToString());

                                }

                                bookRecord.SaveChanges();
                                bid = bookRecord.RecordId;
                                string wfid = "13";
                                string curren = "10234";
                                string currstate = "55";

                                WriteEntity = "WorkFlowInstance";
                                parms = new string[] { "WkIn_WorkflowId", "WkIn_CurrentEntityId", "WkIn_CurrentRecordId", "WkIn_CurrentStateId" };
                                values = new string[] { wfid, curren, bid.ToString(), currstate };
                                Record WFRecord = new Record(WriteEntity);
                                for (int i = 0; i < parms.Length; i++)
                                {
                                    WFRecord.SetField(parms[i].ToString(), values[i].ToString());

                                }
                                WFRecord.SaveChanges();
                                int wid = WFRecord.RecordId;
                                Record r = FindRecord("Booking", "book_BookingId = " + bid.ToString());
                                parms = new string[] { "book_WorkflowId" };
                                values = new string[] { wid.ToString() };
                                for (int i = 0; i < parms.Length; i++)
                                {

                                    r.SetField(parms[i].ToString(), values[i].ToString());

                                }
                                r.SaveChanges();

                                bid2 = bid;
                                //  AddContent("IT IS FINISHED " + bid);
                            }
                            string SQLcl2 = "select * from Client where client_Name like '%" + advertiser + "%'";
                            QuerySelect Client2 = GetQuery();
                            Client2.SQLCommand = SQLcl2;
                            Client2.ExecuteReader();
                            if (!Client2.Eof())
                            {
                                cid = Client2.FieldValue("client_ClientId");
                              //  comtype = Client2.FieldValue("client_Type");
                            }
                            //comtype = comissionconvert(comtype);
                            // comtype = "NonCommission";
                            //  AddContent(mWSheet1.UsedRange.Rows.Count.ToString());

                            // mWSheet1.Rows.ClearFormats();
                            //  AddContent(mWSheet1.UsedRange.Rows.Count.ToString());
                            for (int j = 22; j < mWSheet1.UsedRange.Rows.Count + 22; j++)
                            //for (int j = 22; j < 40; j++)
                            {
                                  AddContent("Welcome +" + (mWSheet1.UsedRange.Rows.Count + 2).ToString());

                                if (string.IsNullOrEmpty((string)(mWSheet1.Cells[j, 2] as Microsoft.Office.Interop.Excel.Range).Value) && !string.IsNullOrEmpty((string)(mWSheet1.Cells[j + 1, 2] as Microsoft.Office.Interop.Excel.Range).Value)) j++;
                                if (string.IsNullOrEmpty((string)(mWSheet1.Cells[j, 5] as Microsoft.Office.Interop.Excel.Range).Value)) break;
                                  AddContent(j.ToString());
                                string eventPlan = "";
                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[j, 16] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                    eventPlan = (string)(mWSheet1.Cells[j, 2] as Microsoft.Office.Interop.Excel.Range).Value; ;
                                //  string quoteR = (string)(mWSheet1.Cells[j, 3] as Microsoft.Office.Interop.Excel.Range).Value;
                                string bookedby = (string)(mWSheet1.Cells[j, 4] as Microsoft.Office.Interop.Excel.Range).Value;
                                //  AddContent("BOOKED");
                                if (!string.IsNullOrEmpty(bookedby))
                                {
                                    if (bookedby.Equals("NW"))
                                        bookedby = "NewsWorks";
                                }
                                //     AddContent("BOOKED2");
                                string Commission = (string)(mWSheet1.Cells[j, 5] as Microsoft.Office.Interop.Excel.Range).Value;

                                if (Commission == "Non Commission") Commission = "NonCommission";
                                // AddContent("BOOKED3");
                                string status = (string)(mWSheet1.Cells[j, 6] as Microsoft.Office.Interop.Excel.Range).Value;

                                if (string.IsNullOrEmpty(status)) break;
                                status = status.ToLower();
                                // AddContent("BOOKED5");
                                string publication = (string)(mWSheet1.Cells[j, 7] as Microsoft.Office.Interop.Excel.Range).Value;
                                //    AddContent("BOOKED");
                                string section = (string)(mWSheet1.Cells[j, 8] as Microsoft.Office.Interop.Excel.Range).Value;
                                string subsection = (string)(mWSheet1.Cells[j, 9] as Microsoft.Office.Interop.Excel.Range).Value;
                                //  AddContent("BOOKED");
                                string day = (string)(mWSheet1.Cells[j, 10] as Microsoft.Office.Interop.Excel.Range).Value;
                                //   AddContent("DAY");
                                double insertiondate2 = (double)(mWSheet1.Cells[j, 11] as Microsoft.Office.Interop.Excel.Range).Value2;
                                DateTime insertiondate = DateTime.FromOADate(insertiondate2);
                                string size = "";
                                string size2 = "";
                                //  AddContent("Insert");
                                // string Commission = (string)(mWSheet1.Cells[j, 12] as Microsoft.Office.Interop.Excel.Range).Value;
                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[j, 12] as Microsoft.Office.Interop.Excel.Range).Text).Trim())) {

                                    size = (string)(mWSheet1.Cells[j, 12] as Microsoft.Office.Interop.Excel.Range).Value;
                                    size2 = StandardSizeCode(size); }
                                // AddContent("HELl");
                                string specs = (string)(mWSheet1.Cells[j, 13] as Microsoft.Office.Interop.Excel.Range).Value;
                                //   AddContent("SPECS");
                                string color = (string)(mWSheet1.Cells[j, 14] as Microsoft.Office.Interop.Excel.Range).Value;
                                //     AddContent("COOOROR");
                                //string cost = (string)(mWSheet1.Cells[j, 15] as Microsoft.Office.Interop.Excel.Range).Value;
                                if (color == "Colour" || color == "COLOUR") color = "COLOR";
                                else color = "NoColor";
                                // AddContent("BEES");
                                string key = "";
                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[j, 15] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                    key = (string)(mWSheet1.Cells[j, 15] as Microsoft.Office.Interop.Excel.Range).Value.ToString();

                                // string key = (string)(mWSheet1.Cells[j, 16].Text.to
                                //   AddContent("KEY");
                                string caption = "";
                                string special = "";
                                string dispatched = "";
                                string valid = "";
                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[j, 16] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                    caption = (string)(mWSheet1.Cells[j, 16] as Microsoft.Office.Interop.Excel.Range).Value;
                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[j, 17] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                    special = (string)(mWSheet1.Cells[j, 17] as Microsoft.Office.Interop.Excel.Range).Value;
                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[j, 19] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                    dispatched = (string)(mWSheet1.Cells[j, 19] as Microsoft.Office.Interop.Excel.Range).Value;
                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[j, 20] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                    valid = (string)(mWSheet1.Cells[j, 20] as Microsoft.Office.Interop.Excel.Range).Value;
                                string rateday = "";
                                if (day.Equals("Thursday"))
                                {
                                    rateday = "Thrusday";
                                }
                                else rateday = day;
                                string dayshort = "";
                                //    AddContent("READ");
                                if (day.Equals(string.Empty) == false) dayshort = day.Substring(0, 3);
                                if (day.Equals("Thrusday")) dayshort = "Thu";

                                if (publication.Equals("HAWKE'S BAY TODAY")) publication = "HAWKE";
                                //   int weeknow = cal.GetWeekOfYear(d1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);

                                char[] delimiterChars = { ' ', ',', '.', ':', '\t' };

                                if (1 == 1)
                                {

                                    if (1 == 1)
                                    {
                                        string strSQLco = "Select * from Company where Comp_Name ='" + agentcy + "' and comp_deleted is null";
                                        QuerySelect objCurrencyCo = GetQuery();
                                        objCurrencyCo.SQLCommand = strSQLco;
                                        objCurrencyCo.ExecuteReader();
                                        if (!objCurrencyCo.Eof())
                                        {
                                            comp = objCurrencyCo.FieldValue("comp_CompanyId");
                                            agentcode = objCurrencyCo.FieldValue("comp_AgencyCode");
                                        }

                                        //string name = "" + agentcy;
                                        // // string WriteEntity = "Booking";
                                        // string[] parms = new string[] { "book_Name", "book_Agency", "book_Contact", "book_Client", "book_CampaignSummary", "book_status", "book_agencycode","book_reference" };
                                        //  string[] values = new string[] { name, comp, pid, cid, Camp, "InProgress", agentcode,GenerateSequecenumber("")};

                                        // Record bookRecord = new Record(WriteEntity);
                                        //  for (int i = 0; i < parms.Length; i++)
                                        //  {
                                        //     bookRecord.SetField(parms[i].ToString(), values[i].ToString());

                                        //}
                                        //  bookRecord.SaveChanges();
                                        // bid = bookRecord.RecordId;
                                        string wfid = "13";
                                        string curren = "10234";
                                        string currstate = "55";

                                        // WriteEntity = "WorkFlowInstance";
                                        //   parms = new string[] { "WkIn_WorkflowId", "WkIn_CurrentEntityId", "WkIn_CurrentRecordId", "WkIn_CurrentStateId" };
                                        //   values = new string[] { wfid, curren, bid.ToString(), currstate };
                                        //  Record WFRecord = new Record(WriteEntity);
                                        //  for (int i = 0; i < parms.Length; i++)
                                        //  {
                                        //      WFRecord.SetField(parms[i].ToString(), values[i].ToString());

                                        //   }
                                        //   WFRecord.SaveChanges();
                                        //   int wid = WFRecord.RecordId;
                                        //   Record r = FindRecord("Booking", "book_BookingId = " + bid.ToString());
                                        //    parms = new string[] { "book_WorkflowId" };
                                        //    values = new string[] { wid.ToString() };
                                        //      for (int i = 0; i < parms.Length; i++)
                                        //      {

                                        //          r.SetField(parms[i].ToString(), values[i].ToString());

                                        //      }
                                        //     r.SaveChanges();
                                        //     bid2 = bid;
                                    }

                                }
                                // AddContent("RAT");
                                string pubid = "";
                                string SQLpubcat = "select * from Publications where pblc_Name like'%" + publication + "%' and pblc_deleted is null";
                                QuerySelect Publ = GetQuery();
                                Publ.SQLCommand = SQLpubcat;
                                Publ.ExecuteReader();
                                if (!Publ.Eof())
                                { pubid = Publ.FieldValue("pblc_PublicationsId"); }
                                string secid = "";
                                string SQLsec = "select * from Sections where sctn_Name like'" + section + "' and sctn_publicationid = '" + pubid + "'";
                                QuerySelect sec = GetQuery();
                                sec.SQLCommand = SQLsec;
                                sec.ExecuteReader();
                                if (!sec.Eof())
                                { secid = sec.FieldValue("sctn_sctn_sectionid"); }

                                string subsecid = "";
                                string SQLsubsec = "select * from SubSection where suse_Name like'" + subsection + "' and suse_section= '" + secid + "'";
                                QuerySelect subsec = GetQuery();
                                subsec.SQLCommand = SQLsubsec;
                                subsec.ExecuteReader();
                                if (!subsec.Eof())
                                { subsecid = subsec.FieldValue("suse_subsectionid"); }

                                string Rateid = "";
                                string rateamount = "";

                                string SQLRate = "Select * from RatesCard where rate_PublicationsID = '" + pubid + "' and rate_commissiontype = '" + Commission + "' and rate_Section = '" + secid + "' and rate_Day like '%" + dayshort + "%' and rate_standardsize ='" + size2 + "' and rate_color = '" + color + "'";

                                sec = GetQuery();
                                sec.SQLCommand = SQLRate;
                                sec.ExecuteReader();
                                AddContent("Testing code");
                              //  AddContent(sec.FieldValue("rate_" + rateday));
                                if (!sec.Eof())
                                {
                                    Rateid = sec.FieldValue("rate_RatesCardID");
                                    rateamount = sec.FieldValue("rate_" + rateday);
                                }
                                string plancheck = "";
                                string dayuse = "";

                                if (day.Equals("Thursday")) dayuse = "Thur";
                                else dayuse = dayshort;
                                AddContent("Testing code");
                                AddContent(status);
                                //  if (secid.Equals("")) { plancheck = "select * from PlanBuilder where pnbr_plan = " + bid + " and pnbr_commissiontype = '" + Commission + "'and pnbr_Publications = " + pubid + " and pnbr_sections is null and pnbr_day = '" + dayuse + "' and pnbr_Deleted is null"; }
                                //  else { plancheck = "select * from PlanBuilder where pnbr_plan = " + bid + " and pnbr_commissiontype = '" + Commission + "'and pnbr_Publications = " + pubid + " and pnbr_sections = '" + secid + "' and pnbr_day like '%" + dayuse + "%' and pnbr_Deleted is null"; }

                                //   QuerySelect plancheckQ = GetQuery();
                                //   plancheckQ.SQLCommand = plancheck;
                                //  plancheckQ.ExecuteReader();

                                //  AddContent("cool");

                                if (status.Equals("cancel") || status.Equals("Cancel") || status.Equals("CANCEL"))
                                {
                                    Record r = null;
                                    string searchfield = "pnbr_pnbr_planbuilderid";

                                    if (revised) searchfield = "pnbr_revised";

                                    string pbid = (string)((mWSheet1.Cells[j, 18] as Microsoft.Office.Interop.Excel.Range).Value.ToString());

                                    r = FindRecord("PlanBuilder", searchfield + " = '" + pbid + "'");
                                    if (r.GetFieldAsString("pnbr_action") == "Booked" || r.GetFieldAsString("pnbr_action") == "Amend" || r.GetFieldAsString("pnbr_action") == "Cancel")
                                    {

                                        string[] prams = { "pnbr_action" };
                                        string[] updat = { "Cancel" };
                                        for (int i = 0; i < prams.Length; i++)
                                        {

                                            r.SetField(prams[0].ToString(), updat[0].ToString());

                                        }
                                    }
                                    else
                                    {
                                        string[] prams = { "pnbr_deleted" };
                                        string[] updat = { "1" };
                                        for (int i = 0; i < prams.Length; i++)
                                        {

                                            r.SetField(prams[0].ToString(), updat[0].ToString());

                                        }
                                    }
                                    r.SaveChanges();
                                    mWSheet1.Cells[j, 18] = r.RecordId.ToString();
                                }



                                else if (status.Equals("change") || status.Equals("Change") || status.Equals("CHANGE"))
                                {
                                    Record r = null;
                                    //    AddContent("DASTYDCSGBUH + " + day);

                                    string searchfield = "pnbr_pnbr_planbuilderid";

                                    if (revised) searchfield = "pnbr_revised";

                                    string pbid = (string)((mWSheet1.Cells[j, 18] as Microsoft.Office.Interop.Excel.Range).Value.ToString());

                                    r = FindRecord("PlanBuilder", searchfield + " = '" + pbid + "'");


                               //     DateTimeFormatInfo format = new DateTimeFormatInfo();
                                 //   format.ShortDatePattern = "dd-MM-yy";
                                   // format.DateSeparator = "-";
                                    //DateTime old = Convert.ToDateTime(objPblcnRec.FieldValue("pnbr_date"));


                                   DateTime old = r.GetFieldAsDateTime("pnbr_date");

                                    if (insertiondate.Year != old.Year || insertiondate.Month != old.Month || insertiondate.Day!= old.Day && (r.GetFieldAsString("pnbr_action") == "Booked" || r.GetFieldAsString("pnbr_action") == "Amend"))
                                    {

                                        string[] prams = { "pnbr_action" };
                                        string[] updat = { "Cancel" };
                                        for (int i = 0; i < prams.Length; i++)
                                        {

                                            r.SetField(prams[0].ToString(), updat[0].ToString());

                                        }
                                        r.SaveChanges();
                                        Record newr = new Record("Planbuilder");
                                        string[] planprams = new string[] { "pnbr_color", "pnbr_days", "pnbr_plan", "pnbr_publications", "pnbr_sections", "pnbr_size", "pnbr_Standardsize", "pnbr_commissiontype", "pnbr_ratecard", "pnbr_total", "pnbr_commissiontype", "pnbr_Note", "pnbr_date", "pnbr_keynumber", "pnbr_subsection", "pnbr_standardrate" };
                                        string[] planvalues = new string[] { color, day, bid.ToString(), pubid, secid, "standard", size2, Commission, Rateid, rateamount, Commission, caption, insertiondate.ToString(), key, subsecid, rateamount };
                                      
                                        for (int i = 0; i < planprams.Length; i++)
                                        {
                                            //       AddContent(i.ToString());
                                            newr.SetField(planprams[i].ToString(), planvalues[i].ToString());

                                        }
                                        // AddContent("added");
                                        newr.SaveChanges();
                                        InsertCount++;
                                        int plannbuildid = newr.RecordId;
                                        mWSheet1.Cells[j, 18] = plannbuildid.ToString();

                                    }
                                    //r.SaveChanges();


                                   else if (r.GetFieldAsString("pnbr_action") == "Booked" || r.GetFieldAsString("pnbr_action") == "Amend")
                                    {
                                        string[] prams = new string[] { "pnbr_color", "pnbr_days", "pnbr_plan", "pnbr_publications", "pnbr_sections", "pnbr_size", "pnbr_Standardsize", "pnbr_commissiontype", "pnbr_ratecard", "pnbr_cost", "pnbr_commissiontype", "pnbr_note", "pnbr_date", "pnbr_keynumber", "pnbr_subsection", "pnbr_standardrate", "pnbr_action","pnbr_total" };
                                        string[] updat = new string[] { color, day, bid.ToString(), pubid, secid, "standard", size2, Commission, Rateid, rateamount, Commission, caption, insertiondate.ToString(), key, subsecid, rateamount, "Amend" ,rateamount};

                                        for (int i = 0; i < prams.Length; i++)
                                        {

                                            r.SetField(prams[i].ToString(), updat[i].ToString());

                                        }
                                    }
                                    else
                                    {
                                        string[] prams = new string[] { "pnbr_color", "pnbr_days", "pnbr_plan", "pnbr_publications", "pnbr_sections", "pnbr_size", "pnbr_Standardsize", "pnbr_commissiontype", "pnbr_ratecard", "pnbr_cost", "pnbr_commissiontype", "pnbr_note", "pnbr_date", "pnbr_keynumber", "pnbr_subsection", "pnbr_standardrate","pnbr_total" };
                                        string[] updat = new string[] { color, day, bid.ToString(), pubid, secid, "standard", size2, Commission, Rateid, rateamount, Commission, caption, insertiondate.ToString(), key, subsecid, rateamount ,rateamount};

                                        for (int i = 0; i < prams.Length; i++)
                                        {

                                            r.SetField(prams[i].ToString(), updat[i].ToString());
                                        }
                                    }
                                    r.SaveChanges();
                                    mWSheet1.Cells[j, 18] = r.RecordId.ToString();

                                }


                                else if (status.Equals("new") || status.Equals("NEW") || status.Equals("New"))
                                {
                                       AddContent("!" + caption + "!");
                                    //      AddContent("DASTYDCSGBUH + " + day);
                                    string PlanEnity = "Planbuilder";
                                    string[] planprams = new string[] { "pnbr_color", "pnbr_days", "pnbr_plan", "pnbr_publications", "pnbr_sections", "pnbr_size", "pnbr_Standardsize", "pnbr_commissiontype", "pnbr_ratecard", "pnbr_total", "pnbr_commissiontype", "pnbr_Note","pnbr_date","pnbr_keynumber","pnbr_subsection","pnbr_standardrate" };
                                    string[] planvalues = new string[] { color, day, bid.ToString(), pubid, secid, "standard", size2, Commission, Rateid, rateamount, Commission, caption,insertiondate.ToString(),key,subsecid,rateamount };

                                    int plannbuildid = 0;
                                    Record planRecord = new Record(PlanEnity);
                               //     AddContent("PRE");
                                    for (int i = 0; i < planprams.Length; i++)
                                    {
                                 //       AddContent(i.ToString());
                                        planRecord.SetField(planprams[i].ToString(), planvalues[i].ToString());

                                    }
                                    AddContent("added");
                                    planRecord.SaveChanges();
                                    InsertCount++;
                                    plannbuildid = planRecord.RecordId;
                                    mWSheet1.Cells[j, 18] = plannbuildid.ToString();

                                }
                                else
                                {
                                    string searchfield = "pnbr_pnbr_planbuilderid";

                                    if (revised) searchfield = "pnbr_revised";

                                    string pbid = (string)((mWSheet1.Cells[j, 18] as Microsoft.Office.Interop.Excel.Range).Value.ToString());
                                
                                    Record r = FindRecord("PlanBuilder", searchfield + " = '" + pbid + "'");


                                    string[] prams = { "pnbr_action" };
                                    string[] updat = { "No Change" };
                                    for (int i = 0; i < prams.Length; i++)
                                    {

                                        r.SetField(prams[0].ToString(), updat[0].ToString());

                                    }
                                    

                                    mWSheet1.Cells[j, 18] = r.RecordId.ToString();
                                    r.SaveChanges();
                                    
                                }

                           //     AddContent("END");
                            }
                            mWorkBook.SaveAs(GetLibraryPath() + @"Plan\" + agentcy + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
                   Missing.Value, Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                   Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true, Missing.Value,
                   Missing.Value, Missing.Value);


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
                            string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                            sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&dotnetfunc=RunPlanImportStatusPage&inserted=" + InsertCount + "&name=" + bid + "&Failed=0";
                            getSMTPDetails();
                            getSMPTPassword();
                            SendEmail();
                            Dispatch.Redirect(sURL);

                        }
                        catch (IndexOutOfRangeException e)
                        {
                            string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                            sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=ImportCodenew&dotnetfunc=RunPlanImportStat&inserted=" + InsertCount + "&name=" + bid2;
                            //  Dispatch.Redirect(sURL);

                        }


                        catch (Exception Ex)
                        {

                            this.AddError(Ex.Message);

                        }

                        finally
                        {

                        }


                    }
                }
            }
            // }
            catch (Exception Ex)
            {
                this.AddError(Ex.Message);
            }



        }


        public void getSMTPDetails()
        {
            try
            {
                string sSQL = "select Parm_Name,Parm_Value from Custom_SysParams where Parm_Name IN('SMTPServer','SMTPPort','SMTPPassword','SMTPUserName')";
                //AddContent("<br> Executed SQL Query");
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
                // AddContent(ex.Message.ToString());
            }
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
            //msg.To.Add("amitpardhedev@gmail.com");
            msg.To.Add(sToEmailAddress);
            msg.Subject = agentcy + " News Works Plan" + DateTime.Now.ToShortDateString();
            sb.Append("<br>");
            Record ObjEmailTemplate = FindRecord("EmailTemplates", "EmTe_Name='PlanToSelf' and EmTe_Entity='Booking'");
            if (!ObjEmailTemplate.Eof())
                msg.Body = ObjEmailTemplate.GetFieldAsString("EmTe_Comm_Email");

            msg.IsBodyHtml = true;
            System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(servername);
            // AddContent("Attaching");
            //string sFileName = GetLibraryPath() + "VisitServiceReports\\" + objFileRec.GetFieldAsString("libr_FileName");
            string sFileName = GetLibraryPath() + @"Plan\" + agentcy + ".xlsx";

            try
            {
                if (sFileName != null)
                    msg.Attachments.Add(new Attachment(sFileName));
                //msg.Attachments.Add(new Attachment(@"C:\Program Files (x86)\Sage\CRM\CRM\Library\Plan\QuoteSendToSelf\Template\SendToSelf.xls"));

            }
            catch (Exception ex)
            {
                AddError("Unable To Send Attachment");
                //    AddContent(ex.Message.ToString());
                AddUrlButton("Continue", "Continue.gif", UrlDotNet("PlanImportNew", "RunPlanImportNew" + ""));
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
            }
            catch (Exception ex) { }
            // bool wrkflwResult;

            //    wrkflwResult = crmHelperObj.ProgressWorkflow(sBookingID, "Booking", "Booking Workflow", "Self");
            //    if (wrkflwResult)
            //    {
            //        AddInfo("Workflow progressed successfully to Self state");

            //        crmHelperObj.SetStageStatus("Booking", sBookingID, "sendtoself", "InProgress");
            //    }
            //    else
            //    {
            //        AddInfo("Error Occurred during worlflow progress");
            //    }
            //    string sURL = Url("432") + "&book_bookingid=" + sBookingID;
            //    AddUrlButton("Continue", "Continue.gif", UrlDotNet("NZPACRM", "RunPlanDataPage" + "&book_bookingid=" + sBookingID + ""));
            //}
            //catch (Exception ex)
            //{
            //    AddUrlButton("Continue", "Continue.gif", UrlDotNet("NZPACRM", "RunPlanDataPage" + "&book_bookingid=" + sBookingID + ""));
            //    AddError("Error Occured While Sending An Email: " + ex.Message);
            //}
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


        private string comissionconvert(string co)
        {
            return co;
        }
        private void InsertIntoCaption(string sStandardSize, string p)
        {
            string sCode = "";
            Record objCustomCaption = FindRecord("Custom_captions", " LOWER(LTRIM(RTRIM(capt_family)))='" + p.ToLower().Trim() + "' and LOWER(LTRIM(RTRIM(capt_us)))='" + sStandardSize.ToLower().Trim() + "'");
            sCode = sStandardSize.Replace(" ", "");
            if (!objCustomCaption.Eof())
            {
                objCustomCaption.SetField("capt_code", sCode);
                objCustomCaption.SetField("capt_US", sStandardSize);

                objCustomCaption.SaveChanges();
            }
            else
            {
                Record objCustomCaptionInsert = new Record("Custom_captions");

                objCustomCaptionInsert.SetField("capt_US", sStandardSize);
                objCustomCaptionInsert.SetField("capt_code", sCode);
                objCustomCaptionInsert.SetField("capt_family", p);

                objCustomCaptionInsert.SaveChanges();
            }


        }


        #region Save File in  Rate Card folder
        private string SaveRateCardRLocation()
        {


            string FileName = "";
            string newFullPath = "";
            string UploadfilePath = Dispatch.ContentField("HIDDEN_FilePath");
            //string UploadfilePath = "C:\\Users\\Administrator\\Desktop\\Sia.xlsx";
            //AddContent(UploadfilePath);
            string LibPath = GetLibraryPath();
            string NewPath = LibPath.Replace("\\Library", "");
            FileName = Dispatch.ContentField("HIDDEN_FileName");
            NewPath += @"\WWWRoot\CustomPages\NZPAImport\ImportedFiles\Plans\";
            if (Directory.Exists(NewPath))
            {
                //FileUpload1.PostedFile.FileName

                NewPath = NewPath + FileName;

                #region To check if file allready exists
                int count = 1;
                string fileNameOnly = Path.GetFileNameWithoutExtension(NewPath);
                string extension = Path.GetExtension(NewPath);
                string path = Path.GetDirectoryName(NewPath);
                if (!Path.IsPathRooted(NewPath))
                    NewPath = Path.GetFullPath(NewPath);

                string[] files = Directory.GetFiles(path);
                if (File.Exists(NewPath))
                {
                    count = files.Length;
                    newFullPath = Path.Combine(path, String.Format("{0} ({1}){2}", fileNameOnly, (count++), extension));
                    NewPath = newFullPath;
                    File.Copy(UploadfilePath, newFullPath);
                    // AddContent("HORSE");
                }

                else
                {

                    File.Copy(UploadfilePath, NewPath);

                }
            }
            #endregion
            return NewPath;
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

        #region Get Caption Code
        public string GetCaptionCode(string sCaption, string sFamily)
        {
            string sCode = "";
            Record objCustomCaption = FindRecord("Custom_captions", "LOWER(LTRIM(RTRIM(capt_us)))='" + sCaption.ToLower().Trim() + "' and LOWER(LTRIM(RTRIM(capt_family)))='" + sFamily.ToLower().Trim() + "'");

            if (!objCustomCaption.Eof())
            {
                sCode = objCustomCaption.GetFieldAsString("capt_code");
            }

            return sCode;
        }
        #endregion

        #region Get Base Currency
        public string GetBaseCurrency()
        {
            string sBaseCurrency = "";
            string strSQL = "select parm_value from Custom_SysParams  where Parm_Name ='BaseCurrency' and Parm_Deleted is null";
            QuerySelect objCurrencyRec = GetQuery();
            objCurrencyRec.SQLCommand = strSQL;
            objCurrencyRec.ExecuteReader();

            if (!objCurrencyRec.Eof())
            {
                sBaseCurrency = objCurrencyRec.FieldValue("parm_value");
            }
            return sBaseCurrency;
        }
        #endregion
        #endregion

        public void GeneratelogFile(string Logcontent)
        {
            string LibPath = GetLibraryPath();

            string NewPath = LibPath.Replace("\\Library", "");
            NewPath += "\\WWWRoot\\CustomPages\\NZPAImport\\";
            string sInstallDirName = new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).DirectoryName;
            string Logspath = null;
            try
            {
                string currentPath = NewPath;
                if (!Directory.Exists(Path.Combine(currentPath, "LogFiles")))
                    Directory.CreateDirectory(Path.Combine(currentPath, "LogFiles"));

                DateTime theDate = DateTime.Now;
                string ymd = System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString() + theDate.ToString("yyyyMMdd") + "ImportRateCard.txt";
                LogfileName = ymd;
                Logspath = NewPath + "\\LogFiles\\" + ymd;

                if (!File.Exists(Logspath))
                {
                    File.Create(Logspath).Close();
                    using (StreamWriter stream = new StreamWriter(Logspath, true))
                    {
                        stream.WriteLine(Logcontent);
                        stream.Flush();
                        stream.Close();
                    }
                }
                else
                {
                    using (StreamWriter stream = new StreamWriter(Logspath, true))
                    {
                        stream.WriteLine(Logcontent);
                        stream.Flush();
                        stream.Close();
                    }
                }
            }

            catch (Exception ex)
            {
                try
                {
                    using (StreamWriter stream = new StreamWriter(Logspath, true))
                    {
                        stream.WriteLine("error occurred =" + Logcontent + "Description= " + ex.Message);
                        stream.Flush();
                        stream.Close();
                    }
                }
                catch (Exception)
                {

                }
            }
        }
    
    public string GenerateSequecenumber(string entityId)
    {
        string strUniqueID = "";
        try
        {
            Record recCustomSysParams = FindRecord("custom_sysparams", "Parm_Name = 'Bookingbook_BookingID'");
            if (!recCustomSysParams.Eof())
            {
                int intUniqueID = recCustomSysParams.GetFieldAsInt("Parm_Value");

                if (intUniqueID > 0)
                {
                    intUniqueID = intUniqueID + 1;
                    recCustomSysParams.SetField("Parm_Value", intUniqueID);
                    recCustomSysParams.SaveChanges();
                    strUniqueID = intUniqueID.ToString();
                    DateTime today = DateTime.Now;
                    int YY = today.Year;
                    int MM = today.Month;
                    int DD = today.Day;
                    StringBuilder builder = new StringBuilder();
                    string year = YY.ToString();
                    strUniqueID = this.CurrentUser.UserId.ToString() + "-" + Convert.ToInt32(strUniqueID.ToString()).ToString("D5");

                }
            }
        }
        catch (Exception ex)
        {
            this.AddError(ex.Message);
        }
        return strUniqueID;
    }
        public string StandardSizeCode(string ss)
        {
            if (ss.Equals("Double Page Spread") || ss.Equals("DOUBLE PAGE SPREAD")) return "DoublePageSpread";
            if (ss.Equals("Full Page") || ss.Equals("FULL PAGE")) return "FullPage";
            if (ss.Equals("2/3 Page Horizontal") || ss.Equals("2/3 PAGE HORIZONTAL")) return "2/3PageHorizontal";
            if (ss.Equals("2/3 Page Vertical") || ss.Equals("2/3 PAGE VERTICAL")) return "2/3PageVertical";
            if (ss.Equals("1/2 Page Horizontal") || ss.Equals("1/2 PAGE HORIZONTAL")) return "1/2PageHorizontal";
            if (ss.Equals("1/2 Page Vertical") || ss.Equals("1/2 PAGE VERTICAL")) return "1/2PageVertical";
            if (ss.Equals("Junior Page") || ss.Equals("JUNIOR PAGE")) return "JuniorPage";
            if (ss.Equals("1/3 Page Horizontal") || ss.Equals("1/3 PAGE HORIZONTAL")) return "1/3PageHorizontal";
            if (ss.Equals("1/3 Page Vertical") || ss.Equals("1/3 PAGE VERTICAL")) return "1/3PageVertical";
            if (ss.Equals("1/3 Page Island") || ss.Equals("1/3 PAGE ISLAND")) return "1/3PageIsland";
            if (ss.Equals("1/4 Page Horizontal") || ss.Equals("1/4 PAGE HORIZONTAL")) return "1/4PageHorizontal";
            if (ss.Equals("1/4 Page Vertical") || ss.Equals("1/4 PAGE VERTICAL")) return "1/4PageVertical";
            if (ss.Equals("1/4 Page Island") || ss.Equals("1/4 PAGE ISLAND")) return "1/4PageIsland";
            if (ss.Equals("1/6 Page Horizontal") || ss.Equals("1/6 PAGE HORIZONTAL")) return "1/6PageHorizontal";
            if (ss.Equals("1/6 Page Vertical") || ss.Equals("1/6 PAGE VERTICAL")) return "1/6PageVertical";
            if (ss.Equals("1/6 Page Island") || ss.Equals("1/6 PAGE ISLAND")) return "1/6PageIsland";
            if (ss.Equals("1/8 Page Horizontal") || ss.Equals("1/8 PAGE HORIZONTAL")) return "1/8PageHorizontal";
            if (ss.Equals("1/8 Page Vertical") || ss.Equals("1/8 PAGE VERTICAL")) return "1/8PageVertical";
            if (ss.Equals("1/8 Page Island") || ss.Equals("1/8 PAGE ISLAND")) return "1/8PageIsland";
            if (ss.Equals("1/8 Page Standard") || ss.Equals("1/8 PAGE STANDARD")) return "1/8PageStandard";
            if (ss.Equals("1/16 Page Vertical") || ss.Equals("1/16 PAGE VERTICAL")) return "1/16PageVertical";
            if (ss.Equals("1/16 Page Island") || ss.Equals("1/16 PAGE ISLAND")) return "1/16PageIsland";
            if (ss.Equals("Square Large") || ss.Equals("SQUARE LARGE")) return "SquareLarge";
            if (ss.Equals("Square Medium") || ss.Equals("SQUARE MEDIUM")) return "SquareMedium";
            if (ss.Equals("Square Small") || ss.Equals("SQUARE SMALL")) return "SquareSmall";
            if (ss.Equals("Front Page Solus") || ss.Equals("FRONT PAGE SOLUS")) return "FrontPageSolus";
            return ss;
        }

        public void copyplanbuilders(string oldplanid, string newplanid)
        {
            Record RecEachPlanBuilder = FindRecord("PlanBuilder", "pnbr_plan = '" + oldplanid + "'");

            while (!RecEachPlanBuilder.Eof())
            {
                //AddContent("<BR>PlanID 1=" + PlanID);
                Record objPlanBuilder = new Record("PlanBuilder");
                objPlanBuilder.SetField("pnbr_plan", newplanid);
                objPlanBuilder.SetField("pnbr_publications", RecEachPlanBuilder.GetFieldAsString("pnbr_publications"));
                objPlanBuilder.SetField("pnbr_ratecard", RecEachPlanBuilder.GetFieldAsString("pnbr_ratecard"));
                objPlanBuilder.SetField("pnbr_sections", RecEachPlanBuilder.GetFieldAsString("pnbr_sections"));
                objPlanBuilder.SetField("pnbr_other", RecEachPlanBuilder.GetFieldAsString("pnbr_other"));
                objPlanBuilder.SetField("pnbr_subsection", RecEachPlanBuilder.GetFieldAsString("pnbr_subsection"));
                objPlanBuilder.SetField("pnbr_days", RecEachPlanBuilder.GetFieldAsString("pnbr_days"));
                objPlanBuilder.SetField("pnbr_date", RecEachPlanBuilder.GetFieldAsString("pnbr_date"));
                objPlanBuilder.SetField("pnbr_size", RecEachPlanBuilder.GetFieldAsString("pnbr_size"));
                objPlanBuilder.SetField("pnbr_height", RecEachPlanBuilder.GetFieldAsString("pnbr_height"));
                objPlanBuilder.SetField("pnbr_width", RecEachPlanBuilder.GetFieldAsString("pnbr_width"));
                objPlanBuilder.SetField("pnbr_custom", RecEachPlanBuilder.GetFieldAsString("pnbr_custom"));
                objPlanBuilder.SetField("pnbr_color", RecEachPlanBuilder.GetFieldAsString("pnbr_color"));
                objPlanBuilder.SetField("pnbr_loading", RecEachPlanBuilder.GetFieldAsString("pnbr_loading"));
                objPlanBuilder.SetField("pnbr_standardrate", RecEachPlanBuilder.GetFieldAsString("pnbr_standardrate"));
                objPlanBuilder.SetField("pnbr_loadingvalue", RecEachPlanBuilder.GetFieldAsString("pnbr_loadingvalue"));
                objPlanBuilder.SetField("pnbr_discount", RecEachPlanBuilder.GetFieldAsString("pnbr_discount"));
                objPlanBuilder.SetField("pnbr_total", RecEachPlanBuilder.GetFieldAsString("pnbr_total"));
                objPlanBuilder.SetField("pnbr_commissiontype", RecEachPlanBuilder.GetFieldAsString("pnbr_commissiontype"));
                objPlanBuilder.SetField("pnbr_standardsize", RecEachPlanBuilder.GetFieldAsString("pnbr_standardsize"));
                objPlanBuilder.SetField("pnbr_keynumber", RecEachPlanBuilder.GetFieldAsString("pnbr_keynumber"));
                objPlanBuilder.SetField("pnbr_revised", RecEachPlanBuilder.GetFieldAsInt("pnbr_pnbr_planbuilderid"));
                objPlanBuilder.SetField("pnbr_action", RecEachPlanBuilder.GetFieldAsString("pnbr_action"));
                //objPlanBuilder.SetField("pnbr_plan", intNextBookingID);

                objPlanBuilder.SaveChanges();
                RecEachPlanBuilder.GoToNext();
            }
        }


    }
}