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
    public class PlanSendWhenBooked : Web
    {

        Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        Microsoft.Office.Interop.Excel.Application oXL;

        double totalamount = 0;
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
        string type = "";
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
        List<string> emails = new List<string>();
        CRMHelper crmHelperObj = new CRMHelper();
        public PlanSendWhenBooked()
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
            bookContact = recBook.GetFieldAsString("book_contact");
        }
        public override void BuildContents()
        {
          //  ReadExistingExcel();
            string getpubs = "select * from Planbuilder where pnbr_plan = " + sBookingID + "and pnbr_deleted is null";
            List<int> pubids = new List<int>();
            QuerySelect objAddressRec = GetQuery();

            objAddressRec.SQLCommand = getpubs;
            objAddressRec.ExecuteReader();
          //  AddContent("papercut");
            while (!objAddressRec.Eof())
            {
                int id = Int32.Parse(objAddressRec.FieldValue("pnbr_publications"));
                emailtoUpdate(id);
                ReadExistingExcel(id);
               // AddContent("GOING");
                if (!pubids.Contains(id))
                {
                    string getrecs = "select * from Planbuilder where pnbr_plan = " + sBookingID + " and pnbr_publications = " + objAddressRec.FieldValue("pnbr_publications");
                //    AddContent(getrecs);
                    QuerySelect objrecs = GetQuery();

                    objrecs.SQLCommand = getrecs;
                    objrecs.ExecuteReader();
                    writeheader();
                 //   AddContent("HEAD DONE");
                    SetAgencyXml(objAddressRec.FieldValue("pnbr_publications"));
                    while (!objrecs.Eof())
                    {
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
                        //   AddContent(pubs);
                      //  AddContent("ENTERING DETS");

                        writexmldetails(sections, subsections, color, ss, high, days, date, key, rc, pubs);
                     //   AddContent("FINISHED DETS");
                        objrecs.Next();
                    }
                   // AddContent("HELP");
                    string end = "\t\t<date>" + DateTime.Now.ToString() + "</date>\n";
                    end += "\t</export>\n";
                    end += "</booking>";
                    xmlpath = GetLibraryPath() + @"\Plan\" + max + ".xml";
                    xmlwrite = header + body + end;
                    System.IO.File.WriteAllText(xmlpath, xmlwrite);
                    max++;

                    pubids.Add(id);
                }
                header = "";
                body = "";
                
                objAddressRec.Next();
            }

            //writeheader();
            //SetAgencyXml();

            getSMTPDetails();
            getSMPTPassword();
            SendEmail();
        }
        public void emailtoUpdate(int pid)
        {
            string Psql = "select * from Publications where Pblc_PublicationsID = " + pid;
            QuerySelect objAddressRecsection = GetQuery();
            objAddressRecsection.SQLCommand = Psql;
            objAddressRecsection.ExecuteReader();
            if (!objAddressRecsection.Eof())
            {
                emails.Add(objAddressRecsection.FieldValue("pblc_Bookingemailaddress"));
            }

        }
        public void writeheader()
        {
            xmlwrite += "<booking>";
            xmlwrite += "\t<export>";
        }


        public void ReadExistingExcel(int pubid)
        {
            //string path = @"E:\Projects\ETL914-C\Docs\SendToSelf.xls";            
            orgFilePath = GetLibraryPath() + @"Plan\Booking.xls";

            string checkPathGenTemplate = GetLibraryPath() + @"\Plan";

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
                    //AddContent("on form");
                    mWorkBook = oXL.Workbooks.Open(orgFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                   // AddContent("on form");
                    mWorkSheets = mWorkBook.Worksheets;
                   // AddContent("on form");
                    //Get all the sheets in the workbook
                    //Get the already exists sheet
                    
                    mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item(1);
                   // AddContent("on form");
                    Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
               //     AddContent("on form");
                    //'Set Agency Name
                    SetAgency();
               //     AddContent("CAT"); //set detail grid value
                    SetDetailGrid(pubid);
                }
                catch (Exception ex)
                {
                    AddContent(ex.Message.ToString());
                }
          //      AddContent("BEAR");
                string date = "";
             //   AddContent("Dragon");
                date = DateTime.Now.Date.ToShortDateString().Replace(@"/", "-");
                mergedFilePath = GetLibraryPath() + @"Plan\Booking_" + max + ".xls";
              //  AddContent("SAFE");
                string sSavedCPSpath = mergedFilePath;
            //    AddContent(mergedFilePath);
                try
                {
                   // AddContent("SAVING");
                    mWorkBook.SaveAs(sSavedCPSpath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                    Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value);
                   // AddContent("CLOSING");
                    mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);

                    Process[] processes = Process.GetProcessesByName("EXCEL");
               //    AddContent("KILLING");
                    foreach (var process in processes)
                    {
                        if (process.MainWindowTitle == "")
                            process.Kill();
                    }
                    mWSheet1 = null;
                    mWorkBook = null;
               //     AddContent("CLEANING");
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    //' mWorkBook.Close();
                   // AddContent("PUTTING AWAY");
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
                    //   AddContent(ex.Message.ToString());
                    //MessageBox.Show("Failed to generte CPS work order sheet. File is open or may be path is not defined to save the file.");

                }
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(mWorkBook);
            }
        }
        private void SetAgency( )
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
           // sNewsWorksRef = objAddressRec.FieldValue("book_Reference").ToString();
            //sQuoteVersionRef = objAddressRec.FieldValue("book_costingversion").ToString();
            if (sBilledBy == "Works") sBilledBy = "NewsWorks";
            sStatus = objAddressRec2.FieldValue("book_Status");
            sNewsWorksRef = objAddressRec2.FieldValue("book_Reference");
            sQuoteVersionRef = objAddressRec2.FieldValue("book_costingversion");

            mWSheet1.Cells[20, 12] = sAgency;
            mWSheet1.Cells[21, 12] = sAgencyContact;
            mWSheet1.Cells[22, 12] = sAdvertiser;
            mWSheet1.Cells[23, 12] = sCampaign;
            mWSheet1.Cells[24, 12] = sBilledBy;
            mWSheet1.Cells[20, 69] = sCreationDate;
            mWSheet1.Cells[21, 69] = sCreatedBy;
            mWSheet1.Cells[22, 69] = sQuoteVersionRef;
            mWSheet1.Cells[23, 69] = sNewsWorksRef;
            mWSheet1.Cells[24, 69] = sStatus;
        }
        private void SetDetailGrid(int id)
        {
         //   AddContent("starting details");
            int j = 29;
            int end = 53;
            float commissionTotal = 0;
            int insertation = 0;
            totalamount = 0;
            string sCode = "";
            string strSQL = "select pblc_Name,pblc_CreatedBy,pblc_CreatedDate,pblc_commision, pblc_Status book_reference,pnbr_pnbr_planbuilderid,pnbr_action, pnbr_days, pnbr_discount, pnbr_color, pnbr_ratecard, pnbr_cost, pnbr_size, pnbr_sections,*";
                 
         
            strSQL += " from  Booking left  join Planbuilder on book_BookingID=pnbr_plan left join Publications on pnbr_publications = pblc_PublicationsID left JOIN Sections on sctn_Sctn_sectionid=pnbr_sections ";
            
            strSQL += " where pblc_PublicationsId=" + id  + " and book_BookingID=" + sBookingID  + "and pnbr_deleted is null" ;


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

            if (count > 23)
            {
                int insert = count - 23;
                for (int k = 0; k < insert; k++)
                {
                    Microsoft.Office.Interop.Excel.Range row = range.EntireRow;
                    row.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, RngToCopy.Copy(Type.Missing));
                    end++;
                }
            }

            objPblcnRec.SQLCommand = strSQL;
            objPblcnRec.ExecuteReader();



            while (!objPblcnRec.Eof())

            {
                if (string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_Action")))
                {
                    Record r = FindRecord("planbuilder", "pnbr_pnbr_planbuilderid = '" + objPblcnRec.FieldValue("pnbr_pnbr_planbuilderid") + "'");
                    r.SetField("pnbr_action", "Booked");
                    r.SaveChanges();
                    mWSheet1.Cells[j, 4] = "Book";
                }

               else if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_Action"))){
                    string action = objPblcnRec.FieldValue("pnbr_Action").ToString();
                    
                    mWSheet1.Cells[j, 4] = objPblcnRec.FieldValue("pnbr_Action").ToString();
                    if (action == "Amend" || action == "No Change" || action == "NoChange") 
                    {
                        Record r = FindRecord("planbuilder", "pnbr_pnbr_planbuilderid = '" + objPblcnRec.FieldValue("pnbr_pnbr_planbuilderid") + "'");
                        r.SetField("pnbr_action", "Booked");
                        r.SaveChanges();
                    }
                    else if(action == "Cancel" || action == "cancel")
                    {
                        Record r = FindRecord("planbuilder", "pnbr_pnbr_planbuilderid = '" + objPblcnRec.FieldValue("pnbr_pnbr_planbuilderid") + "'");
                        r.SetField("pnbr_deleted", "1");
                        r.SaveChanges();
                    }
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                    mWSheet1.Cells[j, 11] = objPblcnRec.FieldValue("pblc_Name").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_commissiontype").ToString()))
                {
                    float commiss = float.Parse(objPblcnRec.FieldValue("pblc_commision"), System.Globalization.CultureInfo.InvariantCulture);
                    type = objPblcnRec.FieldValue("pnbr_commissiontype").ToString();
                    if (type != "Non Commission" && type!= "NonCommission" && type!= "Noncommission" && type!= "NonCommission")
                    {
                        //float rate = commiss / 100;
                        mWSheet1.Cells[j, 64] = commiss.ToString() + "%";
                    }
                    else { mWSheet1.Cells[j, 64] = "0%"; }
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("sctn_name").ToString()))
                {
                    mWSheet1.Cells[j, 25] = objPblcnRec.FieldValue("sctn_name").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_subsection").ToString()))
                {
                    string subsql = "Select * from subsection where suse_subsectionid = '" + objPblcnRec.FieldValue("pnbr_subsection").ToString() + "'";

                    QuerySelect objRecsub = GetQuery();
                    objRecsub.SQLCommand = subsql;
                    objRecsub.ExecuteReader();
                    mWSheet1.Cells[j, 26] = objRecsub.FieldValue("suse_name").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_days").ToString()))
                {
                    mWSheet1.Cells[j, 14] = GetDays(objPblcnRec.FieldValue("pnbr_days").ToString());
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_date").ToString()))
                {
                    DateTimeFormatInfo format = new DateTimeFormatInfo();
                    format.ShortDatePattern = "dd-MM-yy";
                    format.DateSeparator = "-";
                    DateTime date = Convert.ToDateTime(objPblcnRec.FieldValue("pnbr_date"));
                    mWSheet1.Cells[j, 19] = Convert.ToDateTime(date, format);
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                {
                   // mWSheet1.Cells[j, 5] = objPblcnRec.FieldValue("pblc_Name").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_size").ToString()))
                {
                    mWSheet1.Cells[j, 27] = objPblcnRec.FieldValue("pnbr_standardsize").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_keynumber").ToString()))
                {
                    mWSheet1.Cells[j, 46] = objPblcnRec.FieldValue("pnbr_keynumber").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_inserts").ToString()))
                {
                    mWSheet1.Cells[j, 39] = objPblcnRec.FieldValue("pnbr_inserts").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_caption").ToString()))
                {
                    mWSheet1.Cells[j, 40] = objPblcnRec.FieldValue("pnbr_caption").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_plan").ToString()))
                {
                    string ratesql = "Select * from RatesCard where rate_RatesCardid = '" + objPblcnRec.FieldValue("pnbr_ratecard").ToString() + "'";

                    QuerySelect objRec = GetQuery();
                    objRec.SQLCommand = ratesql;
                    objRec.ExecuteReader();
                    if (!objRec.Eof())
                    {
                      //  AddContent(objPblcnRec.FieldValue("pnbr_date").ToString());
                      AddContent("FREAK");
                       // DateTime insert = objPblcnRec.FieldValueAsDate("pnbr_date");
                       AddContent("PASS");
                      //  string before = objRec.FieldValue("rate_BookinDeadlineDays").ToString().Substring(0,1);
                       // int bre = Int32.Parse(before);
                       
                      //  insert.AddDays((bre*-1));
                        //mWSheet1.Cells[j, 81] = insert.ToShortDateString() + " " + objRec.FieldValue("rate_BookingDeadlineTime").ToString();
                        int widthlim = -1;
                        if (!string.IsNullOrEmpty(objRec.FieldValue("rate_size").ToString()) && (!string.IsNullOrEmpty(objRec.FieldValue("rate_standardsize").ToString())))
                        {
                            if (objRec.FieldValue("rate_Size").ToString() != "Standard" || objRec.FieldValue("rate_standardsize").Contains("Module") || objRec.FieldValue("rate_standardsize").Contains("cm") || objRec.FieldValue("rate_standardsize").Contains("Cm") || objRec.FieldValue("rate_standardsize").Contains("CM"))
                            {
                                for (int i = 1; i < 13; i++)
                                {
                                    if (!string.IsNullOrEmpty((objRec.FieldValue("rate_width" + i.ToString()).ToString())))
                                    {
                                        string num = objRec.FieldValue("rate_width" + i.ToString());
                                        //AddContent("STRING" + i + " " + num);
                                        double lim = Double.Parse(num);
                                        if (lim == 0)
                                        {
                                            widthlim = i - 1;
                                            break;
                                        }
                                    }
                                }
                                if (widthlim == -1)
                                {
                                    widthlim = 12;
                                }
                            }
                            else
                            {
                               // AddContent("OTHER");
                                if (!string.IsNullOrEmpty(objRec.FieldValue("rate_SetSizesWidth").ToString()))
                                {
                                //    AddContent("GO");
                                    string num = objRec.FieldValue("rate_SetSizesWidth");
                                    int lim = Int32.Parse(num);
                                 //   AddContent("DIE");
                                    widthlim = lim;
                                }
                            }
                        }
                         //  AddContent("HAPPY");
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
                        mWSheet1.Cells[j, 33] = specs;
                        
                        string specsize = objRec.FieldValue("rate_height") + "x" + objRec.FieldValue("rate_width");
                      //  mWSheet1.Cells[j, 51] = specsize;
                      //  AddContent("CATSds");
                        //insert = objPblcnRec.FieldValueAsDate("pnbr_date");
                       // string before2 = objRec.FieldValue("rate_MaterialDeadlinDays").ToString().Substring(0, 1);
                        //int bre2= Int32.Parse(before2);

                       // insert.AddDays((bre2 * -1));
                       // mWSheet1.Cells[j, 87] = insert.ToShortDateString() + " " + objRec.FieldValue("rate_MaterialDeadlineTime").ToString(); 


                    }
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_width").ToString()))
                {
                    //mWSheet1.Cells[j, 40] = objPblcnRec.FieldValue("pnbr_width").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_color").ToString()))
                {
                    string color = objPblcnRec.FieldValue("pnbr_color").ToString();
                    if (color == "Color" || color == "COLOR") color = "Colour";
                    if (color == "NoColor" || color == "NOCOLOR") color = "Mono";
                    mWSheet1.Cells[j, 52] = color;
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_standardrate").ToString()))
                {
                    string temp = objPblcnRec.FieldValue("pnbr_standardrate").ToString();
                    double cur = double.Parse(temp);
                    //mWSheet1.Cells[j, 63] = cur.ToString("C2");
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_total").ToString()))
                {
                    string temp = objPblcnRec.FieldValue("pnbr_total").ToString();
                    double cur = double.Parse(temp);
                    totalamount += cur;
                    mWSheet1.Cells[j, 70] = cur.ToString("C2");
                }

                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_package").ToString()))
                {

                    string subsql = "Select * from packages where pack_packagesid = '" + objPblcnRec.FieldValue("pnbr_package").ToString() + "'";

                    QuerySelect objRecsub = GetQuery();
                    objRecsub.SQLCommand = subsql;
                    objRecsub.ExecuteReader();
                    //mWSheet1.Cells[j, 26] = objRecsub.FieldValue("suse_name").ToString();
                    mWSheet1.Cells[j, 58] = objRecsub.FieldValue("pack_name").ToString();
                }
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_Note").ToString()))
                {
                    mWSheet1.Cells[j,76 ] = objPblcnRec.FieldValue("pnbr_Note").ToString();
                }
                //if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_color").ToString()))
                //{
                //    mWSheet1.Cells[j, 9] = objPblcnRec.FieldValue("pnbr_color").ToString();
                //}
                //if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_ratecard").ToString()))
                //{
                //    mWSheet1.Cells[j, 10] = objPblcnRec.FieldValue("pnbr_ratecard").ToString();
                //}
                //if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_discount").ToString()))
                //{
                //    mWSheet1.Cells[j, 11] = objPblcnRec.FieldValue("pnbr_discount").ToString();
                //}
                //if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_total").ToString()))
                //{
                //    mWSheet1.Cells[j, 12] = objPblcnRec.FieldValue("pnbr_total").ToString();
                //}
                //if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                //{
                //    mWSheet1.Cells[j, 13] = "";
                //}
                //if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pblc_Name").ToString()))
                //{
                //    mWSheet1.Cells[j, 14] = "";
                //}
                float totalint = 0;
                if (!string.IsNullOrEmpty(objPblcnRec.FieldValue("pnbr_total").ToString()))
                 totalint = float.Parse(objPblcnRec.FieldValue("pnbr_total"), System.Globalization.CultureInfo.InvariantCulture);
                float commission = float.Parse(objPblcnRec.FieldValue("pblc_commision"), System.Globalization.CultureInfo.InvariantCulture);
                float totalcom = totalint * (commission/100);
                if (type != "NonCommission")
                    commissionTotal += totalcom;
                
                j++;
                objPblcnRec.Next();
                insertation++;
            }
            mWSheet1.Cells[end - 1, 71] = totalamount.ToString("C2");
            mWSheet1.Cells[end, 64] = "Total Insertions: " + insertation;

            mWSheet1.Cells[end, 74] = "Commission: " + commissionTotal.ToString("C2");
            // AddContent("ending details");
        }

       


        public void writexmldetails(string sections, string subsection, string color, string standardsize, string height, string days, string date, string key, string rate, string pubname)
        {



            body += "\t\t<ad>\n";
            body += "\t\t\t<publication>\n";
            body += "\t\t\t\t<pub_id>" + "" + "</pub_id>\n";
            string ratesql = "select * from Publications where pblc_publicationsid = " + pubname;
            QuerySelect objgetrate = GetQuery();
            objgetrate.SQLCommand = ratesql;
            objgetrate.ExecuteReader();
            if (!string.IsNullOrEmpty(objgetrate.FieldValue("pblc_Name").ToString()))
                body += "\t\t\t\t<pub_name>" + objgetrate.FieldValue("pblc_Name") + "</pub_name>\n";
            else body += "\t\t\t\t<pub_name>" + "" + "</pub_name>\n";
            body += "\t\t\t</publication>\n";
           // AddContent("1");

            body += "\t\t\t<placement_details>\n";
           
            body += "\t\t\t\t<class_id>" + "" + "</class_id>\n"; // PLEASE CHANGE ME!!!!!
            string strSQLsection = "select * from Sections where sctn_sctn_sectionid = " + sections;
            //string sCode = "";

            QuerySelect objAddressRecsection = GetQuery();
            objAddressRecsection.SQLCommand = strSQLsection;
            objAddressRecsection.ExecuteReader();
            if (!string.IsNullOrEmpty(objAddressRecsection.FieldValue("sctn_name").ToString()))
                body += "\t\t\t\t<class_name>" + objAddressRecsection.FieldValue("sctn_name") + "</class_name>\n";
            else body += "\t\t\t\t<class_name>" + "" + "</class_name>\n";
            body += "\t\t\t\t<slot_id>" + "" + "</slot_id>\n";
            //AddContent("A");
            //strSQLsection = "select * from Subsection where suse_subsectionid = '" + subsection + "'";

           
            //string sCode = "";

            //objAddressRecsection = GetQuery();
            //objAddressRecsection.SQLCommand = strSQLsection;
            //objAddressRecsection.ExecuteReader();
            //if (!objAddressRecsection.Eof())
            
                //if (!string.IsNullOrEmpty(objAddressRecsection.FieldValue("suse_name").ToString()))
                    body += "\t\t\t\t<slotname> </slotname>\n";
               // else body += "\t\t\t\t<sub_section_name>" + "" + "</sub_section_name>\n";
            //}
            //else body += "\t\t\t\t<sub_section_name>" + "" + "</sub_section_name>\n";
            body += "\t\t\t\t<colour>" + color + "</colour>\n";
            body += "\t\t\t\t<key_number>" + key + "</key_number>\n";
            body += "\t\t\t\t<caption>" + "" + "</caption>\n";
            body += "\t\t\t\t<placement_comment>" + "" + "</placement_comment>\n";
            body += "\t\t\t</placement_details>\n";
            body += "\t\t\t<adsize>\n";
           
            if (!string.IsNullOrEmpty(standardsize.ToString()))
            {
               // AddContent("10");
                body += "\t\t\t\t<ad_izename>" + standardsize + "</adsizename>\n";
                body += "\t\t\t\t<depth>" + "" + "</depth>\n";
                body += "\t\t\t\t<depth_unit>" + "" + "</depth_unit>\n";
                body += "\t\t\t\t<columns>" + "" + "</columns>\n";
                //AddContent("100");
            }
            else
            {
               //AddContent("20");
                body += "\t\t\t\t<ad_size_name>" + "" + "</ad_size_name>\n";
                body += "\t\t\t\t<depth>" + height + "</depth>\n";
                body += "\t\t\t\t<depthunit>" + "" + "</depthunit>\n";
                body += "\t\t\t\t<columns>" + "" + "</columns>\n";
                //AddContent("200");
            }
           // AddContent(days);
            body += "\t\t\t</adsize>\n";
            body += "\t\t\t<schedule>\n";
            body += "\t\t\t\t<run_dates>\n";
         //   AddContent("100HI");
            string rateday = "";
            if (days.Length > 0)
            {
                string realday = days.Substring(0, 3);
               // AddContent("DAY");
                rateday = dayswap(realday);
            }
          
            // AddContent("NIGHT");
           // AddContent("2");
            body += "\t\t\t\t\t<run_date>\n";
            string[] dates;
            dates = date.Split(' ');
            body += "\t\t\t\t\t\t<date>" + dates[0] + "</date>\n";
            string rateme = "select * from RatesCard where rate_RatesCardID = '" + rate + "'";
            QuerySelect objrate = GetQuery();
            objrate.SQLCommand = rateme;
            objrate.ExecuteReader();
            if (days.Length > 0)
            {
                if (!string.IsNullOrEmpty(objrate.FieldValue("rate_" + rateday).ToString()))
                    body += "\t\t\t\t\t\t<price>" + objrate.FieldValue("rate_" + rateday) + "</price>\n";
                else body += "\t\t\t\t\t\t<price>" + "" + "</price>\n";
            }
            else { body += "\t\t\t\t\t\t<price>" + "" + "</price>\n"; }
           
            body += "\t\t\t\t\t</run_date>\n";
            body += "\t\t\t\t</run_dates>\n";
            body += "\t\t\t</schedule>\n";
            body += "\t\t</ad>\n";
          //  AddContent("3");
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
            sDay = sDay.Replace(",","");
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
               // AddContent(ex.Message.ToString());
            }
        }
        public void SendEmail()
        {
            for (int i = 0; i < emails.Count && i < max;i++ )
            {

                userID = this.CurrentUser.UserId.ToString();
                string sToEmailAddress = "";
                // for (int i = 0; i < emails.Count; i++)
                //{
                sToEmailAddress += emails[i] ;
                //sToEmailAddress +=  "awells@enabling.net";
                //   if (i < emails.Count - 1) sToEmailAddress += ",";
                // }
                // string sToEmailAddress = ToEmailID();
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
                        msg.Attachments.Add(new Attachment(GetLibraryPath() + @"Plan\Booking_" + i + ".xls"));
                        //  for (int i = 0; i < max; i++)
                        // {
                        msg.Attachments.Add(new Attachment(GetLibraryPath() + @"\Plan\" + i.ToString() + ".xml"));
                    // }
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
                    //AddInfo("Email Sent Successfully");
                    bool wrkflwResult;

                    wrkflwResult = crmHelperObj.ProgressWorkflow(sBookingID, "Booking", "Booking Workflow", "Booked");
                    if (wrkflwResult)
                    {
                       

                        crmHelperObj.SetStageStatus("Booking", sBookingID, "sendtoagency", "InProgress");
                    }
                    else
                    {
                        AddInfo("Error Occurred during worlflow progress");
                    }
                   // AddUrlButton("Continue", "Continue.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage" + "&book_bookingid=" + sBookingID + ""));
                }
                catch (Exception ex)
                {
                   // AddUrlButton("Continue", "Continue.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage" + "&book_bookingid=" + sBookingID + ""));

                    AddError("Error Occured While Sending An Email: " + ex.Message);
                }
            }
            AddInfo("Email Sent Successfully");
            AddInfo("Workflow progressed successfully to Booked state");
            AddUrlButton("Continue", "Continue.gif", UrlDotNet("NZPACRM", "RunPlanDataPage" + "&book_bookingid=" + sBookingID + ""));
        }
        public string ToEmailID()
        {
            string sToEmail = "";
           // string pubSql = "select "
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
            header += "\t\t<id>" + recBook.GetFieldAsInt("book_bookingid").ToString() + "</id>\n";
            if (!string.IsNullOrEmpty(recBook.GetFieldAsString("book_name").ToString()))
                header += "\t\t<action>" + recBook.GetFieldAsString("book_name") + "</action>\n";
            else header += "\t\t<action>" + "" + "</action>\n";
            header += "\t\t<customer>\n";
            if (!string.IsNullOrEmpty(recBook.GetFieldAsString("book_Client").ToString()))
                header += "\t\t\t<client_id>" + recBook.GetFieldAsString("book_Client") + "</client_id>\n";
            else header += "\t\t\t<client_id>" + "" + "</client_id>\n";
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
            else header += "\t\t\t<agency_id>" + "" + "</agency_id>\n";


            string strSQLComp = "select * from Company where comp_companyid = " + recBook.GetFieldAsString("book_agency");
            //string sCode = "";

            QuerySelect objAddresscomp = GetQuery();
            objAddresscomp.SQLCommand = strSQLComp;
            objAddresscomp.ExecuteReader();
            ////'Declare variables
            ////AddContent(objBookingList.GetFieldAsString("comp_name"));
            if (!string.IsNullOrEmpty(objAddresscomp.FieldValue("Comp_Name").ToString()))
                header += "\t\t\t<agency_name>" + objAddresscomp.FieldValue("Comp_Name") + "</agency_name>\n";
            else header += "\t\t\t<agency_name>" + "" + "</agency_name>\n";
            string sqlpubcom = "select * from Publications where pblc_PublicationsID = " + pubxml;
            QuerySelect obj = GetQuery();
            obj.SQLCommand = sqlpubcom;
            obj.ExecuteReader();
            if (!string.IsNullOrEmpty(obj.FieldValue("pblc_Commision").ToString()))
                header += "\t\t\t<commission>" + obj.FieldValue("pblc_Commision") + "%</commission>\n";
            else header += "\t\t\t<commission>" + "" + "</commission>\n";
            header += "\t\t</customer>\n";

        }
    }
}
