using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using System.IO;
using Sage.CRM.Data;
namespace NZPACRM.Plan
{
  public  class importNabsfile: DataPageNew 

    {
        string shttpURL = "";
        Stream filestream;
        StreamReader sr;
        int countall = 1;
        string bookfin;

        public importNabsfile(): base ("booking","book_bookingid"){

            string CurrUser = CurrentUser.UserId.ToString(); // the current user
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


        }

        public override void BuildContents()
        {
            AddContent("<script type='text/javascript' src='../js/custom/ClientFuncs.js'></script>");
            string sSuccessErrorMessage = "";
            int iFailedCount = 0;
            int InsertCount = 0;
            int iDupRecord = 0;
            string isDatavalid = "";

            AddContent(HTML.Form());

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
            AddContent(HTML.Box("Import Nabs Data Beta Version", sHTML));

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

                    
                    using (sr = new StreamReader(SavedFilePath))
                    {
                        
                        countall++;
                        
                        string header = sr.ReadLine();

                        while (!sr.EndOfStream)
                        {
                            AddContent("start");
                            string booking = sr.ReadLine();
                            if (booking.Substring(0, 2) == "EF" || booking.Substring(0, 1) == "H") { countall++; }
                            else
                            {
                               // string book = sr.ReadLine();
                               // AddContent(book.Length.ToString());
                                HandleBooking(booking);
                                countall++;
                            }
                        }
                        sr.Close();
                    }
                    string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                    sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&dotnetfunc=RunPlanImportStatusPage&inserted=" + InsertCount + "&name=" +bookfin  + "&Failed=0";
                   Dispatch.Redirect(sURL);

                }
            }
            //base.BuildContents();
                }




        private void HandleBooking(string bookingline)
        {
            AddContent(bookingline);
            int count = 0;
            int i;
            string head = bookingline.Substring(0, 2);
            if (head == "IH") return;
            string status = bookingline.Substring(2, 10);
            if (status.Contains("NEW") || status.Contains("New"))
            {
                AddContent(bookingline.Length.ToString());
                string writtingentity = "Booking";
                string countstr = bookingline.Substring(185, 5);
                AddContent("CAT" + countstr);
                count = Int32.Parse(countstr);

                string pegAge = bookingline.Substring(12, 6);
                pegAge = pegAge.TrimEnd();
                string compSQL = "Select * from Company where comp_pegid = '" + pegAge + "' and comp_deleted is null";
                QuerySelect compQS = GetQuery();
                compQS.SQLCommand = compSQL;
                compQS.ExecuteReader();
                if (compQS.Eof())
                {
                    AddError("This Agency is either not in the database or has not been assigned a pegID. Line number " + countall.ToString());
                    throw new Exception();
                }
                string companyid = compQS.FieldValue("comp_companyid");

                string clientnab = bookingline.Substring(42, 50).TrimEnd();
                string sqlclient = "select * from Client where client_name like '" + clientnab + "%'";
                QuerySelect clientQS = GetQuery();
                clientQS.SQLCommand = sqlclient;
                clientQS.ExecuteReader();
                int clientid = 0;
                if (clientQS.Eof())
                {
                    // AddError("This Client is either not in the database or the name in the nab does not match the name in crm. Line number " + countall.ToString());
                    //throw new Exception();
                    Record newclient = new Record("Client");
                    string[] prams = new string[] { "client_name", "client_companyid" };
                    string[] valuess = new string[] { clientnab, companyid };

                    for (int k = 0; k < prams.Length; k++)
                    {

                        newclient.SetField(prams[k].ToString(), valuess[k].ToString());

                    }
                    newclient.SaveChanges();
                    clientid = newclient.RecordId;

                }
                else
                {
                    clientid = Int32.Parse(clientQS.FieldValue("client_clientid"));
                }
                Record BookRec = new Record("Booking");
                BookRec.SetField("book_status", "InProgress");
                BookRec.SetField("book_agency", companyid);
                BookRec.SetField("book_client", clientid);
                BookRec.SetField("book_stage", "Booking_Confirmed");
                BookRec.SetField("book_name", Dispatch.ContentField("HIDDEN_FileName").Substring(0, Dispatch.ContentField("HIDDEN_FileName").Length- 4) + " " + bookingline.Substring(112, 30).Trim(' '));
                BookRec.SetField("book_Schedid", bookingline.Substring(18, 12).Trim(' '));
                BookRec.SetField("book_Billedby", "Works");

                BookRec.SaveChanges();
                int recid = BookRec.RecordId;

                string WriteEntity = "WorkFlowInstance";
                string[] parms = new string[] { "WkIn_WorkflowId", "WkIn_CurrentEntityId", "WkIn_CurrentRecordId", "WkIn_CurrentStateId" };
                string[] values = new string[] { "13", "10234", recid.ToString(), "1075" };

                Record WFRecord = new Record(WriteEntity);
                for (int t = 0; t < parms.Length; t++)
                {
                    WFRecord.SetField(parms[t].ToString(), values[t].ToString());

                }
                WFRecord.SaveChanges();
                int wfid = WFRecord.RecordId;

                Record bookagain = FindRecord("Booking", "Book_bookingid = '" + recid.ToString() + "'");
                bookagain.SetField("book_workflowid", wfid.ToString());
                bookagain.SaveChanges();
                i = bookagain.RecordId;
            }
            else
            {
                string pegage = bookingline.Substring(12, 6);
                string countstr = bookingline.Substring(185, 5);
                count = Int32.Parse(countstr);
                string compSQL = "Select * from Company where comp_pegid = '" + pegage + "' and comp_deleted is null";
                QuerySelect compQS = GetQuery();
                compQS.SQLCommand = compSQL;
                compQS.ExecuteReader();
                if (compQS.Eof())
                {
                    AddError("This Agency is either not in the database or has not been assigned a pegID. Line number " + countall.ToString());
                    throw new Exception();
                }
                string companyid = compQS.FieldValue("comp_companyid");

                string clientnab = bookingline.Substring(42, 50).TrimEnd();
                string sqlclient = "select * from Client where client_name like '" + clientnab + "%'";
                QuerySelect clientQS = GetQuery();
                clientQS.SQLCommand = sqlclient;
                clientQS.ExecuteReader();
                int clientid = 0;
                if (clientQS.Eof())
                {
                    AddError("This Client is either not in the database or the name in the nab does not match the name in crm. Line number " + countall.ToString());
                    throw new Exception();
                }
                else
                {
                    clientid = Int32.Parse(clientQS.FieldValue("client_clientid"));
                }

                string booksql = "Select * from Booking where book_agency = '" + companyid.ToString() + "' and book_Client = '" + clientid.ToString() + "'and book_name = '" + Dispatch.ContentField("HIDDEN_FileName").Substring(0, Dispatch.ContentField("HIDDEN_FileName").Length - 4) + " " + bookingline.Substring(112, 30).Trim(' ') + "' and book_SchedID = '" + bookingline.Substring(18, 12).Trim(' ') + "'";
                QuerySelect bookselect = GetQuery();
                bookselect.SQLCommand = booksql;
                bookselect.ExecuteReader();

                if (bookselect.Eof())
                {
                    AddError("This Booking is either not in the database or the name in the booking does not match the name in crm. Line number " + countall.ToString());
                    throw new Exception();
                }
                else
                {
                    i = Int32.Parse(bookselect.FieldValue("book_bookingid"));
                }
            }
                string ins = bookingline.Substring(190, 1);
                AddContent("Count is " + count);
                bookfin = i.ToString();
                if (ins == "Y")
                {
                    countall++;
                    handleInst(sr.ReadLine(), i);
                }
                while (count > 0)
                {
                    countall++;
                    handledetails(sr.ReadLine(), i);
                    count--;
                }


                
            }



        private void handleInst(string ins, int booki)
        {
           
            string count = ins.Substring(12, 3);
            int cou = Int32.Parse(count);
            string ing = "";
            AddContent(cou.ToString());
            while (cou > 0)
            {
                ing += sr.ReadLine();
                countall++;
                cou--;
                AddContent(ing + "GH");
            }
            AddContent(ing + "GH");
            Record bookrec = FindRecord("Booking", "Book_bookingid = '" + booki + "'");
            AddContent(ing + "Gg");
            bookrec.SetField("book_PlanNotes", ing);
            AddContent(ing + "Gf");
            bookrec.SaveChanges();
            AddContent("INS COMPLETE");
        }

        private void handledetails(string details, int booki)
        {
            string stat = details.Substring(2, 10);
            stat = stat.Trim(' ');
            string pubcode = details.Substring(22, 2);
            string pubsql = "Select * from publications where pblc_Nab like '%" + pubcode + "%'";
            QuerySelect pubselect = GetQuery();
            pubselect.SQLCommand = pubsql;
            pubselect.ExecuteReader();

            if (pubselect.Eof())
            {
                AddError("This Publication is either not in the database or has not been assigned a pegID. Line number " + countall.ToString());
                throw new Exception();
            }
            string pubid = pubselect.FieldValue("pblc_publicationsid");
            string date = details.Substring(34, 8);
            string year = date.Substring(0, 4);
            string month = date.Substring(4, 2);
            string day = date.Substring(6, 2);
            string finaldate = year + "-" + month + "-" + day + " 00:00:00.000";
            string sudodate = year + "-" + month + "-" + day + " 12:00:00.000";

            string color = "";
            string section = "";
            string key = details.Substring(52, 12);
            key = key.Replace(" ", "");

            if (stat == "NEW" || stat == "New")
            {
                string amount = details.Substring(159, 10);
                string money = amount.Substring(0, 8) + "." + amount.Substring(8, 2);
                AddContent(money + "CASH MONEY");
                string colorcode = details.Substring(133, 2);
                if (colorcode != "FC" && colorcode != "07") color = "NoColor";
                else color = "Color";
                string com = details.Substring(171, 5);
                string comdi = com.Substring(0, 3) + "." + com.Substring(2, 2);
                string caption = details.Substring(95, 30);
                string pos = details.Substring(125, 4);
                caption = caption.Replace(' ', ' ');
                string cols = details.Substring(135, 3);
                double coloumns = Double.Parse(cols);
                string cm = details.Substring(138, 3);
                double height = Double.Parse(cm);
                Record pb = new Record("Planbuilder");
                pb.SetField("pnbr_publications", pubid);
                DateTime dt = Convert.ToDateTime(finaldate);
                pb.SetField("pnbr_date", dt.ToString());
                pb.SetField("pnbr_keynumber", key);
                pb.SetField("pnbr_color", color);
                pb.SetField("pnbr_keynumber", key);
                pb.SetField("pnbr_total",amount );
                pb.SetField("pnbr_caption", caption);
                pb.SetField("pnbr_width", coloumns);
                pb.SetField("pnbr_height", height);

                pos = pos.Trim();
                if (pos!= "")
                {
                    string secsql = "select * from sections where sctn_publicationID = '" + pubid.ToString() + "' and sctn_name = 'as instructed' and sctn_deleted is null";
                    QuerySelect secselect = GetQuery();
                    secselect.SQLCommand = secsql;
                    secselect.ExecuteReader();

                    if (!secselect.Eof())
                    {
                        pb.SetField("pnbr_sections", secselect.FieldValue("sctn_sctn_sectionid"));
                    }
                }
                pb.SetField("pnbr_nabcommission", comdi);
                pb.SetField("pnbr_plan", booki);
                pb.SaveChanges();

            }
            else if (stat == "AMEND" || stat == "Amend")
            {
                string amount = details.Substring(146, 10);
                string money = amount.Substring(0, 8) + "." + amount.Substring(8, 2);
                string caption = details.Substring(95, 30);
                string pos = details.Substring(125, 4);
                caption = caption.Replace(' ', ' ');
                string cols = details.Substring(135, 3);
                double coloumns = Double.Parse(cols);
                string cm = details.Substring(138, 3);
                double height = Double.Parse(cm);
                string colorcode = details.Substring(133, 2);
                if (colorcode != "FC" && colorcode != "07") color = "NoColor";
                else color = "Color";
                string com = details.Substring(171, 5);
                string comdi = com.Substring(0, 3) + "." + com.Substring(2, 2);
                pos = pos.Trim();
                int secid = 0;

                if (pos != "")
                {
                    string secsql = "select * from sections where sctn_publicationID = '" + pubid.ToString() + "' and sctn_name = 'as instructed' and sctn_deleted is null";
                    QuerySelect secselect = GetQuery();
                    secselect.SQLCommand = secsql;
                    secselect.ExecuteReader();

                    if (!secselect.Eof())
                    {
                       secid = Int32.Parse(secselect.FieldValue("sctn_sctn_sectionid"));
                        
                    }
                }
                string where = "pnbr_publications ='" + pubid.ToString() + "'and pnbr_color = '" + color + "' and pnbr_date = '" + sudodate + "' and  pnbr_plan = '" + booki.ToString() + "'";
                if (secid != 0) where += "and pnbr_sections = '" + secid.ToString() + "'";
                string booksql = "Select * from planbuilder where " + where;
                QuerySelect planque = GetQuery();
                planque.SQLCommand = booksql;
                planque.ExecuteReader();
                if (planque.Eof())
                {
                    AddError("This Detail is either not in the database or the details in the file are incorrect.  Line number " + countall.ToString());
                    throw new Exception();
                }
                Record planb = FindRecord("Planbuilder", "pnbr_pnbr_planbuilderid = '" + planque.FieldValue("pnbr_pnbr_planbuilderid") + "'");
                string newday = details.Substring(42, 8);
                newday = newday.Replace(" ", "");
                if (newday != "")
                {
                    string yearn = newday.Substring(0, 4);
                    string monthn = newday.Substring(4, 2);
                    string dayn = newday.Substring(6, 2);
                    finaldate = yearn + "-" + monthn + "-" + dayn + " 00:00:00.000";
                }
                // Note: why is this date methoid not implemented for new bookings....
                string newkey = details.Substring(64, 12);
                newkey = newkey.Replace(" ", "");
                if (newkey != "")
                {
                    planb.SetField("pnbr_keynumner", newkey);
                }
                DateTime dt = Convert.ToDateTime(finaldate);
                planb.SetField("pnbr_date", dt);
                planb.SetField("pnbr_total", amount);
                planb.SetField("pnbr_caption", caption);
                planb.SetField("pnbr_width", coloumns);
                planb.SetField("pnbr_nabcommission", comdi);
                planb.SetField("pnbr_height", height);
                planb.SaveChanges();

            }

            else
            {
                string SQLpb = "pnbr_publications ='" + pubid.ToString() + "' and pnbr_date = '" + sudodate + "' and  pnbr_plan = '" + booki.ToString() + "'";
                string booksql = "Select * from planbuilder where " + SQLpb;
                QuerySelect planque = GetQuery();
                planque.SQLCommand = booksql;
                planque.ExecuteReader();
                if (planque.Eof())
                {
                    AddError("This Detail is either not in the database or the details in the file are incorrect.  Line number " + countall.ToString());
                    throw new Exception();
                }
                Record deleteme = FindRecord("Planbuilder", "pnbr_pnbr_planbuilderid = '" + planque.FieldValue("pnbr_pnbr_planbuilderid") + "'");
                deleteme.SetField("pnbr_deleted", "1");
                deleteme.SaveChanges();
            }
            }



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


        public string GetLibraryPath()
        {
            string Path = "";
            Record RecPath = FindRecord("Custom_SysParams", "parm_name = 'DocStore'");
            Path = RecPath.GetFieldAsString("Parm_Value");
            return Path;
        }

    }
}
