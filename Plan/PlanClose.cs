using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;
using System.Data;
using System.IO;
using Sage.CRM.UI;
using System.Diagnostics;
using System.Net.Mail;
using System.Net;

namespace NZPACRM.Plan
{
    public class PlanClose : Web
    {
        CRMHelper objCrmHelp = new CRMHelper();

        string sBook_bookingid = "";
        string sToEmailAddress = "";
        string sFromEmailAddress = "";

        string smtpusername = "";
        string smtppwd = "";
        string servername = "";
        string smtpport = "";
        string hMode = "";

        public PlanClose()
            : base()
        {
            if (!string.IsNullOrEmpty(Dispatch.EitherField("book_BookingID")))
                sBook_bookingid = Dispatch.EitherField("book_BookingID");

            else if (!string.IsNullOrEmpty(Dispatch.EitherField("Key58")))
                sBook_bookingid = Dispatch.EitherField("Key58");

            else if (!string.IsNullOrEmpty(Dispatch.EitherField("Key37")))
                sBook_bookingid = Dispatch.EitherField("Key37");

            //Get STMP Detail
            objCrmHelp.getSMTPDetails();

            smtpusername = objCrmHelp.smtpusername;
            smtppwd = objCrmHelp.smtppwd;
            servername = objCrmHelp.servername;
            smtpport = objCrmHelp.smtpport;
            getSMPTPassword();

            //Hidden Mode
            
            if (!String.IsNullOrEmpty(Dispatch.EitherField("HiddenMode")))
            {
                hMode = Dispatch.EitherField("HiddenMode");
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
        public override void BuildContents()
        {
            AddContent(HTML.Form());
            EntryGroup entryBookClosePlan = new EntryGroup("BookClosePlan");
            //display selection field with options
            //if misc show multiline text box with 60 chara
            //Save Data & set workflow
            //send mail with selected option
            if (hMode == "save")
            {
                try
                {
                    Record objExisBookRec = FindRecord("Booking", "book_bookingid=" + sBook_bookingid);
                    if (!objExisBookRec.Eof())
                    {
                        objExisBookRec.SetField("book_closecode", Dispatch.ContentField("book_closecode"));
                        objExisBookRec.SetField("book_misc", Dispatch.ContentField("book_misc"));
                        objExisBookRec.SetField("book_status", "Closed");
                        objExisBookRec.SetField("book_stage", "Planclosed");
                        objExisBookRec.SaveChanges();
                        AddInfo("Plan successfully closed");

                        //Set WorkFlow
                        if (!objCrmHelp.ProgressWorkflow(sBook_bookingid, "Booking", "Booking Workflow", "Quote Closed"))
                        {
                            AddError("Error Occurred during workflow progress");
                        }
                        else
                        {
                            objCrmHelp.SetStageStatus("Booking", sBook_bookingid, "Planclosed", "Closed");
                        }
                        //Send Email
                        if (!SendEmail("amit.pardhe@greytrixindia.com", "admin@panoply-tech.com", "ClosePlan", "Booking"))
                        {
                            AddError("Email sending fail");
                        }
                        else
                        {
                            Dispatch.Redirect(UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage&=" + sBook_bookingid));
                        }                        
                        //entryBookClosePlan.Fill(objExisBookRec);
                        //base.OnLoad = "javascript:hideBookMisc('" + Dispatch.ContentField("book_closecode") + "')";
                        //entryBookClosePlan.GetHtmlInViewMode(objExisBookRec);
                        //AddContent(entryBookClosePlan);
                    }
                }
                catch (Exception ex)
                {
                    AddError(ex.Message.ToString());
                }
            }
            else
            {
                AddContent(entryBookClosePlan.GetHtmlInEditMode());
            }
            AddContent(HTML.InputHidden("HiddenMode", ""));
            AddSubmitButton("Save", "save.gif", "javascript:document.EntryForm.HiddenMode.value='save';document.EntryForm.submit();");
            AddUrlButton("Back to Plan Summary", "back.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage&=" + sBook_bookingid));
        }

        public bool SendEmail(string ToEmailAddress, string FromEmailAddress, string TemplateName, string EntityName)
        {
            
            if (ToEmailAddress.ToString() == "")
            {
                AddError("Unable To Send Email: To Email address is not available.");
            }

            if (FromEmailAddress.ToString() == "")
            {
                AddError("Unable To Send Email: From Email address is not available.");
            }

            StringBuilder sb = new StringBuilder();
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress(FromEmailAddress);
            msg.To.Add(ToEmailAddress);
            msg.Subject = "Close Plan";
            sb.Append("<br>");
            Record ObjEmailTemplate = FindRecord("EmailTemplates", "EmTe_Name='"+TemplateName+"' and EmTe_Entity='"+EntityName+"'");
            if (!ObjEmailTemplate.Eof())
                msg.Body = ObjEmailTemplate.GetFieldAsString("EmTe_Comm_Email");

            msg.IsBodyHtml = true;
            System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(servername);
            
            //string sFileName = GetLibraryPath() + "VisitServiceReports\\" + objFileRec.GetFieldAsString("libr_FileName");
            if (false)
            {
                string sFileName = "";
                try
                {
                    if (sFileName != null)
                        msg.Attachments.Add(new Attachment(sFileName));
                }
                catch (Exception ex)
                {
                    AddError("Unable To Send Attachment");
                    AddContent(ex.Message.ToString());
                    AddUrlButton("Continue", "Continue.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage" + "&book_bookingid=" + sBook_bookingid + ""));
                }
            }
            try
            {                
                var _with1 = smtpClient;
                smtpClient.Port = Convert.ToInt32(smtpport);
                smtpClient.Credentials = new NetworkCredential(smtpusername, smtppwd);
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.EnableSsl = true;
                
                _with1.Send(msg);
                return true;
            }
            catch (Exception Ex)
            {                
                AddError(Ex.Message.ToString());
                return false;
            }
        }
    }
}
