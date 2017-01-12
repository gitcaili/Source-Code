using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;
using System.Data;
using System.IO;

namespace NZPACRM.Plan
{
    public class PlanDeletePage : DataPageDelete
    {
        int iPlanID = 0;
        public PlanDeletePage()
            : base("Booking", "book_bookingid", "BookingNewEntry")
        {
            AddTopContent(GetCustomEntityTopFrame("Booking"));
            if (!String.IsNullOrEmpty(Dispatch.EitherField("book_bookingid")))
            {
                iPlanID = Convert.ToInt32(Dispatch.EitherField("book_bookingid"));
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key37")))
            {
                iPlanID = Convert.ToInt32(Dispatch.EitherField("Key37"));
            }
            
            AddContent("<script type='text/javascript' src='../CustomPages/Booking/ClientFuncs.js'></script>");
            this.SaveMethod = "RunPlanFindPage&book_bookingid=" + iPlanID;
            this.CancelMethod = "RunPlanSummaryPage&book_bookingid=" + iPlanID;
            GetTabs("Booking");
        }

        public override void BuildContents()
        {
            try
            {
                AddContent(HTML.Form());

                if (!String.IsNullOrEmpty(Dispatch.EitherField("HiddenMode")))
                {
                    if (Dispatch.EitherField("HiddenMode") == "delete")
                    {
                        Record objPlanRecordDel = FindRecord("Booking", "book_bookingid=" + iPlanID);
                        objPlanRecordDel.SetField("book_Deleted", 1);
                        objPlanRecordDel.SaveChanges();

                        Dispatch.Redirect(UrlDotNet(ThisDotNetDll, "RunPlanSearchPage&J=Booking&T=find"));
                    }
                }
                EntryGroup objPlanSummary = new EntryGroup("BookingNewEntry");
                Record objPlanRecord = FindRecord("Booking", "book_bookingid=" + iPlanID);
                AddContent(objPlanSummary.GetHtmlInViewMode(objPlanRecord));

                AddContent(HTML.InputHidden("HiddenMode", ""));
                AddUrlButton("Confirm Delete", "delete.gif", "javascript:document.EntryForm.HiddenMode.value='delete';document.EntryForm.submit();");
                AddUrlButton("Cancel", "cancel.gif", UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage") + "&J=Booking Summary&T=Booking&book_bookingid=" + iPlanID);

            }
            catch (Exception Ex)
            {
 
            }
        }

    }
}
