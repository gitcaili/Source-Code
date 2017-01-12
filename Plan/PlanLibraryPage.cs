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
    public class PlanLibraryPage : ListPage
    {
        string sPlanID = "";
        public PlanLibraryPage()
            : base("Library", "PlanLibrarylist", "LibraryFilterBox")
        {
            if (!String.IsNullOrEmpty(Dispatch.EitherField("book_bookingid")))
            {
                sPlanID = Dispatch.EitherField("book_bookingid").ToString();
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key58")))
            {
                sPlanID = Dispatch.EitherField("Key58").ToString();
            }
        }

        public override void BuildContents()
        {
            AddTopContent(GetCustomEntityTopFrame("Booking"));

            int iKey1 = 0;
            int iKey2 = 0;

            if (!String.IsNullOrEmpty(GetContextInfo("Booking", "book_agency")))
                iKey1 = Convert.ToInt32(GetContextInfo("Booking", "book_agency"));

            if (!String.IsNullOrEmpty(GetContextInfo("Booking", "book_contact")))
                iKey2 = Convert.ToInt32(GetContextInfo("Booking", "book_contact"));

            this.ResultsGrid.Filter = "libr_bookingid=" + sPlanID;
            base.BuildContents();
            AddUrlButton("Add File", "FileUpload.gif", Url("343") + "&Key-1=58&Key1=" + iKey1 + "&Key2=" + iKey2 + "&book_bookingid=" + sPlanID+"&Key58="+sPlanID);            
        }

        public override void AddNewButton()
        {
            //base.AddNewButton();
        }
    }
}
