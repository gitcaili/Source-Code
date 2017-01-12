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
    class PlanCommunication : ListPage
    {
        public PlanCommunication()
            : base("Booking", "CommunicationPlanGrid", "PlanCommunicationFilterBox")
        {
            AddTopContent(GetCustomEntityTopFrame("Booking"));
           
            #region Add HTML Form so that Navigation will work as expected
            AddContent(HTML.Form());
            #endregion
        }

        public override void BuildContents()
        {
            
            string sPlanId = "";
            
            if (!String.IsNullOrEmpty(Dispatch.EitherField("book_bookingid")))
                sPlanId = Dispatch.EitherField("book_bookingid");
            else
                sPlanId = Dispatch.EitherField("Key58");
            
            #region Add Filter to the List
            List objCommunicationList = new List("CommunicationPlanGrid");
            this.ResultsGrid.RowsPerScreen = 10; 
            this.ResultsGrid.Filter = "comm_bookingid=" + sPlanId;
            #endregion

            base.BuildContents();
            #region Add New Button with List
            AddUrlButton("New Task", "newtask.png", Url("361") + "&book_bookingid=" + sPlanId);
            AddUrlButton("New Appointment", "newappointment.png", Url("362") + "&book_bookingid=" + sPlanId);
            #endregion
            
        }

        public override void AddNewButton()
        {
            //'base.AddNewButton();
        }
    }
}
