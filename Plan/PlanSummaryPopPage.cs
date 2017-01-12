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
    public class PlanSummaryPopPage : DataPage
    {
        private EntryGroup objPlanSummaryBox;
        int iPLanId = 0;

        public PlanSummaryPopPage()
            : base("Booking", "book_bookingid")
        {
            #region Get current Booking id

            if (!String.IsNullOrEmpty(Dispatch.EitherField("book_bookingid")))
            {
                iPLanId = Convert.ToInt32(Dispatch.EitherField("book_bookingid"));
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key58")))
            {
                iPLanId = Convert.ToInt32(Dispatch.EitherField("Key58"));
            }
            #endregion
        }

        public override void BuildContents()
        {
            try
            {
                #region Bulid UI

                EntryGroup objTopContent = new EntryGroup("BookingTopContent");
                objPlanSummaryBox = AddEntryGroup("BookingNewEntry", "Plan");                

                AddUrlButton("Cancel", "cancel.gif", "javascript:window.close();");
                #endregion
            }

            catch (Exception Ex)
            {
                AddError(Ex.Message);
            }
            base.BuildContents();
        }

        public override void PreBuildContents()
        {
            base.PreBuildContents();
        }

        public override void AddContinueButton()
        {
            //base.AddContinueButton();
        }

        public override void AddEditButton()
        {
            //'base.AddEditButton();
        }
    }
}
