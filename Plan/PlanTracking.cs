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
   public class PlanTracking:ListPage
    {
       public PlanTracking()
           : base("Booking", "BookingProgressList", "BookingProgressScreen")
       {
           AddTopContent(GetCustomEntityTopFrame("Booking"));

           #region Add HTML Form so that Navigation will work as expected
           AddContent(HTML.Form());
           #endregion
           
       }
       public override void BuildContents()
       {
           try
           {
               string sPlanID = "";
               if (!String.IsNullOrEmpty(Dispatch.EitherField("book_bookingid")))
                   sPlanID = Dispatch.EitherField("book_bookingid");
               else
                   sPlanID = Dispatch.EitherField("Key58");

               this.ResultsGrid.RowsPerScreen = 10;
               #region Add Filter to the List
               this.ResultsGrid.Filter = "book_bookingid=" + sPlanID;
               #endregion
               base.BuildContents();
               base.OnLoad = "javascript:SetPlanTopContext();";
               base.OnLoad = "javascript:HideFilterButton();";
               
           }
           catch (Exception Ex)
           {
               this.AddError(Ex.Message);
           }
           
       }

       public override void AddNewButton()
       {
           //base.AddNewButton();
       } 
    }
}
