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
    class PlanReffralPage : Web
    {
        public PlanReffralPage()
        {
            GetTabs("Booking");
        }

        public override void BuildContents()
        {
            try
            {
                AddContent(HTML.Form());

                base.OnLoad = "javascript:OnloadScriptToChangeLink();";

                List objPlanRefferal = new List("PlanList");
                objPlanRefferal.Filter = "book_revisedBookId = " + Dispatch.EitherField("book_BookingID");                
                AddContent(objPlanRefferal);
            }
            catch (Exception ex)
            { 
            
            }
        }
    }

}
