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
using System.Globalization;

namespace NZPACRM.Plan
{
    class ReactivatePlanPage : Web
    {
        CRMHelper objHelper = new CRMHelper();
        string sPlanId = "";
        public ReactivatePlanPage()
        { 
        }

        public override void BuildContents()
        {
            try
            {
                if (!String.IsNullOrEmpty(Dispatch.EitherField("book_BookingID")))
                {
                    sPlanId = Dispatch.EitherField("book_BookingID");
                }
               
                if (sPlanId != "")
                {
                    Record recPlanRecord = FindRecord("Booking", "book_revisedBookId =" + sPlanId);
                    if (!recPlanRecord.Eof())
                    {
                        while (!recPlanRecord.Eof())
                        {
                            Record recPlan = FindRecord("Booking", "book_BookingID =" + recPlanRecord.GetFieldAsString("book_BookingID"));
                            recPlan.SetField("book_Status", "InActive");
                            recPlan.SaveChanges();
                            recPlanRecord.GoToNext();
                        }
                    }

                    Record CurrPlanRec = FindRecord("Booking", "book_BookingID =" + sPlanId);
                    CurrPlanRec.SetField("book_Status", "InProgress");
                    CurrPlanRec.SaveChanges();

                    bool wrkflwResult;

                    wrkflwResult = objHelper.ProgressWorkflow(sPlanId, "Booking", "Booking Workflow", "Agency");
                    if (wrkflwResult)
                    {
                        AddInfo("Workflow progressed successfully to InProgress");

                        objHelper.SetStageStatus("Booking", sPlanId, "", "InProgress");
                    }
                    else
                    {
                        AddInfo("Error Occurred during worlflow progress");
                    }         

                    Dispatch.Redirect(UrlDotNet(ThisDotNetDll, "RunPlanSummaryPage"));
                }
            }                    
            catch (Exception ex)
            { }
        }
    }    
}
