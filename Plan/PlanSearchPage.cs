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
    public class PlanSearchPage : SearchPage
    {
        public PlanSearchPage()
            :base("BookingSearchBox","BookingGrid")
        {
            this.ResultsGrid.RowsPerScreen = 10;
            this.SavedSearch = true;
        }

        public override void  BuildContents()
        {
            if (!String.IsNullOrEmpty(Dispatch.EitherField("RecentValue")))
            {
               Dispatch.Redirect("/" + Dispatch.InstallName + "/eware.dll/Do?SID=" + Dispatch.EitherField("SID") + "&Act=432&Mode=1&CLk=T&Key0=58&func=datapageurl&dotnetdll=NZPACRM&dotnetfunc=RunPlanSearchPage&T=Find&J=Booking");
            }
            base.BuildContents();
        }
    }
}
