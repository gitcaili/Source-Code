using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;
using System.Data;
using System.IO;

namespace NZPACRM.RateCard
{
    public class RateCardEditPage : DataPageEdit
    {
        string sRatesCardID = "";
        public RateCardEditPage()
            : base("RatesCard", "rate_RatesCardID", "RatesCardSummaryScreen")
        {
            #region Sub Section ID
            if (!String.IsNullOrEmpty(Dispatch.EitherField("suse_subsectionid")))
            {
                sRatesCardID = Dispatch.EitherField("suse_subsectionid");
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key37")))
            {
                sRatesCardID = Dispatch.EitherField("Key37");
            }
            #endregion
            
            this.SaveMethod = "RunPlanPage&rate_RatesCardID=" + sRatesCardID;
            this.CancelMethod = "RunPlanPage&rate_RatesCardID=" + sRatesCardID;
            this.DeleteMethod = "RunRateCardDeletePage&rate_RatesCardID=" + sRatesCardID;
            GetTabs("Publications", "Rate Card");
        }
    }
}
