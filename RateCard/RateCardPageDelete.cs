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
    public class RateCardPageDelete : DataPageDelete
    {
        int iRateCardID = 0;
        public RateCardPageDelete()
            : base("rate_RatesCardID", "rate_RatesCardID", "RatesCardSummaryScreen")
        {
            
            #region Sub Section ID
            if (!String.IsNullOrEmpty(Dispatch.EitherField("rate_RatesCardID")))
            {
                iRateCardID = Convert.ToInt32(Dispatch.EitherField("rate_RatesCardID"));
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key37")))
            {
                iRateCardID = Convert.ToInt32(Dispatch.EitherField("Key37"));
            }
            #endregion

            this.CancelMethod = "RunPlanPage&rate_RatesCardID=" + iRateCardID;
        }

        public override void BuildContents()
        {
            base.BuildContents();
        }

        public override void AfterSave(EntryGroup screen)
        {
            Dispatch.Redirect(UrlDotNet(ThisDotNetDll, "RunPlanPage") + "&rate_RatesCardID=" + iRateCardID);

            base.AfterSave(screen);
        }
    }
}
