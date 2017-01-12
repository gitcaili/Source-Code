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
    class RateCardListPage : ListPage
    {
        int iPublicationID = 0;
        public RateCardListPage()
            : base("publications", "PublicationRateCardGrid", "PublicationsSummaryScreen")
        {
            #region Redirect
            if (!String.IsNullOrEmpty(Dispatch.EitherField("F")))
            {
                //AddContent("tab = " + Dispatch.EitherField("F"));
                string sUrl = "http://"+Dispatch.Host+ "/" + Dispatch.InstallName + "//eware.dll/Do?SID="+Dispatch.EitherField("SID")+ "&Act=432&Mode=1&CLk=T&Key0=58";
                        sUrl += "&Key58="+Dispatch.EitherField("pblc_PublicationsID") + "&pblc_PublicationsID=" +Dispatch.EitherField("pblc_PublicationsID");
                        sUrl += "&dotnetdll=NZPACRM&dotnetfunc=RunPlanPage&J=Rate Card&T=Publications";                       
                Dispatch.Redirect(sUrl);
            }
            #endregion
            base.OnLoad = "javascript:HideFilterScreen();";
            
            #region Get current Publications id

           if (!String.IsNullOrEmpty(Dispatch.EitherField("pblc_PublicationsID")))
            {
               iPublicationID = Convert.ToInt32(Dispatch.EitherField("pblc_PublicationsID"));
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key58")))
            {
                iPublicationID = Convert.ToInt32(Dispatch.EitherField("Key58"));
            }
            else
            {
                iPublicationID = Convert.ToInt32(GetContextInfo("publications", "pblc_PublicationsID"));
            }
            #endregion
        }
        public override void BuildContents()
        {            
            this.ResultsGrid.Filter = "rate_PublicationsID=" + iPublicationID;
            base.BuildContents();            
            AddTopContent(GetCustomEntityTopFrame("publications"));
        }
        public override void AddNewButton()
        {
            //base.AddNewButton();
        }
    }
}
