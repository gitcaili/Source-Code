using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;

namespace NZPACRM
{
    public class RedirectorPublicationList : Web
    {
        public RedirectorPublicationList()
        { 
            
        }
        public override void BuildContents()
        {
            AddContent(HTML.Form());
            GetTabs("Publications");
            string sPublicationId = "";
            string sCompanyid = "";
            string sPersonId = "";
            #region Get Publication id
            if (!String.IsNullOrEmpty(Dispatch.EitherField("pblc_PublicationsID")))
                sPublicationId = Dispatch.EitherField("pblc_PublicationsID");
            else
                sPublicationId = Dispatch.EitherField("Key58");
            #endregion
            #region Get Company ID
            if (!String.IsNullOrEmpty(Dispatch.EitherField("comp_companyid")))
                sCompanyid = Dispatch.EitherField("comp_companyid");
            else
                sCompanyid = Dispatch.EitherField("Key1");
            #endregion
            #region Get Person Id
            if (!String.IsNullOrEmpty(Dispatch.EitherField("comp_primarypersonid")))
                sPersonId = Dispatch.EitherField("comp_primarypersonid");
            else
                sPersonId = Dispatch.EitherField("Key2");
            #endregion

            string sURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/do?SID=" + Dispatch.EitherField("SID") + "&Act=432&Key0=58" + "&Key1=" + sCompanyid + "&Key2=" + sPersonId + "&pblc_PublicationsID=" + sPublicationId + "&Key58=" + sPublicationId + "&comp_companyid=" + sCompanyid + "&comp_primarypersonid=" + sPersonId + "&func=baseUrl&dotnetdll=NZPACRM&dotnetfunc=RunPublicationAddressList&J=Addresses";//  UrlDotNet(ThisDotNetDll, "RunClientAddressList") + "&J=Addresses&T=Client";
            Dispatch.Redirect(sURL);
        }
    }
}
