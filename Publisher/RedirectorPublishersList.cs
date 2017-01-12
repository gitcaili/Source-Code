using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.Blocks;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using Sage.CRM.HTML;
using Sage.CRM.Utils;
using Sage.CRM.WebObject;
using Sage.CRM.UI;
using NZPACRM.Common;

namespace NZPACRM
{
    public class RedirectorPublishersList : Web
    {
        public RedirectorPublishersList()
        { 
        }
        public override void BuildContents()
        {
            AddContent(HTML.Form());
            GetTabs("Publishers");
            string sPublishrId = "";
            string sCompanyid = "";
            string sPersonId = "";
            #region Set Publisher ID
            if (!String.IsNullOrEmpty(Dispatch.EitherField("pbls_PublishersID")))
                sPublishrId = Dispatch.EitherField("pbls_PublishersID");
            else
                sPublishrId = Dispatch.EitherField("Key58");
            #endregion
            #region Set Company Id
            if (!String.IsNullOrEmpty(Dispatch.EitherField("comp_companyid")))
                sCompanyid = Dispatch.EitherField("comp_companyid");
            else
                sCompanyid = Dispatch.EitherField("Key1");
            #endregion

            #region Set Person Id
            if (!String.IsNullOrEmpty(Dispatch.EitherField("comp_primarypersonid")))
                sPersonId = Dispatch.EitherField("comp_primarypersonid");
            else
                sPersonId = Dispatch.EitherField("Key2");
            #endregion
            string sURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/do?SID=" + Dispatch.EitherField("SID") + "&Act=432&Key0=58" + "&Key1=" + sCompanyid + "&Key2=" + sPersonId + "&pbls_PublishersID=" + sPublishrId + "&Key58=" + sPublishrId + "&comp_companyid=" + sCompanyid + "&comp_primarypersonid=" + sPersonId + "&func=baseUrl&dotnetdll=NZPACRM&dotnetfunc=RunPublisherAddressList&J=Addresses";//  UrlDotNet(ThisDotNetDll, "RunClientAddressList") + "&J=Addresses&T=Client";
            Dispatch.Redirect(sURL);
        }
    }
    
}
