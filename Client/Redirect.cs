using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;

namespace NZPACRM
{
    public class Redirect : Web
    {
        public Redirect()
        {


        }

        public override void BuildContents()
        {
            string sCompanyid = "";
            string sPersonId = "";
            AddContent(HTML.Form());
            GetTabs("Client");

            string sClientId = "";
            if (!String.IsNullOrEmpty(Dispatch.EitherField("client_ClientID")))
                sClientId = Dispatch.EitherField("client_ClientID");
            else
                sClientId = Dispatch.EitherField("Key58");

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

            string sURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/do?SID=" + Dispatch.EitherField("SID") + "&Act=432&Key0=58" + "&Key1=" + sCompanyid + "&Key2=" + sPersonId + "&client_clientid=" + sClientId + "&Key58=" + sClientId + "&comp_companyid=" + sCompanyid + "&comp_primarypersonid=" + sPersonId + "&func=baseUrl&dotnetdll=NZPACRM&dotnetfunc=RunClientAddressList&J=Addresses";//  UrlDotNet(ThisDotNetDll, "RunClientAddressList") + "&J=Addresses&T=Client";
            Dispatch.Redirect(sURL);

        }
    }
}
