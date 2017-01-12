using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;

namespace NZPACRM
{
    public class RedirectClientPhoneEmailPage : Web
    {
        public RedirectClientPhoneEmailPage()
        { 
        }
        public override void BuildContents()
        {
            try
            {
                AddContent(HTML.Form());
                GetTabs("Client");
                string sClientID = "";
                string Key58 = "";   

                #region Get Publication id
                if (!String.IsNullOrEmpty(Dispatch.EitherField("client_ClientID")))
                {
                    sClientID = Dispatch.EitherField("client_ClientID");
                    Key58 = Dispatch.EitherField("client_ClientID");
                }
                else
                {
                    Key58 = Dispatch.EitherField("Key58");
                    sClientID = Dispatch.EitherField("Key58");
                }
                #endregion                
                AddContent("sClientID=" + sClientID + "Key58=" + Key58);
                string sURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/do?SID=" + Dispatch.EitherField("SID") + "&Act=432&&Mode=1&Key0=58" + "&Key58=" + Key58 + "&client_ClientID=" + sClientID +  "&dotnetdll=NZPACRM&dotnetfunc=RunClientPhoneEmail&J=Phone/E-mail&T=Client";//  UrlDotNet(ThisDotNetDll, "RunClientAddressList") + "&J=Addresses&T=Client";
                AddContent(sURL);
                Dispatch.Redirect(sURL);
                
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
        }
    }
}
