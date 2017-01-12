using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;

namespace NZPACRM
{
    public class RedirectPublicationPhoneEmailPage : Web
    {
        public RedirectPublicationPhoneEmailPage()
        { 
        }
        public override void BuildContents()
        {
            try
            {
                AddContent(HTML.Form());
                GetTabs("Publications");
                string sPublication = "";
                string Key58 = "";   

                #region Get Publication id
                if (!String.IsNullOrEmpty(Dispatch.EitherField("pblc_PublicationsID")))
                {
                    sPublication = Dispatch.EitherField("pblc_PublicationsID");
                    Key58 = Dispatch.EitherField("pblc_PublicationsID");
                }
                else
                {
                    Key58 = Dispatch.EitherField("Key58");
                    sPublication = Dispatch.EitherField("Key58");
                }
                #endregion                
                //AddContent("sClientID=" + sPublication + "Key58=" + Key58);
                string sURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/do?SID=" + Dispatch.EitherField("SID") + "&Act=432&&Mode=1&Key0=58" + "&Key58=" + Key58 + "&pblc_PublicationsID=" + sPublication +  "&dotnetdll=NZPACRM&dotnetfunc=RunPublicationPhoneEmail&J=Phone/E-mail&T=Publications";//  UrlDotNet(ThisDotNetDll, "RunClientAddressList") + "&J=Addresses&T=Client";
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
