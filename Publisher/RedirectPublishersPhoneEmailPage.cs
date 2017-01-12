using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;


namespace NZPACRM
{
    class RedirectPublishersPhoneEmailPage : Web
    {
        public RedirectPublishersPhoneEmailPage()
        { 
        
        }
        public override void BuildContents()
        {
            try
            {
                AddContent(HTML.Form());
                GetTabs("Publishers");
                string sPublisherID = "";
                string Key58 = "";

                #region Get Publication id
                if (!String.IsNullOrEmpty(Dispatch.EitherField("pbls_PublishersID")))
                {
                    sPublisherID = Dispatch.EitherField("pbls_PublishersID");
                    Key58 = Dispatch.EitherField("pbls_PublishersID");
                }
                else
                {
                    Key58 = Dispatch.EitherField("Key58");
                    sPublisherID = Dispatch.EitherField("Key58");
                }
                #endregion
                
                string sURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/do?SID=" + Dispatch.EitherField("SID") + "&Act=432&&Mode=1&Key0=58" + "&Key58=" + Key58 + "&pbls_PublishersID=" + sPublisherID + "&dotnetdll=NZPACRM&dotnetfunc=RunPublisherPhoneEmailPage&J=Phone/E-mail&T=Publishers";//  UrlDotNet(ThisDotNetDll, "RunClientAddressList") + "&J=Addresses&T=Client";
            //http://grey014/etl_crm/eware.dll/Do?SID=134408729743713&Act=432&Mode=1&CLk=T&Key0=58&Key58=9&pbls_PublishersID=9&dotnetdll=NZPACRM&dotnetfunc=RunPublisherPhoneEmailPage&J=Phone/E-mail&T=Publishers
                Dispatch.Redirect(sURL);
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
        }
    }
}
