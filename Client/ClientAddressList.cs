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

namespace NZPACRM
{
    public class ClientAddressList : ListPage
    {
        public ClientAddressList()
            : base("Client", "ClientAddressList", "ClientSummaryScreen")
        {
            GetTabs("Client");

            #region Set js file reference path
            AddContent("<script type='text/javascript' src='../CustomPages/Client/ClientFuncs.js'></script>");
            #endregion
        } 
        public override void BuildContents()
        {
            try
            {
               

                string sClientId = "";
                if (!String.IsNullOrEmpty(Dispatch.EitherField("client_ClientID")))
                    sClientId = Dispatch.EitherField("client_ClientID");
                else
                    sClientId = Dispatch.EitherField("Key58");

            

                #region Add HTML Form so that Navigation will work as expected
                AddContent(HTML.Form());
                #endregion

                List objAddressList = new List("ClientAddressList");
                objAddressList.Filter = " adli_clientid=" + sClientId;

                #region Add New Button with List
                AddUrlButton("New Address", "new.gif", UrlDotNet(ThisDotNetDll, "RunClientAddress"));

                #endregion

                AddContent(objAddressList);
                

            }
            catch (Exception error)
            {
                this.AddError(error.Message);
            }

        }
    }
}
