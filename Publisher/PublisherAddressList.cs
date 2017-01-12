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
    public class PublisherAddressList : ListPage
    {
        public PublisherAddressList()
            : base("Publishers", "PublisherAddressList", "PublishersSummaryScreen")
        {
            #region Set js file reference path
            AddContent("<script type='text/javascript' src='../CustomPages/Client/ClientFuncs.js'></script>");
            #endregion
        }
        public override void BuildContents()
        {
            try
            {
                string sPublicationId = "";
                if (!String.IsNullOrEmpty(Dispatch.EitherField("pbls_PublishersID")))
                    sPublicationId = Dispatch.EitherField("pbls_PublishersID");
                else
                    sPublicationId = Dispatch.EitherField("Key58");
                
                #region Add HTML Form so that Navigation will work as expected
                AddContent(HTML.Form());
                #endregion

                List objAddressList = new List("PublisherAddressList");
                objAddressList.Filter = " adli_publisherid=" + sPublicationId;

                #region Add New Button with List
                AddUrlButton("New Address", "new.gif", UrlDotNet(ThisDotNetDll, "RunPublisherAddressNew"));

                #endregion

                AddContent(objAddressList);
                GetTabs("Publishers");

            }
            catch (Exception error)
            {
                this.AddError(error.Message);
            }

        }
    }
}
