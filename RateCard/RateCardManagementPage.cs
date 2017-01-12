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
    class RateCardManagementPage : Web
    {
        string RatesCardId = "";
        public RateCardManagementPage()
        {

        }

        public override void BuildContents()
        {
            try
            {
                AddContent(HTML.Form());
                AddTopContent("<img src='/etl_crm/Themes/img/color/Icons/admin.png' hspace='0' border='0' align='TOP' title=''> " + HTML.Span("SpanAdmin", "Administratior", "class='TOPBC'") + " " + "-" + " " + HTML.Span("SpanDataMgt", "Data Management", "class='TOPBC'") + " " + "-" + " " + HTML.Span("SpanDataMgt", "Manage RateCards", "class='TOPBC'"));
                 
                if (!String.IsNullOrEmpty(Dispatch.EitherField("HiddenMode")))
                {
                    RatesCardId = Dispatch.EitherField("HiddenMode");
                }
                //AddContent("RatesCardId =" + RatesCardId);
                EntryGroup entRatesCard = new EntryGroup("RatesCardSearchScreen", "Find RatesCard");
                Record recRatesCard = new Record("RatesCard");
                entRatesCard.Fill(recRatesCard);
                AddContent(entRatesCard.GetHtmlInEditMode());
               
                AddUrlButton("Find", "search.gif", "javascript:FilterRateCardScreen();");
               
                string sBackUrl = "http://" + Dispatch.Host +"/"+ Dispatch.InstallName + "/eware.dll/Do?SID=" + Dispatch.EitherField("SID") + "&Act=1650&Mode=1&CLk=T&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data%20Management";
                AddUrlButton("Back", "PrevCircle.gif", sBackUrl); 
               
                AddContent("<BR>");
                //AddContent("<BR>http://" + Dispatch.Host + "/WebApplication/default.aspx?rate_RatesCardID='" + RatesCardId + "'");
                #region Editable Grid
                //base.OnLoad = "javascript:GetGrid();";
                //AddContent("<table id='tblAppendGrid'></table>");http://localhost:54946/
                //AddContent("<IFRAME  ID='iframePage' src='http://" + Dispatch.Host + "/WebApplication/default.aspx?rate_RatesCardID='" + RatesCardId + "' width='1050px' height='900px' frameborder='0' scrolling='yes' ></IFRAME>");
                //string sUrl = "http://localhost:54946/default.aspx?rate_RatesCardID=" + RatesCardId + "";
                string sUrl = "http://" + Dispatch.Host + "/WebApplication/default.aspx?rate_RatesCardID=" + RatesCardId + "";
                //AddContent("<BR>sUrl=" + sUrl);
                AddContent("<IFRAME  ID='iframePage' src=" + sUrl + " width='1050px' height='900px' frameborder='0' scrolling='yes' ></IFRAME>");
                #endregion
                //AddContent("<input type='button' value='Save'> <BR>");
                #region Hidden Field
                AddContent(HTML.InputHidden("HiddenMode", ""));
                #endregion
            }
            catch (Exception ex)
            {
            }
        }
    }
}
