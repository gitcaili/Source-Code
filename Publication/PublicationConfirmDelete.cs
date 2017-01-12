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
    public class PublicationConfirmDelete : Web
    {
        #region Variable Declaration
        string sPublicationId = "";
        string sCompanyid = "";
        string sPersonId = "";
        string sAddressId = "";
        string sAddress1 = "";
        string sAddress2 = "";
        string sAddress3 = "";
        string sAddress4 = "";
        string sCity = "";
        string sState = "";
        string sZipCode = "";
        string sCountry = "";
        string sDefaultAddress = "";
        string sType = "";
        string sHidden = "";
        string sHTML = "";
        string sHiddenDelete = "";
        string sFormAddress1 = "";
        string sFormAddress2 = "";
        string sFormAddress3 = "";
        string sFormAddress4 = "";
        string sFormCity = "";
        string sFormState = "";
        string sFormZipCode = "";
        string sFormCountry = "";
        string sFormDefaultAddress = "";
        string sFormBusinessType = "";
        string sFormBillingType = "";
        string sFormShippingType = "";
        int iAddressId = 0;
        #endregion
        public PublicationConfirmDelete()
        {
            GetTabs("Publications");
            AddTopContent(GetCustomEntityTopFrame("Publications"));
            #region Set Publisher ID
            if (!String.IsNullOrEmpty(Dispatch.EitherField("pblc_PublicationsID")))
                sPublicationId = Dispatch.EitherField("pblc_PublicationsID");
            else
                sPublicationId = Dispatch.EitherField("Key58");
            #endregion
            #region Set Address Id
            if (!string.IsNullOrEmpty(Dispatch.EitherField("addr_addressid")))
                sAddressId = Dispatch.EitherField("addr_addressid");
            else if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_addressid")))
                sAddressId = Dispatch.ContentField("addr_addressid");
            #endregion
            #region Add HTML Form so that Navigation will work as expected
            AddContent(HTML.Form());
            #endregion
            #region Set js file reference path
            AddContent("<script type='text/javascript' src='../CustomPages/Client/ClientFuncs.js'></script>");
            #endregion
        }
        public override void BuildContents()
        {
            try
            {
                #region Get publisher Address
                string sAddressSQL = " select * from vPublicationAddress where pblc_PublicationsID=" + sPublicationId + " and adli_Addressid=" + sAddressId + " ";
                QuerySelect sQueryObj = GetQuery();
                sQueryObj.SQLCommand = sAddressSQL;
                sQueryObj.ExecuteReader();
                #endregion
                if (!sQueryObj.Eof())
                {
                    if (!string.IsNullOrEmpty(sQueryObj.FieldValue("addr_address1")))
                        sAddress1 = sQueryObj.FieldValue("addr_address1");
                    else
                        sAddress1 = "";

                    if (!string.IsNullOrEmpty(sQueryObj.FieldValue("addr_address2")))
                        sAddress2 = sQueryObj.FieldValue("addr_address2");
                    else
                        sAddress2 = "";

                    if (!string.IsNullOrEmpty(sQueryObj.FieldValue("addr_address3")))
                        sAddress3 = sQueryObj.FieldValue("addr_address3");
                    else
                        sAddress3 = "";

                    if (!string.IsNullOrEmpty(sQueryObj.FieldValue("addr_address4")))
                        sAddress4 = sQueryObj.FieldValue("addr_address4");
                    else
                        sAddress4 = "";

                    if (!string.IsNullOrEmpty(sQueryObj.FieldValue("addr_city")))
                        sCity = sQueryObj.FieldValue("addr_city");
                    else
                        sCity = "";

                    if (!string.IsNullOrEmpty(sQueryObj.FieldValue("addr_state")))
                        sState = sQueryObj.FieldValue("addr_state");
                    else
                        sState = "";

                    if (!string.IsNullOrEmpty(sQueryObj.FieldValue("addr_postcode")))
                        sZipCode = sQueryObj.FieldValue("addr_postcode");
                    else
                        sZipCode = "";

                    if (!string.IsNullOrEmpty(sQueryObj.FieldValue("addr_country")))
                        sCountry = sQueryObj.FieldValue("addr_country");
                    else
                        sCountry = "";

                    if (!string.IsNullOrEmpty(sQueryObj.FieldValue("Type")))
                        sType = sQueryObj.FieldValue("Type");
                    else
                        sType = "";

                    if (!string.IsNullOrEmpty(sQueryObj.FieldValue("pblc_primarypublicationid")))
                        sDefaultAddress = sQueryObj.FieldValue("pblc_primarypublicationid");
                    else
                        sDefaultAddress = "";

                    sHTML += HTML.Form();

                    sHTML += HTML.StartTable();

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData("Address 1:", "VIEWBOXCAPTION");
                    sHTML += HTML.TableData("Address 2:", "VIEWBOXCAPTION");
                    sHTML += HTML.TableData("") + HTML.TableData("Type", "VIEWBOXCAPTION");

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData(sAddress1, "VIEWBOX");
                    sHTML += HTML.TableData(sAddress2, "VIEWBOX");

                    if (sType.ToLower().Trim().Contains("business"))
                        sHTML += HTML.TableData("Business", "VIEWBOXCAPTION") + HTML.TableData(HTML.InputCheckBox("addr_business", true), "VIEWBOXCAPTION");

                    else
                        sHTML += HTML.TableData("Business", "VIEWBOXCAPTION") + HTML.TableData(HTML.InputCheckBox("addr_business", false), "VIEWBOXCAPTION");

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData("Address 3:", "VIEWBOXCAPTION");
                    sHTML += HTML.TableData("Address 4:", "VIEWBOXCAPTION");

                    if (sType.ToLower().Trim().Contains("billing"))
                        sHTML += HTML.TableData("Billing", "VIEWBOXCAPTION") + HTML.TableData(HTML.InputCheckBox("addr_billing", true), "VIEWBOXCAPTION");

                    else
                        sHTML += HTML.TableData("Billing", "VIEWBOXCAPTION") + HTML.TableData(HTML.InputCheckBox("addr_billing", false), "VIEWBOXCAPTION");

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData(sAddress3, "VIEWBOX");
                    sHTML += HTML.TableData(sAddress4, "VIEWBOX");
                    if (sType.ToLower().Trim().Contains("shipping"))
                        sHTML += HTML.TableData("Shipping", "VIEWBOXCAPTION") + HTML.TableData(HTML.InputCheckBox("addr_shipping", true), "VIEWBOXCAPTION");

                    else
                        sHTML += HTML.TableData("Shipping", "VIEWBOXCAPTION") + HTML.TableData(HTML.InputCheckBox("addr_shipping", false), "VIEWBOXCAPTION");

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData("City:", "VIEWBOXCAPTION");
                    sHTML += HTML.TableData("State:", "VIEWBOXCAPTION");

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData(sCity, "VIEWBOX");
                    sHTML += HTML.TableData(sState, "VIEWBOX");

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData("Zip Code:", "VIEWBOXCAPTION");
                    sHTML += HTML.TableData("Country:", "VIEWBOXCAPTION");

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData(sZipCode, "VIEWBOX");

                    string sHTMLCountry = "";
                    sHTMLCountry += "";
                    if (sCountry != "")
                    {
                        sHTMLCountry += (Metadata.GetTranslation("addr_country", sCountry));
                    }

                    sHTML += HTML.TableData(sHTMLCountry, "VIEWBOX");
                    sHTML += HTML.TableRow("");

                    if (sDefaultAddress != "0" && sDefaultAddress != "")
                    {
                        if (sDefaultAddress == sAddressId)
                            sHTML += HTML.TableData("Set as default address for Publication  " + HTML.InputCheckBox("addr_default", true), "VIEWBOXCAPTION");
                        else
                            sHTML += HTML.TableData("Set as default address for Publication  " + HTML.InputCheckBox("addr_default", false), "VIEWBOXCAPTION");
                    }
                    else if (sDefaultAddress == "")
                    {
                        sHTML += HTML.TableData("Set as default address for Publication  " + HTML.InputCheckBox("addr_default", false), "VIEWBOXCAPTION");
                    }

                    sHTML += HTML.EndTable();

                    AddContent(HTML.Box("Address", sHTML));
                    if (!string.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_Delete")))
                        sHiddenDelete = Dispatch.ContentField("HIDDEN_Delete");
                    else
                        sHiddenDelete = "";
                    if (sHiddenDelete == "Delete")
                    {
                        DeleteAddress(sHiddenDelete, "");
                    }
                    else
                    {
                        #region Define the Hidden Fields
                        AddContent(HTML.InputHidden("HIDDEN_Delete", ""));
                        #endregion
                        #region Add Buttons
                        AddUrlButton("Confirm Delete", "delete.gif", "javascript:SetDeleteParam();");
                        AddUrlButton("Cancel", "Cancel.gif", UrlDotNet(this.ThisDotNetDll, "RunPublicationAddressList"));
                        #endregion
                    }
                }
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
        }
        public void DeleteAddress(string HiddenParam, string TypeParam)
        {
            if (HiddenParam != "")
            {
                Record objAddress = FindRecord("Address", "Addr_addressid=" + sAddressId);
                if (!objAddress.Eof())
                {
                    objAddress.SetField("addr_deleted", 1);
                    objAddress.SaveChanges();
                }

                Record objAddresslink = FindRecord("Address_link", "adli_addressid=" + sAddressId);
                if (!objAddresslink.Eof())
                {
                    objAddresslink.SetField("adli_deleted", 1);
                    objAddresslink.SaveChanges();
                }


                string sURL = UrlDotNet(this.ThisDotNetDll, "RunPublicationAddressList");
                Dispatch.Redirect(sURL);
            }
            else
            {
                Record objAddresslink = FindRecord("Address_link", "adli_addressid=" + sAddressId + " and LOWER(LTRIM(RTRIM(adli_type)))='" + TypeParam.ToLower().Trim() + "'");

                if (!objAddresslink.Eof())
                {
                    objAddresslink.SetField("adli_deleted", 1);
                    objAddresslink.SaveChanges();
                }
            }

        }
    }
}
