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
    public class PublicationAddressEdit : Web
    {
        CRMHelper objCRMHelper = new CRMHelper();
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
        public PublicationAddressEdit()
        {
            GetTabs("publications");
            AddTopContent(GetCustomEntityTopFrame("publications"));
            #region Set Publication ID
            if (!String.IsNullOrEmpty(Dispatch.EitherField("pblc_PublicationsID")))
                sPublicationId = Dispatch.EitherField("pblc_PublicationsID");
            else
                sPublicationId = Dispatch.EitherField("Key58");
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
            #region Set Address Id
            if (!string.IsNullOrEmpty(Dispatch.EitherField("Addr_AddressId")))
                sAddressId = Dispatch.EitherField("Addr_AddressId");
            else if (!string.IsNullOrEmpty(Dispatch.ContentField("Addr_AddressId")))
                sAddressId = Dispatch.ContentField("Addr_AddressId");
            #endregion
            #region Add HTML Form so that Navigation will work as expected
            AddContent(HTML.Form());
            #endregion
            #region Define the Hidden Fields
            AddContent(HTML.InputHidden("HIDDEN_Save", ""));
            AddContent(HTML.InputHidden("HIDDEN_Delete", ""));
            #endregion
            #region Add Buttons
            AddUrlButton("Save", "save.gif", "javascript:SetAddressParam();");
            AddUrlButton("Delete", "delete.gif", "javascript:SetDeleteParam();");
            AddUrlButton("Cancel", "Cancel.gif", UrlDotNet(this.ThisDotNetDll, "RunPublicationAddressList"));
            #endregion

            #region Set js file reference path
            AddContent("<script type='text/javascript' src='../CustomPages/Client/ClientFuncs.js'></script>");
            #endregion
        }
        public override void BuildContents()
        {
            try
            {

                string sAddressSQL = " select * from vPublicationAddress where pblc_PublicationsID=" + sPublicationId + " and adli_Addressid=" + sAddressId + " ";
                QuerySelect sQueryObj = GetQuery();
                sQueryObj.SQLCommand = sAddressSQL;
                sQueryObj.ExecuteReader();

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
                    sHTML += HTML.TableData(HTML.InputText("addr_address1", sAddress1, 40, 20, "", "", false, "tabindex=1") + "<font style='color:blue;'>*</font>");
                    sHTML += HTML.TableData(HTML.InputText("addr_address2", sAddress2, 40, 40, "", "", false, "tabindex=2"));

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
                    sHTML += HTML.TableData(HTML.InputText("addr_address3", sAddress3, 40, 40, "", "", false, "tabindex=3"));
                    sHTML += HTML.TableData(HTML.InputText("addr_address4", sAddress4, 40, 40, "", "", false, "tabindex=4"));
                    if (sType.ToLower().Trim().Contains("shipping"))
                        sHTML += HTML.TableData("Shipping", "VIEWBOXCAPTION") + HTML.TableData(HTML.InputCheckBox("addr_shipping", true), "VIEWBOXCAPTION");

                    else
                        sHTML += HTML.TableData("Shipping", "VIEWBOXCAPTION") + HTML.TableData(HTML.InputCheckBox("addr_shipping", false), "VIEWBOXCAPTION");

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData("City:", "VIEWBOXCAPTION");
                    sHTML += HTML.TableData("State:", "VIEWBOXCAPTION");

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData(HTML.InputText("addr_city", sCity, 30, 20, "", "", false, "tabindex=5"));
                    sHTML += HTML.TableData(HTML.InputText("addr_state", sState, 30, 10, "", "", false, "tabindex=6"));

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData("Zip Code:", "VIEWBOXCAPTION");
                    sHTML += HTML.TableData("Country:", "VIEWBOXCAPTION");

                    sHTML += HTML.TableRow("");
                    sHTML += HTML.TableData(HTML.InputText("addr_postcode", sZipCode, 10, 15, "", "", false, "tabindex=7"));

                    string sSQL = "";
                    sSQL = " select LTRIM(RTRIM(cast(Capt_code as nvarchar))) as code, LTRIM(RTRIM(cast(Capt_US as nvarchar))) as Caption from custom_Captions where capt_Deleted is null and capt_family='addr_country'";
                    QuerySelect AddressObj = GetQuery();
                    AddressObj.SQLCommand = sSQL;
                    AddressObj.ExecuteReader();
                    string sHTMLCountry = "";
                    sHTMLCountry = "<style type=text/css> select {  font-family: Tahoma,Arial;font-size:11px; width:150px;color=#4d4f53 }</style>";
                    sHTMLCountry += "";
                    sHTMLCountry += "<select  name=addr_country id=addr_country tabindex=8> <option value=''>--None--</option>";
                    while (!AddressObj.Eof())
                    {
                        if (sCountry.ToLower() == AddressObj.FieldValue("code").ToLower())
                            sHTMLCountry += "<option value=" + AddressObj.FieldValue("code") + " SELECTED>" + AddressObj.FieldValue("Caption") + "</option>";
                        else
                            sHTMLCountry += "<option value=" + AddressObj.FieldValue("code") + ">" + AddressObj.FieldValue("Caption") + "</option>";
                        AddressObj.Next();
                    }

                    sHTMLCountry += "</select>";
                    sHTML += HTML.TableData(HTML.Span("addr_country", sHTMLCountry));
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
                }

                if (!string.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_Save")))
                    sHidden = Dispatch.ContentField("HIDDEN_Save");
                else
                    sHidden = "";

                if (!string.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_Delete")))
                    sHiddenDelete = Dispatch.ContentField("HIDDEN_Delete");
                else
                    sHiddenDelete = "";

                if (sHidden == "Save")
                {
                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_address1")))
                        sFormAddress1 = Dispatch.ContentField("addr_address1");
                    else
                        sFormAddress1 = "";

                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_address2")))
                        sFormAddress2 = Dispatch.ContentField("addr_address2");
                    else
                        sFormAddress2 = "";

                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_address3")))
                        sFormAddress3 = Dispatch.ContentField("addr_address3");
                    else
                        sFormAddress3 = "";

                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_address4")))
                        sFormAddress4 = Dispatch.ContentField("addr_address4");
                    else
                        sFormAddress4 = "";

                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_city")))
                        sFormCity = Dispatch.ContentField("addr_city");
                    else
                        sFormCity = "";

                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_state")))
                        sFormState = Dispatch.ContentField("addr_state");
                    else
                        sFormState = "";

                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_postcode")))
                        sFormZipCode = Dispatch.ContentField("addr_postcode");
                    else
                        sFormZipCode = "";

                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_country")))
                    {
                        sFormCountry = Dispatch.ContentField("addr_country");
                    }
                    else
                        sFormCountry = "";

                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_default")))
                        sFormDefaultAddress = Dispatch.ContentField("addr_default");
                    else
                        sFormDefaultAddress = "";

                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_business")))
                        sFormBusinessType = Dispatch.ContentField("addr_business");
                    else
                        sFormBusinessType = "";

                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_billing")))
                        sFormBillingType = Dispatch.ContentField("addr_billing");
                    else
                        sFormBillingType = "";

                    if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_shipping")))
                        sFormShippingType = Dispatch.ContentField("addr_shipping");
                    else
                        sFormShippingType = "";


                    Record AddressRec = FindRecord("address_link", " adli_addressid='" + sAddressId + "'");
                    if (!AddressRec.Eof())
                    {
                        string sBillingType = "";
                        string sBusinessType = "";
                        string sShippingType = "";
                        while (!AddressRec.Eof())
                        {
                            string sAddressType = AddressRec.GetFieldAsString("adli_type");

                            if (AddressRec.GetFieldAsString("adli_type").ToLower() == "billing")
                            {
                                sBillingType = "true";
                            }

                            if (AddressRec.GetFieldAsString("adli_type").ToLower() == "shipping")
                            {
                                sShippingType = "true";
                            }

                            if (AddressRec.GetFieldAsString("adli_type").ToLower() == "business")
                            {
                                sBusinessType = "true";
                            }
                            if (AddressRec.GetFieldAsString("adli_type").ToLower() == "billing" && sFormBillingType != "")
                            {
                                UpdateAddress();
                            }
                            else if (AddressRec.GetFieldAsString("adli_type").ToLower() == "shipping" && sFormShippingType != "")
                            {
                                UpdateAddress();
                            }
                            else if (AddressRec.GetFieldAsString("adli_type").ToLower() == "business" && sFormBusinessType != "")
                            {
                                UpdateAddress();
                            }
                            else if (AddressRec.GetFieldAsString("adli_type").ToLower() == "billing" && sFormBillingType == "")
                            {
                                DeleteAddress("", sAddressType.Trim());
                            }
                            else if (AddressRec.GetFieldAsString("adli_type").ToLower() == "shipping" && sFormShippingType == "")
                            {
                                DeleteAddress("", sAddressType.Trim());
                            }
                            else if (AddressRec.GetFieldAsString("adli_type").ToLower() == "business" && sFormBusinessType == "")
                            {
                                DeleteAddress("", sAddressType.Trim());
                            }

                            if ((sDefaultAddress == "0" || sDefaultAddress == "") && sFormDefaultAddress != "")
                            {
                                CreateDefaultAddress();
                            }
                            else if (sDefaultAddress != "0" && sFormDefaultAddress == "")
                            {
                                DeleteDefaultAddress();
                            }
                            if (sBillingType == "" && sFormBillingType != "")
                            {
                                CreateNewAddress("Billing");
                            }
                            if (sShippingType == "" && sFormShippingType != "")
                            {
                                CreateNewAddress("Shipping");
                            }
                            if (sBusinessType == "" && sFormBusinessType != "")
                            {
                                CreateNewAddress("Business");
                            }
                            if (sFormBillingType == "" && sFormShippingType == "" && sFormBusinessType == "")
                            {
                                UpdateAddress();
                            }
                            if (sDefaultAddress != sAddressId && sFormDefaultAddress != "")
                            {
                                CreateDefaultAddress();
                            }
                            AddressRec.GoToNext();

                        }

                    }
                    else
                    {
                        if (sFormBillingType != "")
                        {
                            CreateNewAddress("Billing");
                        }

                        if (sFormShippingType != "")
                        {
                            CreateNewAddress("Shipping");
                        }

                        if (sFormBusinessType != "")
                        {
                            CreateNewAddress("Business");
                        }

                    }
                    string sURL = UrlDotNet(this.ThisDotNetDll, "RunPublicationAddressList");
                    Dispatch.Redirect(sURL);
                }
                if (sHiddenDelete == "Delete")
                {
                    //DeleteAddress(sHiddenDelete, "");
                    string sURL = UrlDotNet(this.ThisDotNetDll, "RunPublicationConfirmDelete&pblc_PublicationsID=" + sPublicationId + "&Addr_AddressId=" + sAddressId);
                    Dispatch.Redirect(sURL);
                } 
            }
            catch (Exception error)
            {
                this.AddError(error.Message);
            }
        }
        public void UpdateAddress()
        {
            Record objAddressRec = FindRecord("address", "addr_addressid=" + sAddressId);
            if (!objAddressRec.Eof())
            {

                objAddressRec.SetField("Addr_Address1", sFormAddress1);
                objAddressRec.SetField("Addr_Address2", sFormAddress2);
                objAddressRec.SetField("Addr_Address3", sFormAddress3);
                objAddressRec.SetField("Addr_Address4", sFormAddress4);
                objAddressRec.SetField("Addr_City", sFormCity);
                objAddressRec.SetField("Addr_State", sFormState);
                objAddressRec.SetField("Addr_PostCode", sFormZipCode);
                objAddressRec.SetField("Addr_Country", sFormCountry);
                objAddressRec.SaveChanges();

                iAddressId = objAddressRec.RecordId;
            }
        }

        public void UpdateAddressLink(string sType)
        {
            Record objAddressLinkRec = FindRecord("address_link", "adli_addressid=" + sAddressId + " and adli_publicationid=" + sPublicationId + " and LTRIM(RTRIM(adli_type))='" + sType.Trim() + "' ");
            if (!objAddressLinkRec.Eof())
            {
                objAddressLinkRec.SetField("AdLi_CompanyID", sCompanyid);
                objAddressLinkRec.SetField("AdLi_PersonID", sPersonId);
                objAddressLinkRec.SetField("adli_publicationid", sPublicationId);
                objAddressLinkRec.SetField("AdLi_AddressId", iAddressId);
                objAddressLinkRec.SetField("adli_type", sType);
                objAddressLinkRec.SaveChanges();
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

        public void CreateNewAddress(string sFrom)
        {
            Record objAddressRec = new Record("Address");
            objAddressRec.SetField("Addr_Address1", sAddress1);
            objAddressRec.SetField("Addr_Address2", sAddress2);
            objAddressRec.SetField("Addr_Address3", sAddress3);
            objAddressRec.SetField("Addr_Address4", sAddress4);
            objAddressRec.SetField("Addr_City", sCity);
            objAddressRec.SetField("Addr_State", sState);
            objAddressRec.SetField("Addr_PostCode", sZipCode);
            objAddressRec.SetField("Addr_Country", sCountry);
            objAddressRec.SaveChanges();

            iAddressId = objAddressRec.RecordId;
            CreateNewAddressLink(iAddressId.ToString(), sFrom);

        }

        public void CreateNewAddressLink(string AddressId, string sFrom)
        {
            if (String.IsNullOrEmpty(sCompanyid))
                sCompanyid = "";
            if (String.IsNullOrEmpty(sPersonId))
                sPersonId = "";
            Record objAddressLinkRec = new Record("Address_link");
            objAddressLinkRec.SetField("AdLi_CompanyID", sCompanyid);
            objAddressLinkRec.SetField("AdLi_PersonID", sPersonId);
            objAddressLinkRec.SetField("adli_publicationid", sPublicationId);
            objAddressLinkRec.SetField("AdLi_AddressId", sAddressId);
            objAddressLinkRec.SetField("adli_type", sFrom);
            objAddressLinkRec.SaveChanges();
        }
        public void CreateDefaultAddress()
        {
            Record objBuildingRec = FindRecord("publications", "pblc_PublicationsID=" + sPublicationId + " and pblc_Deleted is null");
            if (!objBuildingRec.Eof())
            {
                objBuildingRec.SetField("pblc_primarypublicationid", sAddressId);
                objBuildingRec.SaveChanges();
            }
        }
        public void DeleteDefaultAddress()
        {
            Record objBuildingRec = FindRecord("publications", "pblc_PublicationsID=" + sPublicationId + " and pblc_Deleted is null");
            if (!objBuildingRec.Eof())
            {
                objBuildingRec.SetField("pblc_primarypublicationid", "");
                objBuildingRec.SaveChanges();
            }
        }
    }
}
