using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;

namespace NZPACRM
{
    public class PublisherAddressNew : Web
    {
        CRMHelper objCRMHelper = new CRMHelper();
        #region Variable Declaration
        string sPublishersid = "";
        string sCompanyid = "";
        string sPersonId = "";       
        string sCountry = "";
        string sAddress1 = "";
        string sAddress2 = "";
        string sAddress3 = "";
        string sAddress4 = "";
        string sCity = "";
        string sState = "";
        string sZipCode = "";
        string sDefaultAddress = "";
        string sBusinessType = "";
        string sBillingType = "";
        string sShippingType = "";
        int AddressId = 0;
        
        #endregion
        public PublisherAddressNew()
        {
            //'Set Tabs
            objCRMHelper.SetTabs("Publishers");

            //'Set Context of Custom Entity
            objCRMHelper.SetCustomEntityTopFrame("Publishers");
            #region Set js file reference path
            AddContent("<script type='text/javascript' src='../CustomPages/Client/ClientFuncs.js'></script>");
            #endregion
        }
        public override void BuildContents()
        {
            try
            {
                #region Adding Html Form
                AddContent(HTML.Form());
                #endregion
                #region Get Address
                objCRMHelper.AddressBox("Publishers");
                #endregion
                #region Get Publication id
                if (!String.IsNullOrEmpty(Dispatch.EitherField("pbls_PublishersID")))
                    sPublishersid = Dispatch.EitherField("pbls_PublishersID");
                else
                    sPublishersid = Dispatch.EitherField("Key58");
                #endregion
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

                #region To Save Address Record
                if (!String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_Save")))
                {
                    if (Dispatch.ContentField("HIDDEN_Save") == "Save")
                    {
                        if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_address1")))
                            sAddress1 = Dispatch.ContentField("addr_address1");
                        else
                            sAddress1 = "";

                        if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_address2")))
                            sAddress2 = Dispatch.ContentField("addr_address2");
                        else
                            sAddress2 = "";

                        if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_address3")))
                            sAddress3 = Dispatch.ContentField("addr_address3");
                        else
                            sAddress3 = "";

                        if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_address4")))
                            sAddress4 = Dispatch.ContentField("addr_address4");
                        else
                            sAddress4 = "";

                        if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_city")))
                            sCity = Dispatch.ContentField("addr_city");
                        else
                            sCity = "";

                        if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_state")))
                            sState = Dispatch.ContentField("addr_state");
                        else
                            sState = "";

                        if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_postcode")))
                            sZipCode = Dispatch.ContentField("addr_postcode");
                        else
                            sZipCode = "";

                        if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_country")))
                        {
                            sCountry = Dispatch.ContentField("addr_country");
                        }
                        else
                            sCountry = "";

                        if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_default")))
                            sDefaultAddress = Dispatch.ContentField("addr_default");
                        else
                            sDefaultAddress = "";

                        if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_business")))
                            sBusinessType = Dispatch.ContentField("addr_business");
                        else
                            sBusinessType = "";

                        if (!string.IsNullOrEmpty(Dispatch.ContentField("add_billing")))
                            sBillingType = Dispatch.ContentField("add_billing");
                        else
                            sBillingType = "";
                        if (!string.IsNullOrEmpty(Dispatch.ContentField("addr_shipping")))
                            sShippingType = Dispatch.ContentField("addr_shipping");
                        else
                            sShippingType = "";
                        if (sBusinessType == "" || sBillingType == "" || sShippingType == "")
                        {

                            #region Method To Create New Address Record
                            AddressId = CreaNewRecord(sAddress1, sAddress2, sAddress3, sAddress4, sCity, sState, sZipCode, sCountry);
                            CreateNewAddressLink(sCompanyid, sPersonId, AddressId, "");
                            #endregion
                        }
                        else
                        {
                            #region Method To Create New Address Record
                            AddressId = CreaNewRecord(sAddress1, sAddress2, sAddress3, sAddress4, sCity, sState, sZipCode, sCountry);
                            CreateNewAddressLink(sCompanyid, sPersonId, AddressId, "");
                            #endregion
                        }
                        #region Method to create address link
                        string sFrom = "";
                        if ((sBusinessType != "" && sBillingType != "") || (sBusinessType != "" && sShippingType != "") || (sBillingType != "" && sShippingType != ""))
                        {
                            sFrom = "None";
                        }                           
                        if (sBusinessType != "")
                        {
                            sFrom = "Business";
                            CreateNewAddressLink(sCompanyid, sPersonId, AddressId, sFrom);
                        }
                        if (sBillingType != "")
                        {
                            sFrom = "Billing";
                            CreateNewAddressLink(sCompanyid, sPersonId, AddressId, sFrom);
                        }
                        if (sShippingType != "")
                        {
                            sFrom = "Shipping";
                            CreateNewAddressLink(sCompanyid, sPersonId, AddressId, sFrom);
                        }
                        #endregion
                        string sURL = UrlDotNet(this.ThisDotNetDll, "RunPublisherAddressList");                        
                        Dispatch.Redirect(sURL);
                        
                    }
                }
                else
                {
                    #region Define the Hidden Fields
                    AddContent(HTML.InputHidden("HIDDEN_Save", ""));
                    #endregion

                    #region Add Buttons
                    //string sUrl = "javascript:document.EntryForm.HIDDEN_Save.value='Save';";
                    //AddSubmitButton("Save", "Save.gif", sUrl);
                    AddUrlButton("Save", "save.gif", "javascript:SetAddressParam();");
                    AddUrlButton("Cancel", "Cancel.gif", UrlDotNet(this.ThisDotNetDll, "RunPublisherAddressList"));     
                    #endregion
                }
                #endregion
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
        }
        public void CreateNewAddressLink(string sCompanyid, string sPersonId, int AddressId, string sFrom)
        {
            //AddContent("AddressIdLink=" + AddressId + "sCompanyid=" + sCompanyid + "sPersonId=" + sPersonId + "AddressId=" + AddressId.ToString() +"sFrom ="+ sFrom + "sPublicationId="+ sPublicationId);
            if (sFrom == "None" || String.IsNullOrEmpty(sFrom))
                sFrom = "";
            if (String.IsNullOrEmpty(sCompanyid))
                sCompanyid = "";
            if (String.IsNullOrEmpty(sPersonId))
                sPersonId = "";
            string objEntity = "Address_Link";
            string[] objParaName = new string[] { "AdLi_CompanyID", "AdLi_PersonID", "AdLi_AddressId", "adli_type", "adli_publisherid" };
            string[] objParaValue = new string[] { sCompanyid, sPersonId, AddressId.ToString(), sFrom, sPublishersid };
            objCRMHelper.CreateNewAddress(objEntity, objParaName, objParaValue);
            if (sDefaultAddress != "")
            {
                Record objClientRec = FindRecord("publishers", "pbls_PublishersID=" + sPublishersid);
                if (!objClientRec.Eof())
                {
                    objClientRec.SetField("pbls_primarypublisherid", AddressId.ToString());
                    objClientRec.SaveChanges();
                }
            }
        }

        public int CreaNewRecord(string sAddress1, string sAddress2, string sAddress3, string sAddress4, string sCity, string sState, string sZipCode, string sCountry)
        {
            string objEntity = "Address";
            string[] objParaName = new string[] { "Addr_Address1", "Addr_Address2", "Addr_Address3", "Addr_Address4", "Addr_City", "Addr_State", "Addr_PostCode", "Addr_Country" };
            string[] objParaValue = new string[] { sAddress1, sAddress2, sAddress3, sAddress4, sCity, sState, sZipCode, sCountry };
            AddressId = objCRMHelper.CreateNewAddress(objEntity, objParaName, objParaValue);            
            return AddressId;
        }

    }
}
