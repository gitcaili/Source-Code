﻿using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using System.Data;
using NZPACRM.Common;
using System.IO;

namespace NZPACRM
{
    public class ImportClientPage : DataPageNew
    {
        CRMHelper objCRM = new CRMHelper();
        string LogfileName = "";
        public ImportClientPage()
            : base("client", "client_ClientID", "ClientNewEntry")
        {

        }
        public override void BuildContents()
        {
            try
            {
                #region Adding Html Form
                AddContent(HTML.Form());
                #endregion
                #region Get Client template block
                objCRM.GetTemplateBlock("Client");
                #endregion

                string SavedFilePath = string.Empty;
                string sFileURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/";
                AddContent("<BR><BR>" + HTML.Box("File", "<br>&nbsp;&nbsp;<input type='file' id='fileupload' name='pic' size='70'>&nbsp;<input type='BUTTON' class='Edit'value='Import'name='upload' onclick='javascipt:CheckFile();'></br></br>"));
                if (!String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_Save")))
                {
                    if (Dispatch.ContentField("HIDDEN_Save") == "Save")
                    {
                        DataSet ds = new DataSet();
                        DataTable dt = new DataTable();
                        if (!String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_FileName")))
                        {
                            #region Save File in Client folder
                            //ds = SaveFileAtClientLocation();
                            SavedFilePath = SaveFileAtClientLocation();
                            #endregion

                            #region Read Saved Excel
                            string extention = Path.GetExtension(SavedFilePath);
                            dt = objCRM.ConvertToDataTable(SavedFilePath);
                            int ImportCount = 0;
                            string EntityID = "0";

                            if (dt.Rows.Count > 0)
                            {
                                if (dt.Rows[0][0].ToString().Trim().ToLower() == "client" || dt.Rows[0]["Type"].ToString().Trim().ToLower() == "client")
                                {
                                    for (int i = 0; i < dt.Rows.Count; i++)
                                    {
                                        int ClNewRecordID = 0;
                                     //   int UpClRecordID = 0;
                                        #region Find EnityID
                                        Record RecEnityID = FindRecord("Custom_Tables", "Bord_Caption = 'Client'");
                                        if (!RecEnityID.Eof())
                                        {
                                            EntityID = RecEnityID.GetFieldAsString("Bord_TableId");
                                        }
                                        #endregion
                                        string Name = "";
                                        string CurrUser = CurrentUser.UserId.ToString();
                                        string CompanyID = "";
                                        string PersonID = "";

                                        #region Insert Update Enity, Address, Phone, Email
                                        #region Entity type record
                                        #region Check for Entity Type

                                        if (dt.Rows[i][0].ToString().Trim().ToLower() == "client" || dt.Rows[i]["Type"].ToString().Trim().ToLower() == "client")
                                        {
                                            #region Create Person Attached with Entity
                                            string PersonFName = "";
                                            string PersonLName = "";
                                            string Website = "";

                                            if (!String.IsNullOrEmpty(dt.Rows[i]["Contact First Name"].ToString().Trim()))
                                                PersonFName = Convert.ToString(dt.Rows[i]["Contact First Name"]).Trim();

                                            if (!String.IsNullOrEmpty(dt.Rows[i]["Contact Last Name"].ToString().Trim()))
                                                PersonLName = Convert.ToString(dt.Rows[i]["Contact Last Name"]).Trim();

                                            string sWhereClause = "";

                                            if (PersonFName == "" && PersonLName == "")
                                                sWhereClause += "1=2";
                                            else
                                                sWhereClause += "isnull(pers_firstname,'') = '" + PersonFName + "' and isnull(pers_lastname,'') = '" + PersonLName + "' ";

                                            if (sWhereClause != "")
                                            {
                                                Record findPerson = FindRecord("person", sWhereClause);

                                                if (findPerson.RecordCount > 0)
                                                {
                                                    PersonID = findPerson.GetFieldAsString("Pers_PersonId");
                                                }
                                                else
                                                {
                                                    if (PersonFName != "" || PersonLName != "")
                                                    {
                                                        string[] ParamName = new string[] { "Pers_FirstName", "Pers_LastName", "Pers_CreatedBy", "Pers_CreatedDate" };
                                                        string[] ParamValue = new string[] { PersonFName, PersonLName, CurrentUser.UserId.ToString(), System.DateTime.Now.ToString() };
                                                        PersonID = Convert.ToString(objCRM.InsertUpdateEnity("person", ParamName, ParamValue));
                                                    }
                                                    else
                                                    {
                                                        Generatelog("[" + System.DateTime.Now.ToString() + "]" + " No Person is attached to Client");
                                                    }
                                                }
                                            }
                                            #endregion

                                            #region Insert Update Client

                                            if (dt.Rows[i]["Name"].ToString().Trim() != "")
                                            {
                                                Name = dt.Rows[i]["Name"].ToString().Trim();
                                                if (Name.Length > 30)
                                                {
                                                    Name = Name.Substring(0, 30);
                                                }
                                            }
                                            if (dt.Rows[i]["Website"].ToString().Trim() != "")
                                            {
                                                Website = dt.Rows[i]["Website"].ToString().Trim();
                                            }

                                            Record recFindClient = FindRecord("client", "client_Name='" + Name.Replace("'", "''") + "' and client_Deleted is null");

                                            if (recFindClient.RecordCount == 0)
                                            {
                                                #region Insert New Client record
                                                string[] ParaName = new string[] { "client_Name", "client_website", "client_CompanyId", "client_PersonId", "client_CreatedBy", "client_CreatedDate" };
                                                string[] ParaValue = new string[] { Name, Website, CompanyID, PersonID, CurrUser, System.DateTime.Now.ToString() };
                                                ClNewRecordID = objCRM.InsertUpdateEnity("client", ParaName, ParaValue);
                                                #endregion
                                            }
                                            else
                                            {
                                                #region Update Client record

                                                ClNewRecordID = Convert.ToInt32(recFindClient.GetFieldAsString("client_ClientID"));
                                                if (!recFindClient.Eof())
                                                {
                                                    recFindClient.SetField("client_website", Website);
                                                    if (PersonID != "")
                                                    {
                                                        recFindClient.SetField("client_PersonId", PersonID);
                                                    }
                                                    recFindClient.SaveChanges();
                                                }
                                                #endregion
                                            }
                                            #endregion

                                        }
                                        #endregion

                                        #region Address record

                                        #region Address Variable Declaration
                                        //string AddressID = "";
                                        //string PostalAddressID = "";
                                        string Address1 = "";
                                        string Address2 = "";
                                        string Address3 = "";
                                        string Address4 = "";
                                        string PhysicalType = "";
                                        string PhysicalCity = "";
                                        string PhysicalZipCode = "";
                                        string PhysicalCountry = "";
                                        string PostalAddr1 = "";
                                        string PostalAddr2 = "";
                                        string PostalAddr3 = "";
                                        string PostalAddr4 = "";
                                        string PostalCity = "";
                                        string PostalZipCode = "";
                                        string PostalCountry = "";
                                        string PostalType = "";

                                        if (Convert.ToString(dt.Rows[i]["Physical Address"]).Trim() != "")
                                        {
                                            Address1 = Convert.ToString(dt.Rows[i]["Physical Address"]).Trim();
                                            if (Address1.Length > 60)
                                            {
                                                Address1 = Address1.Substring(0, 60);
                                            }
                                        }
                                        if (Convert.ToString(dt.Rows[i]["Physical Address2"]).Trim() != "")
                                        {
                                            Address2 = Convert.ToString(dt.Rows[i]["Physical Address2"]).Trim();
                                            if (Address2.Length > 60)
                                            {
                                                Address2 = Address2.Substring(0, 60);
                                            }
                                        }
                                        if (Convert.ToString(dt.Rows[i]["Physical Address3"]).Trim() != "")
                                        {
                                            Address3 = Convert.ToString(dt.Rows[i]["Physical Address3"]).Trim();
                                            if (Address3.Length > 60)
                                            {
                                                Address3 = Address3.Substring(0, 60);
                                            }
                                        }
                                        if (Convert.ToString(dt.Rows[i]["Physical Address4"]).Trim() != "")
                                        {
                                            Address4 = Convert.ToString(dt.Rows[i]["Physical Address4"]).Trim();
                                            if (Address4.Length > 60)
                                            {
                                                Address4 = Address4.Substring(0, 60);
                                            }
                                        }

                                        if (Convert.ToString(dt.Rows[i]["Physical Type"]).Trim() != "")
                                            PhysicalType = Convert.ToString(dt.Rows[i]["Physical Type"]).Trim();

                                        if (Convert.ToString(dt.Rows[i]["Postal Address"]).Trim() != "")
                                        {
                                            PostalAddr1 = Convert.ToString(dt.Rows[i]["Postal Address"]).Trim();
                                            if (PostalAddr1.Length > 60)
                                            {
                                                PostalAddr1 = PostalAddr1.Substring(0, 60);
                                            }
                                        }
                                        if (Convert.ToString(dt.Rows[i]["Postal Address Line 2"]).Trim() != "")
                                        {
                                            PostalAddr2 = Convert.ToString(dt.Rows[i]["Postal Address Line 2"]).Trim();
                                            if (PostalAddr2.Length > 60)
                                            {
                                                PostalAddr2 = PostalAddr2.Substring(0, 60);
                                            }
                                        }
                                        if (Convert.ToString(dt.Rows[i]["Postal Address Line 3"]).Trim() != "")
                                        {
                                            PostalAddr3 = Convert.ToString(dt.Rows[i]["Postal Address Line 3"]).Trim();
                                            if (PostalAddr3.Length > 60)
                                            {
                                                PostalAddr3 = PostalAddr3.Substring(0, 60);
                                            }
                                        }
                                        if (Convert.ToString(dt.Rows[i]["Postal Address Line 4"]).Trim() != "")
                                        {
                                            PostalAddr4 = Convert.ToString(dt.Rows[i]["Postal Address Line 4"]).Trim();
                                            if (PostalAddr4.Length > 60)
                                            {
                                                PostalAddr4 = PostalAddr4.Substring(0, 60);
                                            }
                                        }

                                        if (Convert.ToString(dt.Rows[i]["Postal Type"]).Trim() != "")
                                            PostalType = Convert.ToString(dt.Rows[i]["Postal Type"]).Trim();

                                        if (Convert.ToString(dt.Rows[i]["Physical City"]).Trim() != "")
                                            PhysicalCity = Convert.ToString(dt.Rows[i]["Physical City"]).Trim();
                                        if (Convert.ToString(dt.Rows[i]["Physical Zip Code"]).Trim() != "")
                                            PhysicalZipCode = Convert.ToString(dt.Rows[i]["Physical Zip Code"]).Trim();
                                        if (Convert.ToString(dt.Rows[i]["Physical Country"]).Trim() != "")
                                            PhysicalCountry = Convert.ToString(dt.Rows[i]["Physical Country"]).Trim();
                                        string physcon = "";
                                        if (PhysicalCountry != "")
                                        {
                                            PhysicalCountry = Metadata.GetTranslation("addr_country", PhysicalCountry);
                                            physcon = Metadata.GetTranslation("addr_country", PhysicalCountry);
                                            if (PhysicalCountry != null || PhysicalCountry != "")
                                            {
                                                PhysicalCountry = PhysicalCountry.ToString();
                                                physcon = captioncode(PhysicalCountry);
                                            }
                                            else
                                            {
                                                PhysicalCountry = "";
                                            }
                                        }
                                        if (Convert.ToString(dt.Rows[i]["Postal City"]).Trim() != "")
                                            PostalCity = Convert.ToString(dt.Rows[i]["Postal City"]).Trim();
                                        if (Convert.ToString(dt.Rows[i]["Postal Zip Code"]).Trim() != "")
                                            PostalZipCode = Convert.ToString(dt.Rows[i]["Postal Zip Code"]).Trim();
                                        if (Convert.ToString(dt.Rows[i]["Postal Country"]).Trim() != "")
                                            PostalCountry = Convert.ToString(dt.Rows[i]["Postal Country"]).Trim();
                                        string postalcon = "";
                                        if (PostalCountry != "")
                                        {
                                            PostalCountry = Metadata.GetTranslation("addr_country", PostalCountry);

                                            if (PostalCountry != null || PostalCountry != "")
                                            {
                                                PostalCountry = PostalCountry.ToString();
                                                postalcon = captioncode(PostalCountry);
                                            }
                                            else
                                            {
                                                PostalCountry = "";
                                            }
                                        }
                                        #endregion

                                        if (ClNewRecordID != 0)
                                        {
                                            #region Physical adress
                                            // insert physical address
                                            if (Address1 != "" || Address2 != "" || Address3 != "" || Address4 != "" || PhysicalCity != "" || PhysicalZipCode != "" || PhysicalCountry != "")
                                            {

                                                Record objPhysicalAddressRec = FindRecord("Address", " Addr_address1= '" + Address1 + "' and addr_type ='Physical' and  Addr_AddressId in (select adli_addressid from Address_Link where AdLi_Type is null and adli_ClientId ='" + ClNewRecordID + "' and adli_deleted is null)");
                                                if (!objPhysicalAddressRec.Eof())
                                                {
                                                    //update address
                                                    Record objUpdateAddressRec = FindRecord("Address", "addr_address1='" + Address1 + "' and addr_type ='Physical' and  Addr_AddressId in (select adli_addressid from Address_Link where AdLi_Type is null and adli_ClientId ='" + ClNewRecordID + "' and adli_deleted is null)");
                                                    if (!objUpdateAddressRec.Eof())
                                                    {
                                                        objUpdateAddressRec.SetField("Addr_Address1", Address1);
                                                        objUpdateAddressRec.SetField("Addr_Address2", Address2);
                                                        objUpdateAddressRec.SetField("Addr_Address3", Address3);
                                                        objUpdateAddressRec.SetField("Addr_Address4", Address4);
                                                        objUpdateAddressRec.SetField("Addr_City", PhysicalCity);
                                                        objUpdateAddressRec.SetField("Addr_Country", PhysicalCountry);
                                                        objUpdateAddressRec.SetField("Addr_PostCode", PhysicalZipCode);
                                                        objUpdateAddressRec.SetField("addr_type", "Physical");
                                                        objUpdateAddressRec.SaveChanges();
                                                    }
                                                    #region Update address Link
                                                    Record recAddressLink = FindRecord("address_link", "adli_addressid = " + objUpdateAddressRec.RecordId + "and adli_Type is not null and AdLi_Deleted is null");
                                                    if (PhysicalType == "Shipping" || PhysicalType == "Business" || PhysicalType == "Billing")
                                                    {
                                                        if (!recAddressLink.Eof())
                                                        {
                                                            recAddressLink.SetField("adli_type", PhysicalType);
                                                            recAddressLink.SaveChanges();
                                                        }
                                                        else
                                                        {
                                                            Record objaddresslink = new Record("address_link");
                                                            objaddresslink.SetField("AdLi_AddressId", objUpdateAddressRec.RecordId);
                                                            objaddresslink.SetField("adli_ClientId", ClNewRecordID);
                                                            objaddresslink.SetField("AdLi_PersonID", PersonID);
                                                            objaddresslink.SetField("AdLi_Type", PhysicalType);
                                                            objaddresslink.SaveChanges();
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (PhysicalType == "")
                                                        {
                                                            if (!recAddressLink.Eof())
                                                            {
                                                                recAddressLink.SetField("adli_type", "");
                                                                recAddressLink.SaveChanges();
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                {
                                                    // insert address Physical
                                                    #region Insert New Physical address

                                                    Record objaddress = new Record("address");
                                                    objaddress.SetField("Addr_Address1", Address1);
                                                    objaddress.SetField("Addr_Address2", Address2);
                                                    objaddress.SetField("Addr_Address3", Address3);
                                                    objaddress.SetField("Addr_Address4", Address4);
                                                    objaddress.SetField("Addr_City", PhysicalCity);
                                                    objaddress.SetField("Addr_Country", PhysicalCountry);
                                                    objaddress.SetField("Addr_PostCode", PhysicalZipCode);
                                                    objaddress.SetField("addr_type", "Physical");

                                                    objaddress.SaveChanges();

                                                    #endregion

                                                    Record objClientRec = FindRecord("Client", "client_ClientID=" + ClNewRecordID);
                                                    if (!objClientRec.Eof())
                                                    {
                                                        objClientRec.SetField("client_primaryclientid", objaddress.RecordId);
                                                        objClientRec.SaveChanges();
                                                    }

                                                    #region Instert New Postal Address Link
                                                    Record objaddresslink1 = new Record("address_link");
                                                    objaddresslink1.SetField("AdLi_AddressId", objaddress.RecordId);
                                                    objaddresslink1.SetField("adli_ClientId", ClNewRecordID);
                                                    objaddresslink1.SetField("AdLi_PersonID", PersonID);
                                                    objaddresslink1.SetField("AdLi_Type", "");
                                                    objaddresslink1.SaveChanges();
                                                    // insert address

                                                    #endregion

                                                    if (PhysicalType == "Shipping" || PhysicalType == "Business" || PhysicalType == "Billing")
                                                    {
                                                        Record objaddresslink = new Record("address_link");
                                                        objaddresslink.SetField("AdLi_AddressId", objaddress.RecordId);
                                                        objaddresslink.SetField("adli_ClientId", ClNewRecordID);
                                                        objaddresslink.SetField("AdLi_PersonID", PersonID);
                                                        objaddresslink.SetField("AdLi_Type", PhysicalType);
                                                        objaddresslink.SaveChanges();
                                                    }

                                                }

                                            }
                                            #endregion

                                            #region Postal Address

                                            // insert Postal address
                                            if (PostalAddr1 != "" || PostalAddr2 != "" || PostalAddr3 != "" || PostalAddr4 != "" || PostalCity != "" || PostalCountry != "" || PostalZipCode != "")
                                            {

                                                Record objPhysicalAddressRec = FindRecord("Address", " Addr_address1= '" + PostalAddr1 + "' and addr_type ='Postal' and  Addr_AddressId in (select adli_addressid from Address_Link where AdLi_Type is null and adli_ClientId ='" + ClNewRecordID + "' and adli_deleted is null)");
                                                if (!objPhysicalAddressRec.Eof())
                                                {
                                                    //update address

                                                    Record objUpdateAddressRec = FindRecord("Address", "addr_address1='" + PostalAddr1 + "' and addr_type ='Postal' and  Addr_AddressId in (select adli_addressid from Address_Link where AdLi_Type is null and adli_ClientId ='" + ClNewRecordID + "' and adli_deleted is null)");
                                                    if (!objUpdateAddressRec.Eof())
                                                    {
                                                        objUpdateAddressRec.SetField("Addr_Address1", PostalAddr1);
                                                        objUpdateAddressRec.SetField("Addr_Address2", PostalAddr2);
                                                        objUpdateAddressRec.SetField("Addr_Address3", PostalAddr3);
                                                        objUpdateAddressRec.SetField("Addr_Address4", PostalAddr4);
                                                        objUpdateAddressRec.SetField("Addr_City", PostalCity);
                                                        objUpdateAddressRec.SetField("Addr_Country", PostalCountry);
                                                        objUpdateAddressRec.SetField("Addr_PostCode", PostalZipCode);
                                                        objUpdateAddressRec.SetField("addr_type", "Postal");
                                                        objUpdateAddressRec.SaveChanges();
                                                    }

                                                    #region Update Postal Address Link
                                                    Record recAddressLink = FindRecord("address_link", "adli_addressid = " + objUpdateAddressRec.RecordId + "and adli_Type is not null and AdLi_Deleted is null");
                                                    if (PostalType == "Shipping" || PostalType == "Business" || PostalType == "Billing")
                                                    {
                                                        if (!recAddressLink.Eof())
                                                        {
                                                            recAddressLink.SetField("adli_type", PostalType);
                                                            recAddressLink.SaveChanges();
                                                        }
                                                        else
                                                        {
                                                            Record objaddresslink = new Record("address_link");
                                                            objaddresslink.SetField("AdLi_AddressId", objUpdateAddressRec.RecordId);
                                                            objaddresslink.SetField("adli_ClientId", ClNewRecordID);
                                                            objaddresslink.SetField("AdLi_PersonID", PersonID);
                                                            objaddresslink.SetField("AdLi_Type", PostalType);
                                                            objaddresslink.SaveChanges();
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (PostalType == "")
                                                        {
                                                            if (!recAddressLink.Eof())
                                                            {
                                                                recAddressLink.SetField("adli_type", "");
                                                                recAddressLink.SaveChanges();
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                {
                                                    #region Insert New Postal address

                                                    Record objaddress = new Record("address");
                                                    objaddress.SetField("Addr_Address1", PostalAddr1);
                                                    objaddress.SetField("Addr_Address2", PostalAddr2);
                                                    objaddress.SetField("Addr_Address3", PostalAddr3);
                                                    objaddress.SetField("Addr_Address4", PostalAddr4);
                                                    objaddress.SetField("Addr_City", PostalCity);
                                                    objaddress.SetField("Addr_Country", PostalCountry);
                                                    objaddress.SetField("Addr_PostCode", PostalZipCode);
                                                    objaddress.SetField("addr_type", "Postal");

                                                    objaddress.SaveChanges();

                                                    #endregion

                                                    Record objClientRec = FindRecord("Client", "client_ClientID=" + ClNewRecordID);
                                                    if (!objClientRec.Eof())
                                                    {
                                                        objClientRec.SetField("client_primaryclientid", objaddress.RecordId);
                                                        objClientRec.SaveChanges();
                                                    }

                                                    #region Instert New Postal Address Link
                                                    Record objaddresslink1 = new Record("address_link");
                                                    objaddresslink1.SetField("AdLi_AddressId", objaddress.RecordId);
                                                    objaddresslink1.SetField("adli_ClientId", ClNewRecordID);
                                                    objaddresslink1.SetField("AdLi_PersonID", PersonID);
                                                    objaddresslink1.SetField("AdLi_Type", "");
                                                    objaddresslink1.SaveChanges();
                                                    // insert address

                                                    #endregion

                                                    if (PostalType == "Shipping" || PostalType == "Business" || PostalType == "Billing")
                                                    {
                                                        Record objaddresslink = new Record("address_link");
                                                        objaddresslink.SetField("AdLi_AddressId", objaddress.RecordId);
                                                        objaddresslink.SetField("adli_ClientId", ClNewRecordID);
                                                        objaddresslink.SetField("AdLi_PersonID", PersonID);
                                                        objaddresslink.SetField("AdLi_Type", PostalType);
                                                        objaddresslink.SaveChanges();
                                                    }
                                                }
                                            }

                                            #endregion
                                        }

                                        #endregion

                                        #region Phone record
                                        int PhoneID = 0;
                                        string PhoneNo = "";
                                        if (Convert.ToString(dt.Rows[i]["Phone"]).Trim() != "")
                                            PhoneNo = Convert.ToString(dt.Rows[i]["Phone"]).Trim();
                                        #region Insert New Phone
                                        //if (ClNewRecordID != 0)
                                        //{                                           
                                        //    if (PhoneNo != "")
                                        //    {
                                        //        
                                        //        //string[] ParamName = new string[] { "Phon_Number", "Phon_CreatedBy", "Phon_CreatedDate" };
                                        //        //string[] ParamValue = new string[] { PhoneNo, CurrUser, System.DateTime.Now.ToString() };
                                        //        //PhoneID = objCRM.InsertUpdateEnity("Phone", ParamName, ParamValue);
                                        //        //#endregion
                                        //        //#region Insert into Phone Link
                                        //        //string[] ParaName = new string[] { "PLink_PhoneId", "PLink_Type", "PLink_RecordID", "PLink_EntityID", "PLink_CreatedBy", "PLink_CreatedDate" };
                                        //        //string[] ParaValue = new string[] { PhoneID.ToString(), "Business", ClNewRecordID.ToString(), EntityID, CurrUser, System.DateTime.Now.ToString() };
                                        //        //objCRM.InsertUpdateEnity("PhoneLink", ParaName, ParaValue);
                                        //        
                                        //    }
                                        //}
                                        #endregion
                                        if (ClNewRecordID != 0)
                                        {
                                            string PhoneRecord = "select * from Phone inner join PhoneLink on Phone.Phon_PhoneId = PhoneLink.PLink_PhoneId where PLink_RecordID =" + ClNewRecordID + " and PLink_EntityID =" + EntityID + " and PLink_Deleted is null";
                                            QuerySelect sQueryObj = GetQuery();
                                            sQueryObj.SQLCommand = PhoneRecord;
                                            sQueryObj.ExecuteReader();
                                            if (!sQueryObj.Eof())
                                            {
                                                #region Update Phone
                                                PhoneID = Convert.ToInt32(sQueryObj.FieldValue("Phon_PhoneId"));
                                                Record RecPhone = FindRecord("phone", "Phon_PhoneId =" + PhoneID);
                                                if (!RecPhone.Eof())
                                                {
                                                    RecPhone.SetField("Phon_Number", PhoneNo);
                                                    RecPhone.SaveChanges();
                                                }
                                                #endregion
                                            }
                                            else
                                            {
                                                if (PhoneNo != "")
                                                {
                                                    #region Insert New Phone
                                                    string[] ParamName = new string[] { "Phon_Number", "Phon_CreatedBy", "Phon_CreatedDate" };
                                                    string[] ParamValue = new string[] { PhoneNo, CurrUser, System.DateTime.Now.ToString() };
                                                    PhoneID = objCRM.InsertUpdateEnity("Phone", ParamName, ParamValue);
                                                    #endregion
                                                    #region Insert into Phone Link
                                                    string[] ParaName = new string[] { "PLink_PhoneId", "PLink_Type", "PLink_RecordID", "PLink_EntityID", "PLink_CreatedBy", "PLink_CreatedDate" };
                                                    string[] ParaValue = new string[] { PhoneID.ToString(), "Business", ClNewRecordID.ToString(), EntityID, CurrUser, System.DateTime.Now.ToString() };
                                                    objCRM.InsertUpdateEnity("PhoneLink", ParaName, ParaValue);
                                                    #endregion
                                                }
                                            }
                                        }
                                        #endregion

                                        #region Email record
                                        int EmailID = 0;
                                        string Email = "";
                                        if (Convert.ToString(dt.Rows[i]["Email"]).Trim() != "")
                                            Email = Convert.ToString(dt.Rows[i]["Email"]).Trim();
                                        #region Insert New Email
                                        //if (ClNewRecordID != 0)
                                        //{
                                        //    if (Email != "")
                                        //    {

                                        //string[] paramName = new string[] { "Emai_EmailAddress", "Emai_CreatedBy", "Emai_CreatedDate" };
                                        //string[] paramValue = new string[] { Email, CurrUser, System.DateTime.Now.ToString() };
                                        //EmailID = objCRM.InsertUpdateEnity("Email", paramName, paramValue);
                                        //#endregion

                                        //#region Insert into EmailLink
                                        //string[] paraName = new string[] { "ELink_EmailId", "ELink_Type", "ELink_RecordID", "ELink_EntityID", "ELink_CreatedBy", "ELink_CreatedDate" };
                                        //string[] paraValue = new string[] { EmailID.ToString(), "Business", ClNewRecordID.ToString(), EntityID, CurrUser, System.DateTime.Now.ToString() };
                                        //objCRM.InsertUpdateEnity("EmailLink", paraName, paraValue);

                                        //    }
                                        //}
                                        #endregion

                                        if (ClNewRecordID != 0)
                                        {
                                            string EmailRecord = "select * from Email inner join EmailLink on Email.Emai_EmailId = EmailLink.ELink_EmailId where ELink_RecordID = " + ClNewRecordID + " and ELink_EntityID =" + EntityID + " and ELink_Deleted is null";
                                            QuerySelect sQueryObje = GetQuery();
                                            sQueryObje.SQLCommand = EmailRecord;
                                            sQueryObje.ExecuteReader();
                                            LogMessage(EmailRecord);
                                            if (!sQueryObje.Eof())
                                            {

                                                #region Update Email
                                                //if (Email != "")
                                                //{
                                                EmailID = Convert.ToInt32(sQueryObje.FieldValue("Emai_EmailId"));
                                                Record RecEmail = FindRecord("Email", "Emai_EmailId = " + EmailID);
                                                if (!RecEmail.Eof())
                                                {
                                                    RecEmail.SetField("Emai_EmailAddress", Email);
                                                    RecEmail.SaveChanges();
                                                }
                                                //}
                                                #endregion
                                            }
                                            else
                                            {
                                                if (Email != "")
                                                {
                                                    #region Insert New Email
                                                    string[] paramName = new string[] { "Emai_EmailAddress", "Emai_CreatedBy", "Emai_CreatedDate" };
                                                    string[] paramValue = new string[] { Email, CurrUser, System.DateTime.Now.ToString() };
                                                    EmailID = objCRM.InsertUpdateEnity("Email", paramName, paramValue);
                                                    #endregion

                                                    #region Insert into EmailLink
                                                    string[] paraName = new string[] { "ELink_EmailId", "ELink_Type", "ELink_RecordID", "ELink_EntityID", "ELink_CreatedBy", "ELink_CreatedDate" };
                                                    string[] paraValue = new string[] { EmailID.ToString(), "Business", ClNewRecordID.ToString(), EntityID, CurrUser, System.DateTime.Now.ToString() };
                                                    objCRM.InsertUpdateEnity("EmailLink", paraName, paraValue);
                                                    #endregion
                                                }
                                            }
                                        }
                                        #endregion

                                        #endregion
                                        #endregion
                                        ImportCount++;
                                        Generatelog("[" + System.DateTime.Now.ToString() + "]" + " Client " + Name + " is imported sucessfully in Sage CRM.");
                                    }

                                    string sURL = UrlDotNet(this.ThisDotNetDll, "RunClientImportCompletePage") + "&imported=" + ImportCount + "&LogFileName=" + LogfileName;
                                    Dispatch.Redirect(sURL);
                                }
                                else
                                {
                                    string sFailURL = UrlDotNet(this.ThisDotNetDll, "RunClientImportCompletePage") + "&Fail=Fail";
                                    Dispatch.Redirect(sFailURL);
                                }
                            }
                            #endregion
                        }

                    }
                }
                else
                {
                    #region Define the Hidden Fields
                    AddContent(HTML.InputHidden("HIDDEN_FilePath", ""));
                    AddContent(HTML.InputHidden("HIDDEN_Save", ""));
                    AddContent(HTML.InputHidden("HIDDEN_FileName", ""));
                    #endregion

                    #region Add Buttons
                    string backURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                    backURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=1650&Mode=1&CLk=T&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data Management";
                    AddUrlButton("Back", "prevcircle.gif", backURL);
                    #endregion

                }

            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
                LogMessage(ex.Message);
            }
        }

        #region MyRegion
        public string captioncode(string country)
        {
            string scountry = "";
            Record recCaption = FindRecord("Custom_Caption", "ltrim(rtrim(capt_US)) = '" + country + "'");
            if (!recCaption.Eof())
            {
                scountry = recCaption.GetFieldAsString("capt_code");
            }
            return scountry;
        }
        #endregion
        #region Get Library Path
        public string GetLibraryPath()
        {
            string Path = "";
            Record RecPath = FindRecord("Custom_SysParams", "parm_name = 'DocStore'");
            Path = RecPath.GetFieldAsString("Parm_Value");
            return Path;
        }
        #endregion
        #region Create Log file

        public void Generatelog(string Logcontent)
        {
            string LibPath = GetLibraryPath();
            string NewPath = LibPath.Replace("\\Library", "");
            NewPath += "WWWRoot\\CustomPages\\NZPAImport\\";
            string sInstallDirName = new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).DirectoryName;
            string Logspath = null;
            try
            {
                //    string currentPath = Directory.GetCurrentDirectory();
                string currentPath = NewPath;
                if (!Directory.Exists(Path.Combine(currentPath, "ImportLogs")))
                    Directory.CreateDirectory(Path.Combine(currentPath, "ImportLogs"));

                DateTime theDate = DateTime.Now;
                //string ymd = System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString() + System.DateTime.Now.Second.ToString() + theDate.ToString("yyyyMMdd") + "ClientLog.txt";
                string ymd = System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString() + theDate.ToString("yyyyMMdd") + "ClientLog.txt";
                LogfileName = ymd;
                Logspath = NewPath + "\\ImportLogs\\" + ymd;

                if (!File.Exists(Logspath))
                {
                    File.Create(Logspath).Close();
                    using (StreamWriter stream = new StreamWriter(Logspath, true))
                    {
                        stream.WriteLine(Logcontent);
                        stream.Flush();
                        stream.Close();
                    }
                }
                else
                {
                    using (StreamWriter stream = new StreamWriter(Logspath, true))
                    {
                        stream.WriteLine(Logcontent);
                        stream.Flush();
                        stream.Close();
                    }
                }
            }

            catch (Exception ex)
            {
                try
                {
                    using (StreamWriter stream = new StreamWriter(Logspath, true))
                    {
                        stream.WriteLine("error occurred =" + Logcontent + "Description= " + ex.Message);
                        stream.Flush();
                        stream.Close();
                    }
                }
                catch (Exception)
                {

                }
            }
        }


        #endregion

        #region Save File at Client folder
        public string SaveFileAtClientLocation()
        {
            string FileName = "";
            string newFullPath = "";
            string UploadfilePath = Dispatch.ContentField("HIDDEN_FilePath");
            string LibPath = GetLibraryPath();
            string NewPath = LibPath.Replace("\\Library", "");
            FileName = Dispatch.ContentField("HIDDEN_FileName");
            NewPath += "WWWRoot\\CustomPages\\NZPAImport\\ImportedFiles\\Clients\\";
            if (Directory.Exists(NewPath))
            {
                NewPath = NewPath + FileName;

                #region To check if file allready exists
                int count = 1;
                string fileNameOnly = Path.GetFileNameWithoutExtension(NewPath);
                string extension = Path.GetExtension(NewPath);
                string path = Path.GetDirectoryName(NewPath);
                if (!Path.IsPathRooted(NewPath))
                    NewPath = Path.GetFullPath(NewPath);
                string[] files = Directory.GetFiles(path);
                if (File.Exists(NewPath))
                {
                    count = files.Length;
                    newFullPath = Path.Combine(path, String.Format("{0} ({1}){2}", fileNameOnly, (count++), extension));

                    NewPath = newFullPath;
                    File.Copy(UploadfilePath, newFullPath);
                }
                else
                {
                    File.Copy(UploadfilePath, NewPath);
                }

            }
                #endregion
            return NewPath;
        }
        #endregion



    }
}
