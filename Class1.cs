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

namespace MAM14001.Building
{
    public class BuildingPhoneEmail : Web
    {
        public override void BuildContents()
        {
            AddContent(HTML.Form());
            try
            {
                /* Add your code here */

                #region Set Tab
                GetTabs("Building");
                #endregion

                AddTopContent(GetCustomEntityTopFrame("Building"));

                string sHTMLPhoneDetails = "";
                string sHTMLEmailDetails = "";
                string sHTMLDetails = "";
                string sUseCountryCode = "";
                string sUseAreaCode = "";
               // string sHTMLEmailGrid = "";
                string iEntityRecordID = "";
                //string sPhoneCategory = "";

                if (!String.IsNullOrEmpty(Dispatch.EitherField("buil_buildingid")))
                    iEntityRecordID = Dispatch.EitherField("buil_buildingid");
                else
                    iEntityRecordID = Dispatch.EitherField("Key58");

                #region Define the Hidden Fields
                AddContent(HTML.InputHidden("HIDDENPAGENUMBER_Save", ""));
                #endregion

                #region Set js file reference path
                AddContent("<script type='text/javascript' src='../CustomPages/Building/ClientFuncs.js'></script>");
                #endregion

                #region Check for Area Code and Country Code setting in CRM
                Record objCode = FindRecord("Custom_SysParams", " Parm_Name in ('UseAreaCode','UseCountryCode')");

                if (!objCode.Eof())
                {
                    while (!objCode.Eof())
                    {
                        if (objCode.GetFieldAsString("Parm_name") == "UseAreaCode")
                        {
                            sUseAreaCode = objCode.GetFieldAsString("Parm_value");
                        }
                        else if (objCode.GetFieldAsString("Parm_name") == "UseCountryCode")
                        {
                            sUseCountryCode = objCode.GetFieldAsString("Parm_value");
                        }
                        objCode.GoToNext();
                    }
                }
                #endregion

                #region Build Phone Structure
                sHTMLPhoneDetails = HTML.StartTable().ToString();
                sHTMLPhoneDetails += HTML.TableData("", "", "");
                sHTMLPhoneDetails += HTML.TableRow("", "").ToString();
                sHTMLPhoneDetails += HTML.TableData("");
                sHTMLPhoneDetails += HTML.TableData("Country", "VIEWBOXCAPTION");
                if (sUseAreaCode == "Y")
                    sHTMLPhoneDetails += HTML.TableData("Area", "VIEWBOXCAPTION");
                if (sUseCountryCode == "Y")
                    sHTMLPhoneDetails += HTML.TableData("Number", "VIEWBOXCAPTION");

                Record objPhoneField = FindRecord("Custom_captions", "  Capt_Family = N'Link_CompPhon' and capt_deleted is null ");

                if (!objPhoneField.Eof())
                {
                    while (!objPhoneField.Eof())
                    {
                        string sPhoneType = objPhoneField.GetFieldAsString("Capt_US");
                        if (!String.IsNullOrEmpty(sPhoneType))
                        {
                            sHTMLPhoneDetails += HTML.TableRow("", "");
                            sHTMLPhoneDetails += HTML.TableData("<b>" + objPhoneField.GetFieldAsString("Capt_US") + "</b>", "VIEWBOXCAPTION");

                            #region Phone County Code
                            if (sUseCountryCode == "Y")
                            {
                                #region Get the Phone Detail Record for Current Record
                                string sCountryCodeSQL = "select PLink_Type,Phon_CountryCode,Phon_AreaCode,Phon_Number  from PhoneLink inner join phone on plink_phoneid=phon_phoneid ";
                                sCountryCodeSQL += " where PLink_Deleted is null and Phon_Deleted is null  and plink_recordid=" + iEntityRecordID;
                                QuerySelect sQueryObj = GetQuery();
                                sQueryObj.SQLCommand = sCountryCodeSQL;
                                sQueryObj.ExecuteReader();
                                #endregion

                                #region Update Phone Country Code
                                if (!sQueryObj.Eof())
                                {
                                    while (!sQueryObj.Eof())
                                    {
                                        string sPhoneTypeSQL = "";
                                        string sPhoneCountryCodeSQL = "";

                                        if (!string.IsNullOrEmpty(sQueryObj.FieldValue("PLink_Type")))
                                            sPhoneTypeSQL = sQueryObj.FieldValue("PLink_Type");
                                        else
                                            sPhoneTypeSQL = "";

                                        if (!String.IsNullOrEmpty(sQueryObj.FieldValue("Phon_CountryCode")))
                                            sPhoneCountryCodeSQL = sQueryObj.FieldValue("Phon_CountryCode");
                                        else
                                            sPhoneCountryCodeSQL = "";

                                        if (sPhoneTypeSQL != "")
                                        {
                                            if (sPhoneTypeSQL.ToLower().TrimEnd(':') == objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'))
                                                sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_countrycode" + objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'), sPhoneCountryCodeSQL, 5, 3));
                                            else
                                                sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_countrycode" + objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'), "", 5, 3));
                                        }
                                        sQueryObj.Next();
                                    }
                                }
                                #endregion

                                #region Create New Phone Country Code
                                else
                                    sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_countrycode" + objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'), "", 5, 3));
                                #endregion

                            }
                            #endregion


                            #region Phone Area Code
                            if (sUseAreaCode == "Y")
                            {
                                #region Get the Phone Detail Record for Current Record
                                string sAreaCodeSQL = "select PLink_Type,Phon_CountryCode,Phon_AreaCode,Phon_Number  from PhoneLink inner join phone on plink_phoneid=phon_phoneid ";
                                sAreaCodeSQL += " where PLink_Deleted is null and Phon_Deleted is null  and plink_recordid=" + iEntityRecordID;
                                QuerySelect sQueryObj = GetQuery();
                                sQueryObj.SQLCommand = sAreaCodeSQL;
                                sQueryObj.ExecuteReader();
                                #endregion

                                #region Update Phone Area Code
                                if (!sQueryObj.Eof())
                                {
                                    while (!sQueryObj.Eof())
                                    {
                                        string sPhoneTypeSQL = "";
                                        string sPhoneAreaCodeSQL = "";

                                        if (!string.IsNullOrEmpty(sQueryObj.FieldValue("PLink_Type")))
                                            sPhoneTypeSQL = sQueryObj.FieldValue("PLink_Type");
                                        else
                                            sPhoneTypeSQL = "";

                                        if (!String.IsNullOrEmpty(sQueryObj.FieldValue("Phon_AreaCode")))
                                        {
                                            sPhoneAreaCodeSQL = sQueryObj.FieldValue("Phon_AreaCode");
                                        }
                                        else
                                            sPhoneAreaCodeSQL = "";

                                        if (sPhoneTypeSQL != "")
                                        {
                                            if (sPhoneTypeSQL.ToLower().TrimEnd(':') == objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'))
                                                sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_areacode" + objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'), sPhoneAreaCodeSQL, 20, 4));
                                            else
                                                sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_areacode" + objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'), "", 20, 4));
                                        }

                                        sQueryObj.Next();
                                    }
                                }
                                #endregion

                                #region Create New Phone Area Code
                                else
                                    sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_areacode" + objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'), "", 20, 4));
                                #endregion

                            }
                            #endregion

                            #region Get the Phone Detail Record for Current Record
                            string sNumberSQL = "select PLink_Type,Phon_CountryCode,Phon_AreaCode,Phon_Number  from PhoneLink inner join phone on plink_phoneid=phon_phoneid ";
                            sNumberSQL += " where PLink_Deleted is null and Phon_Deleted is null  and plink_recordid=" + iEntityRecordID;
                            QuerySelect ObjPhoneNumber = GetQuery();
                            ObjPhoneNumber.SQLCommand = sNumberSQL;
                            ObjPhoneNumber.ExecuteReader();
                            if (!ObjPhoneNumber.Eof())
                            {
                                while (!ObjPhoneNumber.Eof())
                                {
                                    string sPhoneTypeSQL = "";
                                    string sPhoneNumberSQL = "";

                                    if (!string.IsNullOrEmpty(ObjPhoneNumber.FieldValue("PLink_Type")))
                                        sPhoneTypeSQL = ObjPhoneNumber.FieldValue("PLink_Type");
                                    else
                                        sPhoneTypeSQL = "";

                                    if (!String.IsNullOrEmpty(ObjPhoneNumber.FieldValue("Phon_Number")))
                                    {
                                        sPhoneNumberSQL = ObjPhoneNumber.FieldValue("Phon_Number");
                                    }
                                    else
                                        sPhoneNumberSQL = "";
                                    if (sPhoneTypeSQL != "")
                                    {
                                        if (sPhoneTypeSQL.ToLower().TrimEnd(':') == objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'))
                                            sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_number" + objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'), sPhoneNumberSQL, 20, 10));
                                        else
                                            sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_number" + objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'), "", 20, 10));
                                    }
                                    ObjPhoneNumber.Next();
                                }
                            }
                            else
                                sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_number" + objPhoneField.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'), "", 20, 10));

                            #endregion
                            objPhoneField.GoToNext();
                        }
                    }
                }

                sHTMLPhoneDetails += HTML.EndTable();
                #endregion

                #region Build Email Structure
                Record objEmailRec = FindRecord("Custom_Captions", " Capt_family=N'Link_CompEmai' and capt_deleted is null ");
                if (!objEmailRec.Eof())
                {
                    sHTMLEmailDetails = HTML.StartTable().ToString();
                    sHTMLEmailDetails += HTML.TableData("", "", "");
                    sHTMLEmailDetails += HTML.TableRow("", "").ToString();
                    sHTMLEmailDetails += HTML.TableData("");
                    sHTMLEmailDetails += HTML.TableData("Email Address:", "VIEWBOXCAPTION");
                    sHTMLEmailDetails += HTML.TableRow("", "").ToString();

                    while (!objEmailRec.Eof())
                    {
                        string sEmailTypeCaption = objEmailRec.GetFieldAsString("Capt_US");

                        if (!String.IsNullOrEmpty(sEmailTypeCaption))
                        {
                            sHTMLEmailDetails += HTML.TableData(sEmailTypeCaption + "</b>", "VIEWBOXCAPTION");
                            #region Get the Phone Detail Record for Current Record
                            string objEmailSQL = "select * from Emaillink inner join email on elink_Emailid=emai_emailid ";
                            objEmailSQL += " where elink_Deleted is null and emai_deleted is null  and elink_recordid=" + iEntityRecordID;
                            QuerySelect sQueryObj = GetQuery();
                            sQueryObj.SQLCommand = objEmailSQL;
                            sQueryObj.ExecuteReader();
                            if (!sQueryObj.Eof())
                            {
                                while (!sQueryObj.Eof())
                                {
                                    string sEmailTypeSQL = "";
                                    string sEmailSQL = "";

                                    if (!string.IsNullOrEmpty(sQueryObj.FieldValue("eLink_Type")))
                                        sEmailTypeSQL = sQueryObj.FieldValue("eLink_Type");
                                    else
                                        sEmailTypeSQL = "";

                                    if (!String.IsNullOrEmpty(sQueryObj.FieldValue("emai_emailaddress")))
                                    {
                                        sEmailSQL = sQueryObj.FieldValue("emai_emailaddress");
                                    }
                                    else
                                        sEmailSQL = "";
                                    if (sEmailTypeSQL != "")
                                    {
                                        if (sEmailTypeSQL.ToLower().TrimEnd(':') == sEmailTypeCaption.ToLower().TrimEnd(':'))
                                        {
                                            sHTMLEmailDetails += HTML.TableData(HTML.InputText("emai_emailaddress" + objEmailRec.GetFieldAsString("Capt_US").ToLower().TrimEnd(':'), sEmailSQL, 255, 20));
                                            sHTMLEmailDetails += HTML.TableRow("", "");
                                        }
                                    }
                                    sQueryObj.Next();
                                }

                            }
                            else
                            {
                                sHTMLEmailDetails += HTML.TableData(HTML.InputText("emai_emailaddress" + sEmailTypeCaption.ToLower().TrimEnd(':'), "", 255, 20));
                                sHTMLEmailDetails += HTML.TableRow("", "");
                            }
                            #endregion

                            objEmailRec.GoToNext();
                        }
                    }

                    sHTMLEmailDetails += HTML.EndTable();

                }
                #endregion

                #region Display Blocks
                sHTMLDetails = HTML.StartTable();
                sHTMLDetails += HTML.TableData("", "", "");
                sHTMLDetails += HTML.TableRow("", "").ToString();
                sHTMLDetails += HTML.TableData(HTML.Box("Phone", sHTMLPhoneDetails), "");
                sHTMLDetails += HTML.TableData("<Table Width=1px '> <TR> <TD> </TD></TR> </Table>");
                sHTMLDetails += HTML.TableData(HTML.Box("E-mail", sHTMLEmailDetails), "", "table-layout:fixed");
                AddContent(HTML.Box("", sHTMLDetails));
                #endregion

                #region Add Buttons
                AddUrlButton("Save", "save.gif", "javascript:SetHiddenParam();");
                AddUrlButton("Cancel", "Cancel.gif", "");
                #endregion

                #region Save Phone/Email Data

                if (!String.IsNullOrEmpty(Dispatch.ContentField("HIDDENPAGENUMBER_Save")))
                {
                    if (Dispatch.ContentField("HIDDENPAGENUMBER_Save") == "Save")
                    {
                        int iEntityId = 0;
                        Record objEntityIdRec = FindRecord("custom_tables", "bord_name='Building' and bord_deleted is null ");
                        if (!objEntityIdRec.Eof())
                            iEntityId = objEntityIdRec.GetFieldAsInt("Bord_tableid");

                        #region Insert/Update Phone Details in CRM
                        Record objPhoneTranslations = FindRecord("Custom_captions", "  Capt_Family = N'Link_CompPhon' and capt_deleted is null ");
                        if (!objPhoneTranslations.Eof())
                        {
                            while (!objPhoneTranslations.Eof())
                            {
                                string sPhoneTypeTranslation = objPhoneTranslations.GetFieldAsString("Capt_US");
                                if (!String.IsNullOrEmpty(Dispatch.ContentField("phon_countrycode" + sPhoneTypeTranslation.ToLower().TrimEnd(':'))) && !String.IsNullOrEmpty(Dispatch.ContentField("phon_areacode" + sPhoneTypeTranslation.ToLower().TrimEnd(':'))) && !String.IsNullOrEmpty(Dispatch.ContentField("phon_number" + sPhoneTypeTranslation.ToLower().TrimEnd(':'))))
                                {
                                    Record objPhoneLinkRec = FindRecord("PhoneLink", "plink_deleted is null and plink_EntityID=" + iEntityId + " and plink_RecordID=" + iEntityRecordID + " and LOWER(LTRIM(RTRIM(plink_type)))='" + sPhoneTypeTranslation.ToLower() + "' ");
                                    if (!objPhoneLinkRec.Eof())
                                    {
                                        while (!objPhoneLinkRec.Eof())
                                        {
                                            int iPhoneId = objPhoneLinkRec.GetFieldAsInt("plink_phoneid");
                                            string sPhonetype = "";
                                            if (!string.IsNullOrEmpty(objPhoneLinkRec.GetFieldAsString("plink_type")))
                                                sPhonetype = objPhoneLinkRec.GetFieldAsString("plink_type");
                                            else
                                                sPhonetype = "";
                                            if (sPhoneTypeTranslation.ToLower() == sPhonetype.ToLower())
                                            {
                                                string sPhoneSQL = "UPDATE Phone SET Phon_CountryCode=" + Dispatch.ContentField("phon_countrycode" + sPhoneTypeTranslation.ToLower().Trim(':'));
                                                sPhoneSQL += " , Phon_AreaCode=" + Dispatch.ContentField("phon_areacode" + sPhoneTypeTranslation.ToLower().Trim(':'));
                                                sPhoneSQL += " ,Phon_Number=" + Dispatch.ContentField("phon_number" + sPhoneTypeTranslation.ToLower().Trim(':'));
                                                sPhoneSQL += " where (phon_phoneid=(Select plink_phoneid from phonelink where plink_Type='" + sPhoneTypeTranslation + "' and plink_phoneid=" + iPhoneId + " and phon_deleted is null))";
                                                QuerySelect PhoneNumberObj = GetQuery();
                                                PhoneNumberObj.SQLCommand = sPhoneSQL;
                                                PhoneNumberObj.ExecuteReader();
                                            }

                                            objPhoneLinkRec.GoToNext();
                                        }
                                    }
                                    else
                                    {
                                        if (!String.IsNullOrEmpty(Dispatch.ContentField("phon_countrycode" + sPhoneTypeTranslation.ToLower().TrimEnd(':'))) && !String.IsNullOrEmpty(Dispatch.ContentField("phon_areacode" + sPhoneTypeTranslation.ToLower().TrimEnd(':'))) && !String.IsNullOrEmpty(Dispatch.ContentField("phon_number" + sPhoneTypeTranslation.ToLower().TrimEnd(':'))))
                                        {
                                            #region Insert New Record in Phone Entity
                                            Record objNewPhoneRec = new Record("Phone");
                                            objNewPhoneRec.SetField("Phon_CountryCode", Dispatch.ContentField("phon_countrycode" + sPhoneTypeTranslation.ToLower().Trim(':')));
                                            objNewPhoneRec.SetField("Phon_AreaCode", Dispatch.ContentField("phon_areacode" + sPhoneTypeTranslation.ToLower().Trim(':')));
                                            objNewPhoneRec.SetField("Phon_Number", Dispatch.ContentField("phon_number" + sPhoneTypeTranslation.ToLower().Trim(':')));
                                            objNewPhoneRec.SaveChanges();
                                            #endregion

                                            #region Insert New Record in Phone Link Entity
                                            Record objNewPhoneLink = new Record("PhoneLink");
                                            objNewPhoneLink.SetField("plink_entityid", iEntityId);
                                            objNewPhoneLink.SetField("Plink_RecordId", iEntityRecordID);
                                            objNewPhoneLink.SetField("Plink_type", sPhoneTypeTranslation.TrimEnd(':'));
                                            objNewPhoneLink.SetField("PLink_PhoneId", objNewPhoneRec.RecordId);
                                            objNewPhoneLink.SaveChanges();
                                            #endregion
                                        }

                                    }
                                }
                                objPhoneTranslations.GoToNext();
                            }
                        }

                        /*#region Create New Phone Record in CRM
                        else
                        {
                            Record objPhoneTranslationRec = FindRecord("Custom_captions", "  Capt_Family = N'Link_CompPhon' and capt_deleted is null ");

                            if (!objPhoneTranslationRec.Eof())
                            {
                                while (!objPhoneTranslationRec.Eof())
                                {
                                    string sPhoneType = objPhoneTranslationRec.GetFieldAsString("Capt_US");

                                    if (!String.IsNullOrEmpty(Dispatch.ContentField("phon_countrycode" + sPhoneType.ToLower().TrimEnd(':'))) && !String.IsNullOrEmpty(Dispatch.ContentField("phon_areacode" + sPhoneType.ToLower().TrimEnd(':'))) && !String.IsNullOrEmpty(Dispatch.ContentField("phon_number" + sPhoneType.ToLower().TrimEnd(':'))))
                                    {
                                        #region Insert New Record in Phone Entity
                                        Record objNewPhoneRec = new Record("Phone");
                                        objNewPhoneRec.SetField("Phon_CountryCode", Dispatch.ContentField("phon_countrycode" + sPhoneType.ToLower().Trim(':')));
                                        objNewPhoneRec.SetField("Phon_AreaCode", Dispatch.ContentField("phon_areacode" + sPhoneType.ToLower().Trim(':')));
                                        objNewPhoneRec.SetField("Phon_Number", Dispatch.ContentField("phon_number" + sPhoneType.ToLower().Trim(':')));
                                        objNewPhoneRec.SaveChanges();
                                        #endregion
                                        
                                        #region Insert New Record in Phone Link Entity
                                        Record objNewPhoneLink = new Record("PhoneLink");
                                        objNewPhoneLink.SetField("plink_entityid", iEntityId);
                                        objNewPhoneLink.SetField("Plink_RecordId", iEntityRecordID);
                                        objNewPhoneLink.SetField("Plink_type", sPhoneType.TrimEnd(':'));
                                        objNewPhoneLink.SetField("PLink_PhoneId", objNewPhoneRec.RecordId);
                                        objNewPhoneLink.SaveChanges();
                                        #endregion
                                    }

                                    objPhoneTranslations.GoToNext();
                                }
                            }
                        }
                        #endregion*/
                #endregion

                        #region Insert/Update Email Details in Sage CRM

                        Record objEmailTranslation = FindRecord("Custom_Captions", " Capt_family=N'Link_CompEmai' and capt_deleted is null ");
                        if (!objEmailTranslation.Eof())
                        {
                            while (!objEmailTranslation.Eof())
                            {
                                string sEmailTypeTranslation = objEmailTranslation.GetFieldAsString("capt_us");
                                sEmailTypeTranslation = sEmailTypeTranslation.ToLower().Trim(':');
                                sEmailTypeTranslation = sEmailTypeTranslation.ToLower().Trim(':');

                                Record objEmailLinkRec = FindRecord("emailLink", "elink_deleted is null and LTRIM(RTRIM(elink_recordid))=" + iEntityRecordID + " and LTRIM(RTRIM(elink_entityid))=" + iEntityId + " and LOWER(LTRIM(RTRIM(elink_type)))='" + sEmailTypeTranslation + "' ");
                                AddContent(objEmailLinkRec.RecordCount.ToString());
                                if (!objEmailLinkRec.Eof())
                                {
                                    while (!objEmailLinkRec.Eof())
                                    {
                                        int iEmailId = objEmailLinkRec.GetFieldAsInt("Elink_Emailid");
                                        string sEmailType = "";
                                        if (!string.IsNullOrEmpty(objEmailLinkRec.GetFieldAsString("elink_type")))
                                            sEmailType = objEmailLinkRec.GetFieldAsString("elink_type");
                                        else
                                            sEmailType = "";
                                        if (sEmailTypeTranslation == sEmailType)
                                        {
                                            string sEmailSQL = "UPDATE Email SET Emai_EmailAddress='" + Dispatch.ContentField("emai_emailaddress" + sEmailTypeTranslation) + "'";
                                            sEmailSQL += " where (emai_emailid=(Select elink_emailid from emaillink where Elink_Type='" + sEmailTypeTranslation + "' ";
                                            sEmailSQL += " and elink_emailid=" + iEmailId + " and ELink_Deleted is null)) ";

                                            QuerySelect EmailAddressrObj = GetQuery();
                                            EmailAddressrObj.SQLCommand = sEmailSQL;
                                            EmailAddressrObj.ExecuteReader();
                                        }

                                        objEmailLinkRec.GoToNext();

                                    }
                                }
                                else
                                {
                                    #region Insert New Record in Email Table
                                    Record objNewEmailRec = new Record("Email");
                                    objNewEmailRec.SetField("emai_emailaddress", Dispatch.ContentField("emai_emailaddress" + sEmailTypeTranslation.ToLower().Trim(':')));
                                    objNewEmailRec.SaveChanges();
                                    #endregion

                                    #region Insert New Record In EmailLink Table
                                    Record objNewEmailLinkRec = new Record("Emaillink");
                                    objNewEmailLinkRec.SetField("elink_entityid", iEntityId);
                                    objNewEmailLinkRec.SetField("elink_recordid", iEntityRecordID);
                                    objNewEmailLinkRec.SetField("elink_type", sEmailTypeTranslation.Trim(':'));
                                    objNewEmailLinkRec.SetField("elink_emailid", objNewEmailRec.RecordId);
                                    objNewEmailLinkRec.SaveChanges();
                                    #endregion
                                }
                                objEmailTranslation.GoToNext();
                            }
                        }

                        #endregion

                        #endregion
                        string sURL = UrlDotNet(this.ThisDotNetDll, "RunBuildingPhoneEmailPage");
                        Dispatch.Redirect(sURL);
                    }

                }

            }

            catch (Exception error)
            {
                this.AddError(error.Message);
            }
        }

    }
}

