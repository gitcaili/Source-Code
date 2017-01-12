using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;

namespace NZPACRM
{
    public class PublicationPhoneEmailPage : Web
    {
        #region Variable Declartion
        int iEntityId = 0;
        string sHTMLPhoneDetails = "";
        string sHTMLEmailDetails = "";
        string sUseCountryCode = "";
        string sUseAreaCode = "";
        #endregion 
        public PublicationPhoneEmailPage()
        {
            #region Get Tab and Top content
            GetTabs("Publications");
            AddTopContent(GetCustomEntityTopFrame("Publications"));
            #endregion  
        }
        public override void BuildContents()
        {
            try
            {
                #region Add Html Form and javascript
                AddContent(HTML.Form());
                AddContent("<script type='text/javascript' src='../CustomPages/Client/ClientFuncs.js'></script>");
                #endregion

                int iEntityRecordID = 0;

                if (!String.IsNullOrEmpty(Dispatch.EitherField("pblc_PublicationsID")))
                    iEntityRecordID = Convert.ToInt32(Dispatch.EitherField("pblc_PublicationsID"));
                else
                    iEntityRecordID = Convert.ToInt32(Dispatch.EitherField("Key58"));


                Record objEntityIdRec = FindRecord("custom_tables", "bord_name='Publications' and bord_deleted is null ");
                if (!objEntityIdRec.Eof())
                    iEntityId = objEntityIdRec.GetFieldAsInt("Bord_tableid");

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



                #region Save Phone Email Record

                if (!String.IsNullOrEmpty(Dispatch.EitherField("HiddenMode")))
                {
                    if (Dispatch.EitherField("HiddenMode") == "save")
                    {
                        InsertUpdatePhoneRecord(iEntityRecordID);
                        InsertUpdateEmailRecord(iEntityRecordID);

                        string sURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName;
                        sURL += "/CustomPages/Publications/PublicationsSummary.asp?SID=" + Dispatch.EitherField("SID");
                        sURL += "&Key0=58&F=&J=Publications/PublicationsSummary.asp&pblc_PublicationsID=" + iEntityRecordID + "&T=Publications";
                        Dispatch.Redirect(sURL);
                    }
                }
                else
                {
                    AddContent(HTML.InputHidden("HiddenMode", ""));
                }
                #endregion

                Record objPhoneRec = FindRecord("PhoneLink", "PLink_RecordID=" + iEntityRecordID);
                int iCompanyPhoneCount = 0;
                int iCompanyEmailCount = 0;
                #region Build Phone UI
                sHTMLPhoneDetails = HTML.StartTable().ToString();
                sHTMLPhoneDetails += HTML.TableData("", "", "");
                sHTMLPhoneDetails += HTML.TableRow("", "").ToString();
                sHTMLPhoneDetails += HTML.TableData("");
                if (sUseCountryCode == "Y")
                    sHTMLPhoneDetails += HTML.TableData("Country", "VIEWBOXCAPTION");
                if (sUseAreaCode == "Y")
                    sHTMLPhoneDetails += HTML.TableData("Area", "VIEWBOXCAPTION");

                sHTMLPhoneDetails += HTML.TableData("Number", "VIEWBOXCAPTION");

                Record objPhoneField = FindRecord("Custom_captions", "  Capt_Family = N'Link_CompPhon' and capt_deleted is null ");
                objPhoneField.OrderBy = "capt_order";
                if (!objPhoneField.Eof())
                {
                    while (!objPhoneField.Eof())
                    {
                        iCompanyPhoneCount++;
                        sHTMLPhoneDetails += HTML.TableRow("", "");
                        sHTMLPhoneDetails += HTML.TableData("<b>" + "&nbsp;" + objPhoneField.GetFieldAsString("Capt_US") + "</b>", "VIEWBOXCAPTION");

                        string sPhoneType = objPhoneField.GetFieldAsString("Capt_US");
                        if (sUseCountryCode == "Y")
                        {
                            sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_countrycode" + objPhoneField.GetFieldAsString("Capt_Code"), GetPhoneValue(iEntityRecordID, iEntityId, objPhoneField.GetFieldAsString("Capt_Code"), "country"), 5, 3));
                        }

                        if (sUseAreaCode == "Y")
                        {
                            sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_areacode" + objPhoneField.GetFieldAsString("Capt_Code"), GetPhoneValue(iEntityRecordID, iEntityId, objPhoneField.GetFieldAsString("Capt_Code"), "area"), 20, 4));
                        }

                        sHTMLPhoneDetails += HTML.TableData(HTML.InputText("phon_number" + objPhoneField.GetFieldAsString("Capt_Code"), GetPhoneValue(iEntityRecordID, iEntityId, objPhoneField.GetFieldAsString("Capt_Code"), "phone"), 20, 27));

                        objPhoneField.GoToNext();
                    }
                }
                sHTMLPhoneDetails += HTML.EndTable();

                #endregion

                #region Build Email UI
                Record objEmailRec = FindRecord("Custom_Captions", " Capt_family=N'Link_CompEmai' and capt_deleted is null ");
                objEmailRec.OrderBy = "capt_order";

                sHTMLEmailDetails = "<label id='lblEmailInfo'>" + HTML.StartTable().ToString();
                sHTMLEmailDetails += HTML.TableData("", "", "");
                sHTMLEmailDetails += HTML.TableRow("", "").ToString();
                sHTMLEmailDetails += HTML.TableData("");
                sHTMLEmailDetails += HTML.TableData("Email Address:", "VIEWBOXCAPTION");
                sHTMLEmailDetails += HTML.TableRow("", "").ToString();

                if (!objEmailRec.Eof())
                {
                    while (!objEmailRec.Eof())
                    {
                        iCompanyEmailCount++;
                        string sEmailTypeCaption = objEmailRec.GetFieldAsString("Capt_US");

                        sHTMLEmailDetails += HTML.TableData("&nbsp;" + sEmailTypeCaption + "</b>", "VIEWBOXCAPTION");
                        sHTMLEmailDetails += HTML.TableData(HTML.InputText("emai_emailaddress" + objEmailRec.GetFieldAsString("Capt_code"), GetEmailValue(iEntityRecordID, iEntityId, objEmailRec.GetFieldAsString("Capt_code")), 255, 30));
                        sHTMLEmailDetails += HTML.TableRow("", "");
                        objEmailRec.GoToNext();
                    }
                    
                }
                #endregion
                if (iCompanyPhoneCount > iCompanyEmailCount)
                {
                    int iDiffRows = iCompanyPhoneCount - iCompanyEmailCount;
                    if (iDiffRows > 0)
                    {
                        for (int i = 0; i < iDiffRows; i++)
                        {
                            sHTMLEmailDetails += HTML.TableData("&nbsp;");
                            sHTMLEmailDetails += HTML.TableRow("", "");

                        }

                    }
                }
                sHTMLEmailDetails += HTML.EndTable() + "</label>";               

                AddContent("<table width ='100%' valign='top'><tr valign='top'><td>" + HTML.TableData(HTML.Box("Phone", sHTMLPhoneDetails), "") + "</td><td>" + HTML.TableData(HTML.Box("E-mail ", sHTMLEmailDetails), "", "") + "</td></tr></tr></table>");

                #region Add Buttons
                AddUrlButton("Save", "save.gif", "javascript:if(ValidateEmail()==true){document.EntryForm.HiddenMode.value='save';document.EntryForm.submit();}");
                string sCancelURL = Url("Publications/PublicationsSummary.asp") + "&pblc_PublicationsID=" + iEntityRecordID;
                AddUrlButton("Cancel", "Cancel.gif", sCancelURL);
                #endregion
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
        }
        public void InsertUpdatePhoneRecord(int iEntityRecordID)
        {
            Record objPhoneField = FindRecord("Custom_captions", "Capt_Family = N'Link_CompPhon' and capt_deleted is null ");
            objPhoneField.OrderBy = "capt_order";
            if (!objPhoneField.Eof())
            {
                while (!objPhoneField.Eof())
                {
                    if (!String.IsNullOrEmpty(Dispatch.ContentField("phon_countrycode" + objPhoneField.GetFieldAsString("capt_code"))) || !String.IsNullOrEmpty(Dispatch.ContentField("phon_areacode" + objPhoneField.GetFieldAsString("capt_code"))) || !String.IsNullOrEmpty(Dispatch.ContentField("phon_number" + objPhoneField.GetFieldAsString("capt_code"))))
                    {
                        Record objPhoneRecord = FindRecord("PhoneLink", "PLink_EntityID='" + iEntityId + "' and plink_recordid='" + iEntityRecordID + "' and plink_type='" + objPhoneField.GetFieldAsString("capt_code") + "'");
                        if (!objPhoneRecord.Eof())
                        {
                            string sPhoneSQL = "UPDATE Phone SET Phon_CountryCode='" + Dispatch.ContentField("phon_countrycode" + objPhoneField.GetFieldAsString("capt_code")) + "'";
                            sPhoneSQL += " , Phon_AreaCode='" + Dispatch.ContentField("phon_areacode" + objPhoneField.GetFieldAsString("capt_code")) + "'";
                            sPhoneSQL += " ,Phon_Number='" + Dispatch.ContentField("phon_number" + objPhoneField.GetFieldAsString("capt_code")) + "'";
                            sPhoneSQL += " where (phon_phoneid=(Select plink_phoneid from phonelink where plink_Type='" + objPhoneField.GetFieldAsString("capt_code") + "' and plink_phoneid=" + objPhoneRecord.GetFieldAsString("plink_phoneid") + " and phon_deleted is null))";
                            QuerySelect PhoneNumberObj = GetQuery();
                            PhoneNumberObj.SQLCommand = sPhoneSQL;
                            PhoneNumberObj.ExecuteReader();
                        }
                        else
                        {
                            #region Insert New Record in Phone Entity
                            Record objNewPhoneRec = new Record("Phone");
                            objNewPhoneRec.SetField("Phon_CountryCode", Dispatch.ContentField("phon_countrycode" + objPhoneField.GetFieldAsString("capt_code")));
                            objNewPhoneRec.SetField("Phon_AreaCode", Dispatch.ContentField("phon_areacode" + objPhoneField.GetFieldAsString("capt_code")));
                            objNewPhoneRec.SetField("Phon_Number", Dispatch.ContentField("phon_number" + objPhoneField.GetFieldAsString("capt_code")));
                            objNewPhoneRec.SaveChanges();
                            #endregion

                            #region Insert New Record in Phone Link Entity
                            Record objNewPhoneLink = new Record("PhoneLink");
                            objNewPhoneLink.SetField("plink_entityid", iEntityId);
                            objNewPhoneLink.SetField("Plink_RecordId", iEntityRecordID);
                            objNewPhoneLink.SetField("Plink_type", objPhoneField.GetFieldAsString("capt_code"));
                            objNewPhoneLink.SetField("PLink_PhoneId", objNewPhoneRec.RecordId);
                            objNewPhoneLink.SaveChanges();
                            #endregion
                        }
                    }
                    else
                    {
                        Record objPhoneRecord = FindRecord("PhoneLink", "PLink_EntityID='" + iEntityId + "' and plink_recordid='" + iEntityRecordID + "' and plink_type='" + objPhoneField.GetFieldAsString("capt_code") + "'");
                        if (!objPhoneRecord.Eof())
                        {
                            string sPhoneSQL = "UPDATE Phone SET Phon_CountryCode='', Phon_AreaCode='',Phon_Number='' where (phon_phoneid=(Select plink_phoneid from phonelink where plink_Type='" + objPhoneField.GetFieldAsString("capt_code") + "' and plink_phoneid=" + objPhoneRecord.GetFieldAsString("plink_phoneid") + " and phon_deleted is null))";
                            QuerySelect PhoneNumberObj = GetQuery();
                            PhoneNumberObj.SQLCommand = sPhoneSQL;
                            PhoneNumberObj.ExecuteReader();
                        }
                    }
                    objPhoneField.GoToNext();
                }
            }
        }

        public void InsertUpdateEmailRecord(int iEntityRecordID)
        {
            Record objEmailField = FindRecord("Custom_captions", "Capt_Family = N'Link_CompEmai' and capt_deleted is null ");
            objEmailField.OrderBy = "capt_order";
            if (!objEmailField.Eof())
            {
                while (!objEmailField.Eof())
                {
                    if (!String.IsNullOrEmpty(Dispatch.ContentField("emai_emailaddress" + objEmailField.GetFieldAsString("capt_code"))))
                    {
                        Record objEmailRecord = FindRecord("EmailLink", "ELink_EntityID='" + iEntityId + "' and Elink_recordid='" + iEntityRecordID + "' and Elink_type='" + objEmailField.GetFieldAsString("capt_code") + "'");
                        if (!objEmailRecord.Eof())
                        {
                            string sEmailSQL = "UPDATE Email SET Emai_EmailAddress='" + Dispatch.ContentField("emai_emailaddress" + objEmailField.GetFieldAsString("capt_code")) + "'";
                            sEmailSQL += " where (emai_emailid=(Select Elink_Emailid from Emaillink where Elink_Type='" + objEmailField.GetFieldAsString("capt_code") + "' and Elink_Emailid=" + objEmailRecord.GetFieldAsString("Elink_Emailid") + " and emai_deleted is null))";
                            QuerySelect EmailObj = GetQuery();
                            EmailObj.SQLCommand = sEmailSQL;
                            EmailObj.ExecuteReader();
                        }
                        else
                        {
                            #region Insert New Record in Phone Entity
                            Record objNewEmailRec = new Record("Email");
                            objNewEmailRec.SetField("emai_emailaddress", Dispatch.ContentField("emai_emailaddress" + objEmailField.GetFieldAsString("capt_code")));
                            objNewEmailRec.SaveChanges();
                            #endregion

                            #region Insert New Record In EmailLink Table
                            Record objNewEmailLinkRec = new Record("Emaillink");
                            objNewEmailLinkRec.SetField("elink_entityid", iEntityId);
                            objNewEmailLinkRec.SetField("elink_recordid", iEntityRecordID);
                            objNewEmailLinkRec.SetField("elink_type", objEmailField.GetFieldAsString("capt_code"));
                            objNewEmailLinkRec.SetField("elink_emailid", objNewEmailRec.RecordId);
                            objNewEmailLinkRec.SaveChanges();
                            #endregion
                        }
                    }
                    else
                    {
                        Record objPhoneRecord = FindRecord("EmailLink", "ELink_EntityID='" + iEntityId + "' and Elink_recordid='" + iEntityRecordID + "' and Elink_type='" + objEmailField.GetFieldAsString("capt_code") + "'");
                        if (!objPhoneRecord.Eof())
                        {
                            string sPhoneSQL = "UPDATE Email SET Emai_EmailAddress='' where (emai_emailid=(Select elink_emailid from Emaillink where elink_Type='" + objEmailField.GetFieldAsString("capt_code") + "' and elink_emailid=" + objPhoneRecord.GetFieldAsString("elink_emailid") + " and emai_deleted is null))";
                            QuerySelect PhoneNumberObj = GetQuery();
                            PhoneNumberObj.SQLCommand = sPhoneSQL;
                            PhoneNumberObj.ExecuteReader();
                        }
                    }
                    objEmailField.GoToNext();
                }
            }
        }
        public string GetPhoneValue(int iEntityRecordID, int iEntityID, string sType, string sPhoneCode)
        {
            string sPhoneValue = "";
            string sSQL = " SELECT PhoneLink.PLink_Type, Phone.Phon_CountryCode, Phone.Phon_AreaCode, Phone.Phon_Number ";
            sSQL += " FROM PhoneLink INNER JOIN Phone ON PhoneLink.PLink_PhoneId = Phone.Phon_PhoneId ";
            sSQL += " WHERE (PhoneLink.PLink_EntityID = '" + iEntityID + "') AND (PhoneLink.PLink_RecordID = '" + iEntityRecordID + "') AND (PhoneLink.PLink_Deleted IS NULL) AND (PhoneLink.PLink_Type = N'" + sType + "') ";
            //AddContent(sSQL + "<br />");
            QuerySelect objPhoneRec = GetQuery();
            objPhoneRec.SQLCommand = sSQL;
            objPhoneRec.ExecuteReader();

            if (!objPhoneRec.Eof())
            {
                if (sPhoneCode == "country")
                    sPhoneValue = objPhoneRec.FieldValue("Phon_CountryCode");
                else if (sPhoneCode == "area")
                    sPhoneValue = objPhoneRec.FieldValue("Phon_AreaCode");
                else if (sPhoneCode == "phone")
                    sPhoneValue = objPhoneRec.FieldValue("Phon_Number");
            }

            return sPhoneValue;
        }

        public string GetEmailValue(int iEntityRecordID, int iEntityID, string sType)
        {
            string sEmailAddress = "";
            string sSQL = " SELECT EmailLink.ELink_RecordID, EmailLink.ELink_Type, Email.Emai_EmailAddress, EmailLink.ELink_EntityID, EmailLink.ELink_Deleted ";
            sSQL += " FROM EmailLink INNER JOIN Email ON EmailLink.ELink_EmailId = Email.Emai_EmailId ";
            sSQL += " WHERE (EmailLink.ELink_RecordID = '" + iEntityRecordID + "') AND (EmailLink.ELink_Type = N'" + sType + "') AND (EmailLink.ELink_EntityID = '" + iEntityID + "') AND (EmailLink.ELink_Deleted IS NULL)";

            QuerySelect objEmailRec = GetQuery();
            objEmailRec.SQLCommand = sSQL;
            objEmailRec.ExecuteReader();

            if (!objEmailRec.Eof())
            {
                sEmailAddress = objEmailRec.FieldValue("Emai_EmailAddress");
            }
            return sEmailAddress;
        }
    }    

}
