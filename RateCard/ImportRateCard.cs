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
    public class ImportRateCard : DataPageNew
    {
        CRMHelper objCRM = new CRMHelper();
        string LogfileName = "";
        string shttpURL = "";
        public ImportRateCard()
            : base("RatesCard", "rate_RatesCardID", "")
        {
            string CurrUser = CurrentUser.UserId.ToString();

            #region get Http from url
            try
            {
                string s = Dispatch.ServerVariable("HTTP_REFERER");
                char[] cSplit = { '/' };
                string[] sHTTP = s.Split(cSplit);

                if (!String.IsNullOrEmpty(sHTTP[0]))
                    shttpURL = sHTTP[0];

                if (CurrentUser.SessionRead("HTTP_REFERER") == null)
                {
                    CurrentUser.SessionWrite("HTTP_REFERER", shttpURL);
                }
            }
            catch (Exception ex)
            {
                if (CurrentUser.SessionRead("HTTP_REFERER") != null)
                {
                    shttpURL = CurrentUser.SessionRead("HTTP_REFERER").ToString();
                }
            }

            string sHostName = Dispatch.Host;
            string sInstallName = Dispatch.InstallName;
            #endregion

        }
        public override void BuildContents()
        {
            try
            {
                AddContent("<script type='text/javascript' src='../CustomPages/Booking/ClientFuncs.js'></script>");
                int ImportCount = 0;
                string EntityID = "0";
                string sPublication = "";
                string sRateCard = "";
                string Section = "";
                string sSubsection = "";
                string sSectionCode = "";
                string sCategory = "";
                string sCategoryCode = "";
                string sPageColor = "";
                string sPageColorCode = "";
                string sMondayRate = "";
                string sTuesdayRate = "";
                string sWednesdayRate = "";
                string sThursdayRate = "";
                string sFridayRate = "";
                string sSaturdayRate = "";
                string sSundayRate = "";
                decimal iMondayRate = 0.00m;
                decimal iTuesdayRate = 0.00m;
                decimal iWednesdayRate = 0.00m;
                decimal iThrusdayRate = 0.00m;
                decimal iFridayRate = 0.00m;
                decimal iSaturdayRate = 0.00m;
                decimal iSundayRate = 0.00m;
                decimal iheight = 0.00m;
                decimal iwidth = 0.00m;



                string sSunday = "";
                string sMonday = "";
                string sTesuday = "";
                string sWednesday = "";
                string sThursday = "";
                string sFriday = "";
                string sSaturday = "";

                string sSize = "";
                string sSizeCode = "";
                string sStandardSize = "";
                string sStandardSizeCode = "";
                string sHeight = "";
                string sWidth = "";
                string sClassType = "";
                string sClassTypeCode = "";
                string sLoading = "";
                decimal iLoading = 0.00m;
                string sSuccessErrorMessage = "";
                int iFailedCount = 0;
                int InsertCount = 0;
                int iDupRecord = 0;
                string isDatavalid = "";
                string sBaseCurrency = "";
                string sName = "";
                string sRateCardDescription = "";

                string sRateMaxHeight = "";
                string sRateWidth1 = "";
                string sRateWidth2 = "";
                string sRateWidth3 = "";
                string sRateWidth4 = "";
                string sRateWidth5 = "";
                string sRateWidth6 = "";
                string sRateWidth7 = "";
                string sRateWidth8 = "";
                string sRateWidth9 = "";
                string sRateWidth10 = "";
                string sRateWidth11 = "";
                string sRateWidth12 = "";

                int iRateMaxHeight = 0;
                int iRateWidth1 = 0;
                int iRateWidth2 = 0;
                int iRateWidth3 = 0;
                int iRateWidth4 = 0;
                int iRateWidth5 = 0;
                int iRateWidth6 = 0;
                int iRateWidth7 = 0;
                int iRateWidth8 = 0;
                int iRateWidth9 = 0;
                int iRateWidth10 = 0;
                int iRateWidth11 = 0;
                int iRateWidth12 = 0;

                #region Adding Html Form
                AddContent(HTML.Form());
                #endregion

                #region Get Template Block
                objCRM.GetTemplateBlock("Rate Card");
                #endregion

                string SavedFilePath = string.Empty;

                #region Adding Html Form
                AddContent(HTML.Form());
                #endregion

                #region Get Fileupload Control on screen
                AddContent("<BR><BR>" + HTML.Box("File", "<br>&nbsp;&nbsp;<input type='file' id='fileupload' name='pic' size='70'>&nbsp;<input type='BUTTON' class='Edit'value='Import'name='upload' onclick='javascipt:CheckFile();'></br></br>"));
                #endregion

                #region Add Buttons

                string backURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                backURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=1650&Mode=1&CLk=T&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data Management";
                AddUrlButton("Back", "prevcircle.gif", backURL);
                #endregion

                #region Define the Hidden Fields
                AddContent(HTML.InputHidden("HIDDEN_FilePath", ""));
                AddContent(HTML.InputHidden("HIDDEN_Save", ""));
                AddContent(HTML.InputHidden("HIDDEN_FileName", ""));
                #endregion

                
                if (!String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_Save")))
                {
                    if (Dispatch.ContentField("HIDDEN_Save") == "Save")
                    {
                        DataSet ds = new DataSet();
                        DataTable dt = new DataTable();
                        if (!String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_FileName")))
                        {
                            
                            #region Save File in Rate Card folder
                            SavedFilePath = SaveRateCardRLocation();                            
                            #endregion
                            
                            #region Read Excel File data
                            string extention = Path.GetExtension(SavedFilePath);
                            dt = objCRM.ConvertToDataTable(SavedFilePath);
                            
                            #region Find EntityID
                            Record RecEntityID = FindRecord("Custom_Tables", "bord_name='ratescard' and bord_deleted is null");
                            if (!RecEntityID.Eof())
                            {
                                EntityID = RecEntityID.GetFieldAsString("Bord_TableId");
                            }
                            #endregion

                            sBaseCurrency = GetBaseCurrency();
                            if (dt.Rows.Count > 0)
                            {                                
                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    sRateCardDescription = "";
                                    string sSectionFlag = "";
                                    string sPublicationId = "";
                                    string isValid = "true";
                                    string isColumnValid = "true";
                                    int iRowCurrent = i + 2;
                                    string sExist = "T";
                                    string sSectionId = "";
                                    string ssectionName = "";
                                    string sSubSectionID = "";
                                    int iRecordcount = 0;
                                    string sAllday = "";
                                    
                                    try
                                    {
                                        if (!String.IsNullOrEmpty(dt.Rows[i]["Publication"].ToString().Trim()))
                                        {
                                            sRateCardDescription = dt.Rows[i]["Publication"].ToString().Trim();

                                            sPublication = dt.Rows[i]["Publication"].ToString().Trim().Replace("'", "''");
                                            Record objPublication = FindRecord("publicationS", "LOWER(LTRIM(RTRIM(pblc_Name)))='" + sPublication.ToLower() + "'");
                                            
                                            if (!objPublication.Eof())
                                            {
                                                sPublicationId = objPublication.GetFieldAsString("pblc_PublicationsID");

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Card Type"].ToString().Trim()))
                                                {
                                                    sRateCard = dt.Rows[i]["Card Type"].ToString().Trim();
                                                    sRateCard = GetCaptionCode(sRateCard, "pnbr_commissiontype");
                                                    sRateCardDescription += " " + dt.Rows[i]["Card Type"].ToString().Trim();
                                                }
                                                else
                                                    sRateCard = "";

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Sections"].ToString().Trim()))
                                                {
                                                    Section = dt.Rows[i]["Sections"].ToString().Trim();
                                                    
                                                    Record objSections = FindRecord("sections", "sctn_name='" + Section + "'");
                                                    if (!objSections.Eof())
                                                    {
                                                        sSectionId = objSections.GetFieldAsString("sctn_Sctn_sectionid");                                                        
                                                    }
                                                    else
                                                    {
                                                        sSuccessErrorMessage += Environment.NewLine + "Row no: " + iRowCurrent + " Section " + Section + " does not exist in Sage CRM";
                                                        sSectionFlag = "F";

                                                        Record objSectionsNew = new Record("sections");
                                                        objSectionsNew.SetField("sctn_name", Section);
                                                        objSectionsNew.SetField("sctn_publicationID", sPublicationId);
                                                        
                                                        objSectionsNew.SaveChanges();
                                                        sSectionId = objSectionsNew.RecordId.ToString();
                                                    }
                                                    sRateCardDescription += " " + dt.Rows[i]["Sections"].ToString().Trim();
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Sub Sections"].ToString().Trim()))
                                                {
                                                    
                                                    sSubsection = dt.Rows[i]["Sub Sections"].ToString().Trim();

                                                    Record objSubSectionRec = FindRecord("SubSection", "suse_name='" + sSubsection + "'");
                                                    if (!objSubSectionRec.Eof())
                                                    {
                                                        sSectionId = objSubSectionRec.GetFieldAsString("suse_section");
                                                        sSubSectionID = objSubSectionRec.GetFieldAsString("suse_subsectionid");
                                                        if (sSectionId == "" || sSectionId == "null" || sSectionId == "undefined") sSectionId = "";

                                                        if (sSectionId != "")
                                                        {
                                                            Record objSections = FindRecord("sections", "sctn_Sctn_sectionid='" + sSectionId + "'");
                                                            if (!objSections.Eof())
                                                            {
                                                                ssectionName = objSections.GetFieldAsString("sctn_name");

                                                                if (ssectionName == "" || ssectionName == "null" || ssectionName == "undefined") ssectionName = "";

                                                                if (ssectionName.ToLower() != Section)
                                                                {
                                                                    sSuccessErrorMessage += Environment.NewLine + "Row no: " + iRowCurrent + " Sections Name difference encountered in Excel file and Sage CRM for " + sSubsection;
                                                                    sSectionFlag = "F";
                                                                }
                                                                else
                                                                    sSectionFlag = "T";
                                                            }
                                                            else
                                                            {
                                                                sSuccessErrorMessage += Environment.NewLine + "Row no: " + iRowCurrent + " Section " + Section + " does not exist in Sage CRM";
                                                                sSectionFlag = "F";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no: " + iRowCurrent + " Section " + Section + " does not exist in Sage CRM";
                                                            sSectionFlag = "F";
                                                        }
                                                    }
                                                }

                                                sMonday = dt.Rows[i]["Mon"].ToString().Trim();
                                                if (sMonday != "")
                                                {
                                                    if (sMonday == "Y")
                                                        sAllday = sAllday + GetCaptionCode("Monday", "rate_day") + ",";
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Mon Rate"].ToString().Trim()))
                                                {
                                                    sMondayRate = dt.Rows[i]["Mon Rate"].ToString().Trim();
                                                    #region Check for Alphabets
                                                    bool result = decimal.TryParse(sMondayRate, out iMondayRate);
                                                    if (iMondayRate == 0 && sMondayRate != "0")
                                                    {
                                                        if (sMondayRate != "0.00")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no: " + iRowCurrent + " An Error occured during import process Invalid Number:  (Monday Rate).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                    sMondayRate = "0.00";

                                                sTesuday = dt.Rows[i]["Tue"].ToString().Trim();
                                                if (sTesuday == "Y")
                                                    sAllday = sAllday + GetCaptionCode("Tuesday", "rate_day") + ",";

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Tue Rate"].ToString().Trim()))
                                                {
                                                    sTuesdayRate = dt.Rows[i]["Tue Rate"].ToString().Trim();
                                                    #region Check for Alphabets
                                                    bool result = decimal.TryParse(sTuesdayRate, out iTuesdayRate);
                                                    if (iTuesdayRate == 0 && sTuesdayRate != "0")
                                                    {
                                                        if (sTuesdayRate != "0.00")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Tuesday Rate).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                    sTuesdayRate = "0.00";

                                                sWednesday = dt.Rows[i]["Wed"].ToString().Trim();
                                                if (sWednesday == "Y")
                                                    sAllday = sAllday + GetCaptionCode("Wednesday", "rate_day") + ",";

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Wed Rate"].ToString().Trim()))
                                                {
                                                    sWednesdayRate = dt.Rows[i]["Wed Rate"].ToString().Trim();
                                                    #region Check for Alphabets
                                                    bool result = decimal.TryParse(sWednesdayRate, out iWednesdayRate);
                                                    if (iWednesdayRate == 0 && sWednesdayRate != "0")
                                                    {
                                                        if (sWednesdayRate != "0.00")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Wednesday Rate).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                    sWednesdayRate = "0.00";

                                                sThursday = dt.Rows[i]["Thur"].ToString().Trim();
                                                if (sThursday == "Y")
                                                    sAllday = sAllday + GetCaptionCode("Thursday", "rate_day") + ",";

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Thur Rate"].ToString().Trim()))
                                                {
                                                    sThursdayRate = dt.Rows[i]["Thur Rate"].ToString().Trim();
                                                    #region Check for Alphabets
                                                    bool result = decimal.TryParse(sThursdayRate, out iThrusdayRate);
                                                    if (iThrusdayRate == 0 && sThursdayRate != "0")
                                                    {
                                                        if (sThursdayRate != "0.00")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Thrusday Rate).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                    sThursdayRate = "0.00";

                                                sFriday = dt.Rows[i]["Fri"].ToString().Trim();
                                                if (sFriday == "Y")
                                                    sAllday = sAllday + GetCaptionCode("Friday", "rate_day") + ",";

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Fri Rate"].ToString().Trim()))
                                                {
                                                    sFridayRate = dt.Rows[i]["Fri Rate"].ToString().Trim();
                                                    #region Check for Alphabets
                                                    bool result = decimal.TryParse(sFridayRate, out iFridayRate);
                                                    if (iFridayRate == 0 && sFridayRate != "0")
                                                    {
                                                        if (sFridayRate != "0.00")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Friday Rate).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                    sFridayRate = "0.00";

                                                sSaturday = dt.Rows[i]["Sat"].ToString().Trim();
                                                if (sSaturday == "Y")
                                                    sAllday = sAllday + GetCaptionCode("Saturday", "rate_day") + ",";
                                                if (sAllday != "")

                                                    if (!String.IsNullOrEmpty(dt.Rows[i]["Sat Rate"].ToString().Trim()))
                                                    {
                                                        sSaturdayRate = dt.Rows[i]["Sat Rate"].ToString().Trim();
                                                        #region Check for Alphabets
                                                        bool result = decimal.TryParse(sSaturdayRate, out iSaturdayRate);
                                                        if (iSaturdayRate == 0 && sSaturdayRate != "0")
                                                        {
                                                            if (sSaturdayRate != "0.00")
                                                            {
                                                                isValid = "false";
                                                                sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Saturday Rate)).";
                                                                iFailedCount++;
                                                            }
                                                        }
                                                        #endregion
                                                    }
                                                    else
                                                        sSaturdayRate = "0.00";

                                                sSunday = dt.Rows[i]["Sun"].ToString().Trim();
                                                if (sSunday != "")
                                                {
                                                    if (sSunday == "Y")
                                                        sAllday = sAllday + GetCaptionCode("Sunday", "rate_day") + ",";

                                                }
                                                sAllday = sAllday.TrimEnd(',');
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Sun Rate"].ToString().Trim()))
                                                {
                                                    sSundayRate = dt.Rows[i]["Sun Rate"].ToString().Trim();
                                                    #region Check for Alphabets
                                                    bool result = decimal.TryParse(sSundayRate, out iSundayRate);
                                                    if (iSundayRate == 0 && sSundayRate != "0")
                                                    {
                                                        if (sSundayRate != "0.00")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Sunday Rate).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                    sSundayRate = "0.00";

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Class Type"].ToString().Trim()))
                                                {
                                                    sClassType = dt.Rows[i]["Class Type"].ToString().Trim();
                                                    sClassTypeCode = GetCaptionCode(sClassType, "rate_classtype");

                                                    sRateCardDescription += " " + dt.Rows[i]["Class Type"].ToString().Trim();
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Size"].ToString().Trim()))
                                                {
                                                    sSize = dt.Rows[i]["Size"].ToString().Trim();
                                                    sSizeCode = GetCaptionCode(sSize, "rate_size");

                                                    sRateCardDescription += " " + dt.Rows[i]["Size"].ToString().Trim();
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Standard Size"].ToString().Trim()))
                                                {
                                                    sStandardSize = dt.Rows[i]["Standard Size"].ToString().Trim();                                                    
                                                    InsertIntoCaption(sStandardSize, "rate_standardsize");
                                                    sStandardSizeCode = GetCaptionCode(sStandardSize, "rate_standardsize");

                                                    sRateCardDescription += " " + dt.Rows[i]["Standard Size"].ToString().Trim();
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Height"].ToString().Trim()))
                                                {
                                                    sHeight = dt.Rows[i]["Height"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = decimal.TryParse(sHeight, out iheight);
                                                    if (iheight == 0 && sHeight != "0")
                                                    {
                                                        if (sHeight != "0.00")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Height).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width"].ToString().Trim()))
                                                {
                                                    sWidth = dt.Rows[i]["Width"].ToString().Trim();
                                                    #region Check for Alphabets
                                                    bool result = decimal.TryParse(sHeight, out iwidth);
                                                    if (iwidth == 0 && sWidth != "0")
                                                    {
                                                        if (sWidth != "0.00")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Page Colour"].ToString().Trim()))
                                                {
                                                    sPageColor = dt.Rows[i]["Page Colour"].ToString().Trim();
                                                    sPageColorCode = GetCaptionCode(sPageColor, "rate_color");
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Loading"].ToString().Trim()))
                                                {
                                                    sLoading = dt.Rows[i]["Loading"].ToString().Trim();
                                                    #region Check for Alphabets
                                                    bool result = decimal.TryParse(sLoading, out iLoading);
                                                    if (iLoading == 0 && sLoading != "0")
                                                    {
                                                        if (sLoading != "0.00")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Loading).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Max Height"].ToString().Trim()))
                                                {
                                                    sRateMaxHeight = dt.Rows[i]["Max Height"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateMaxHeight, out iRateMaxHeight);
                                                    if (iRateMaxHeight == 0 && sRateMaxHeight != "0")
                                                    {
                                                        if (sRateMaxHeight != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Max Height).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 1"].ToString().Trim()))
                                                {
                                                    sRateWidth1 = dt.Rows[i]["Width 1"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth1, out iRateWidth1);
                                                    if (iRateWidth1 == 0 && sRateWidth1 != "0")
                                                    {
                                                        if (sRateWidth1 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 1).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 2"].ToString().Trim()))
                                                {
                                                    sRateWidth2 = dt.Rows[i]["Width 2"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth2, out iRateWidth2);
                                                    if (iRateWidth2 == 0 && sRateWidth2 != "0")
                                                    {
                                                        if (sRateWidth2 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 2).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 3"].ToString().Trim()))
                                                {
                                                    sRateWidth3 = dt.Rows[i]["Width 3"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth3, out iRateWidth3);
                                                    if (iRateWidth3 == 0 && sRateWidth3 != "0")
                                                    {
                                                        if (sRateWidth3 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 3).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 4"].ToString().Trim()))
                                                {
                                                    sRateWidth4 = dt.Rows[i]["Width 4"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth4, out iRateWidth4);
                                                    if (iRateWidth4 == 0 && sRateWidth4 != "0")
                                                    {
                                                        if (sRateWidth4 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 4).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 5"].ToString().Trim()))
                                                {
                                                    sRateWidth5 = dt.Rows[i]["Width 5"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth5, out iRateWidth5);
                                                    if (iRateWidth5 == 0 && sRateWidth5 != "0")
                                                    {
                                                        if (sRateWidth5 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 5).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 6"].ToString().Trim()))
                                                {
                                                    sRateWidth6 = dt.Rows[i]["Width 6"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth6, out iRateWidth6);
                                                    if (iRateWidth6 == 0 && sRateWidth6 != "0")
                                                    {
                                                        if (sRateWidth6 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 6).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 7"].ToString().Trim()))
                                                {
                                                    sRateWidth7 = dt.Rows[i]["Width 7"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth7, out iRateWidth7);
                                                    if (iRateWidth7 == 0 && sRateWidth7 != "0")
                                                    {
                                                        if (sRateWidth7 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 7).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 8"].ToString().Trim()))
                                                {
                                                    sRateWidth8 = dt.Rows[i]["Width 8"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth8, out iRateWidth8);
                                                    if (iRateWidth8 == 0 && sRateWidth8 != "0")
                                                    {
                                                        if (sRateWidth8 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 8).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 9"].ToString().Trim()))
                                                {
                                                    sRateWidth9 = dt.Rows[i]["Width 9"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth9, out iRateWidth9);
                                                    if (iRateWidth9 == 0 && sRateWidth9 != "0")
                                                    {
                                                        if (sRateWidth9 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 9).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 10"].ToString().Trim()))
                                                {
                                                    sRateWidth10 = dt.Rows[i]["Width 10"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth10, out iRateWidth10);
                                                    if (iRateWidth10 == 0 && sRateWidth10 != "0")
                                                    {
                                                        if (sRateWidth10 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 10).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 11"].ToString().Trim()))
                                                {
                                                    sRateWidth11 = dt.Rows[i]["Width 11"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth11, out iRateWidth11);
                                                    if (iRateWidth11 == 0 && sRateWidth11 != "0")
                                                    {
                                                        if (sRateWidth11 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 11).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Width 12"].ToString().Trim()))
                                                {
                                                    sRateWidth12 = dt.Rows[i]["Width 12"].ToString().Trim();

                                                    #region Check for Alphabets
                                                    bool result = int.TryParse(sRateWidth12, out iRateWidth12);
                                                    if (iRateWidth12 == 0 && sRateWidth12 != "0")
                                                    {
                                                        if (sRateWidth12 != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Width 12).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                //Record objExistingRateCardRec = FindRecord("Ratescard", "rate_PublicationsID=" + sPublicationId);
                                                //if (!objExistingRateCardRec.Eof())
                                                //{
                                                //    AddContent("1");
                                                //    sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " aleady exist in Sage CRM.";
                                                //    isDatavalid = "true";

                                                //    iDupRecord++;
                                                //    iRecordcount++;
                                                //}

                                                if (dt.Rows.Count >= 1)
                                                {
                                                    if (isValid == "true")
                                                    {
                                                        isDatavalid = isColumnValid;

                                                        Record objRateCardRec = FindRecord("Ratescard", "rate_PublicationsID=" + sPublicationId + " and rate_name = '" + sRateCardDescription + "' ");

                                                        if (objRateCardRec.Eof())
                                                        {
                                                            objRateCardRec = new Record("ratescard");
                                                            isValid = "true";
                                                        }
                                                        else
                                                        {
                                                            isValid = "false";
                                                        }

                                                        objRateCardRec.SetField("rate_PublicationsID", sPublicationId);
                                                        objRateCardRec.SetField("rate_commissiontype", sRateCard);
                                                        objRateCardRec.SetField("rate_Section", sSectionId);
                                                        objRateCardRec.SetField("rate_subsectionid", sSubSectionID);
                                                        objRateCardRec.SetField("rate_day", sAllday);
                                                        objRateCardRec.SetField("rate_size", sSizeCode);
                                                        objRateCardRec.SetField("rate_standardsize", sStandardSizeCode);
                                                        objRateCardRec.SetField("rate_name", sRateCardDescription);

                                                        if (sHeight != "")
                                                        {
                                                            if (iheight != 0)
                                                                objRateCardRec.SetField("rate_height", iheight.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_height", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_height", "0.00");

                                                        if (sWidth != "")
                                                        {
                                                            if (iwidth != 0)
                                                                objRateCardRec.SetField("rate_width", iwidth.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_width", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_width", "0.00");

                                                        //'objRateCardRec.SetField("rate_category", sCategoryCode);
                                                        objRateCardRec.SetField("rate_color", sPageColorCode);
                                                        objRateCardRec.SetField("rate_classtype", sClassTypeCode);

                                                        if (sLoading != "")
                                                        {
                                                            if (iLoading != 0)
                                                                objRateCardRec.SetField("rate_loading", iLoading.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_loading", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_loading", "0.00");

                                                        if (sMondayRate != "")
                                                        {
                                                            if (iMondayRate != 0)
                                                                objRateCardRec.SetField("rate_Monday", iMondayRate.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_Monday", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_Monday", "0.00");

                                                        if (sTuesdayRate != "")
                                                        {
                                                            if (iTuesdayRate != 0)
                                                                objRateCardRec.SetField("rate_tuesday", iTuesdayRate.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_tuesday", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_tuesday", "0.00");

                                                        if (sWednesdayRate != "")
                                                        {
                                                            if (iWednesdayRate != 0)
                                                                objRateCardRec.SetField("rate_wednesday", iWednesdayRate.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_wednesday", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_wednesday", "0.00");

                                                        if (sThursdayRate != "")
                                                        {
                                                            if (iThrusdayRate != 0)
                                                                objRateCardRec.SetField("rate_thrusday", iThrusdayRate.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_thrusday", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_thrusday", "0.00");

                                                        if (sFridayRate != "")
                                                        {
                                                            if (iFridayRate != 0)
                                                                objRateCardRec.SetField("rate_friday", iFridayRate.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_friday", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_friday", "0.00");

                                                        if (sSaturdayRate != "")
                                                        {
                                                            if (iSaturdayRate != 0)
                                                                objRateCardRec.SetField("rate_saturday", iSaturdayRate.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_saturday", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_saturday", "0.00");

                                                        if (sSundayRate != "")
                                                        {
                                                            if (iSundayRate != 0)
                                                                objRateCardRec.SetField("rate_sunday", iSundayRate.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_sunday", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_sunday", "0.00");

                                                        

                                                        objRateCardRec.SetField("rate_Monday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_tuesday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_wednesday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_thrusday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_friday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_saturday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_sunday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_price_CID", sBaseCurrency);

                                                        objRateCardRec.SetField("rate_maxheight", iRateMaxHeight);
                                                        objRateCardRec.SetField("rate_width1", iRateWidth1);
                                                        objRateCardRec.SetField("rate_width2", iRateWidth2);
                                                        objRateCardRec.SetField("rate_width3", iRateWidth3);
                                                        objRateCardRec.SetField("rate_width4", iRateWidth4);
                                                        objRateCardRec.SetField("rate_width5", iRateWidth5);
                                                        objRateCardRec.SetField("rate_width6", iRateWidth6);
                                                        objRateCardRec.SetField("rate_width7", iRateWidth7);
                                                        objRateCardRec.SetField("rate_width8", iRateWidth8);
                                                        objRateCardRec.SetField("rate_width9", iRateWidth9);
                                                        objRateCardRec.SetField("rate_width10",iRateWidth10);
                                                        objRateCardRec.SetField("rate_width11", iRateWidth11);
                                                        objRateCardRec.SetField("rate_width12", iRateWidth12);

                                                        objRateCardRec.SaveChanges();

                                                        if (isValid == "true")
                                                        {
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " successfully imported in CRM.";
                                                        }
                                                        else if (isValid == "false")
                                                        {
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " successfully updated in CRM.";
                                                        }

                                                        InsertCount++;
                                                    }
                                                    else if (dt.Rows.Count == 1)
                                                    {
                                                        GeneratelogFile("[" + System.DateTime.Now.ToString() + "] " + sSuccessErrorMessage);
                                                        string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                                                        sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&hasrows=Y&dotnetfunc=RunRateCardImportStatusPage&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName;
                                                        Dispatch.Redirect(sURL);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " Publication " + sPublication + " does not exist in Sage CRM.";
                                                isDatavalid = "true";                                             
                                                iRecordcount++;
                                            }
                                        }
                                        else
                                        {
                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " Publication column is empty in Excel Sheet.";
                                            //GeneratelogFile("[" + System.DateTime.Now.ToString() + "] " + sSuccessErrorMessage);                                            
                                        }
                                    }
                                    catch (Exception Ex)
                                    {
                                        isColumnValid = "false";
                                        AddContent(Ex.Message.ToString());
                                        string strURL = UrlDotNet(this.ThisDotNetDll, "RunRateCardImportStatusPage");
                                        strURL += "&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName + "&ValidColumn=ROW";
                                        Dispatch.Redirect(strURL);
                                        return;
                                    }
                                }
                            }
                            else
                            {
                                GeneratelogFile("[" + System.DateTime.Now.ToString() + "]" + " No records found in the excel sheet.");
                                string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                                sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&hasrows=N&dotnetfunc=RunRateCardImportStatusPage&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName;
                                Dispatch.Redirect(sURL);
                            }
                        }
                        #endregion
                        
                        if (isDatavalid == "true" && dt.Rows.Count > 0 && iDupRecord != dt.Rows.Count)
                        {
                            RefreshMetata();
                            GeneratelogFile("[" + System.DateTime.Now.ToString() + "] " + sSuccessErrorMessage);
                            string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                            sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&hasrows=Y&dotnetfunc=RunRateCardImportStatusPage&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName + "&AllDup=N";                            
                            Dispatch.Redirect(sURL);
                        }
                        if (iDupRecord == dt.Rows.Count && isDatavalid == "true")
                        {
                            GeneratelogFile("[" + System.DateTime.Now.ToString() + "] " + sSuccessErrorMessage);
                            string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                            sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&hasrows=Y&dotnetfunc=RunRateCardImportStatusPage&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName + "&AllDup=E";
                            Dispatch.Redirect(sURL);
                        }
                    }
                }
            }
            catch (Exception Ex)
            {                
                this.AddError(Ex.Message);
            }
            //'base.BuildContents();
        }

        private void InsertIntoCaption(string sStandardSize, string p)
        {
            string sCode = "";
            Record objCustomCaption = FindRecord("Custom_captions", " LOWER(LTRIM(RTRIM(capt_family)))='" + p.ToLower().Trim() + "' and LOWER(LTRIM(RTRIM(capt_us)))='" + sStandardSize.ToLower().Trim() + "'");
            sCode = sStandardSize.Replace(" ", "");
            if (!objCustomCaption.Eof())
            {
                objCustomCaption.SetField("capt_code", sCode);
                objCustomCaption.SetField("capt_US", sStandardSize);

                objCustomCaption.SaveChanges();
            }
            else
            {
                Record objCustomCaptionInsert = new Record("Custom_captions");

                objCustomCaptionInsert.SetField("capt_US", sStandardSize);
                objCustomCaptionInsert.SetField("capt_code", sCode);
                objCustomCaptionInsert.SetField("capt_family", p);

                objCustomCaptionInsert.SaveChanges();
            }

            
        }
        #region Save File in  Rate Card folder
        private string SaveRateCardRLocation()
        {
            string FileName = "";
            string newFullPath = "";
            string UploadfilePath = Dispatch.ContentField("HIDDEN_FilePath");
            string LibPath = GetLibraryPath();
            string NewPath = LibPath.Replace("\\Library", "");
            FileName = Dispatch.ContentField("HIDDEN_FileName");
            NewPath += "\\WWWRoot\\CustomPages\\NZPAImport\\ImportedFiles\\RateCard\\";
            if (Directory.Exists(NewPath))
            {
                //FileUpload1.PostedFile.FileName

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

        #region Get Library Path
        public string GetLibraryPath()
        {
            string Path = "";
            Record RecPath = FindRecord("Custom_SysParams", "parm_name = 'DocStore'");
            Path = RecPath.GetFieldAsString("Parm_Value");
            return Path;
        }
        #endregion

        #region Get Caption Code
        public string GetCaptionCode(string sCaption, string sFamily)
        {
            string sCode = "";
            Record objCustomCaption = FindRecord("Custom_captions", "LOWER(LTRIM(RTRIM(capt_us)))='" + sCaption.ToLower().Trim() + "' and LOWER(LTRIM(RTRIM(capt_family)))='" + sFamily.ToLower().Trim() + "'");

            if (!objCustomCaption.Eof())
            {
                sCode = objCustomCaption.GetFieldAsString("capt_code");
            }

            return sCode;
        }
        #endregion

        #region Get Base Currency
        public string GetBaseCurrency()
        {
            string sBaseCurrency = "";
            string strSQL = "select parm_value from Custom_SysParams  where Parm_Name ='BaseCurrency' and Parm_Deleted is null";
            QuerySelect objCurrencyRec = GetQuery();
            objCurrencyRec.SQLCommand = strSQL;
            objCurrencyRec.ExecuteReader();

            if (!objCurrencyRec.Eof())
            {
                sBaseCurrency = objCurrencyRec.FieldValue("parm_value");
            }
            return sBaseCurrency;
        }
        #endregion
        #endregion

        public void GeneratelogFile(string Logcontent)
        {
            string LibPath = GetLibraryPath();

            string NewPath = LibPath.Replace("\\Library", "");
            NewPath += "\\WWWRoot\\CustomPages\\NZPAImport\\";
            string sInstallDirName = new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).DirectoryName;
            string Logspath = null;
            try
            {
                string currentPath = NewPath;
                if (!Directory.Exists(Path.Combine(currentPath, "LogFiles")))
                    Directory.CreateDirectory(Path.Combine(currentPath, "LogFiles"));

                DateTime theDate = DateTime.Now;
                string ymd = System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString() + theDate.ToString("yyyyMMdd") + "ImportRateCard.txt";
                LogfileName = ymd;
                Logspath = NewPath + "\\LogFiles\\" + ymd;
                
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
    }
}
