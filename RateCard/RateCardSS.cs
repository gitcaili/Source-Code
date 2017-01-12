using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;
using System.Data;
using System.IO;
using System.Diagnostics;

namespace NZPACRM.RateCard
{
    public class RateCardSS : DataPageNew

    {
        Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        Microsoft.Office.Interop.Excel._Worksheet mWSheet1;
        Microsoft.Office.Interop.Excel.Application oXL;
        CRMHelper objCRM = new CRMHelper();
        string LogfileName = "";
        string shttpURL = "";
        public RateCardSS()
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
                AddContent("<script type='text/javascript' src='../CustomPages/Booking/ClientFuncsJune.js'></script>");

                int ImportCount = 0;
                string EntityID = "0";
                string sPublication = "";
                string sRateCard = "";
                string Section = "";
                //  string sAgentcy = "";
                string sClient = "";
                string sSubsection = "";
                string sSectionCode = "";
                string sCategory = "";
                string sCategoryCode = "";
                string sPageColor = "";
                // string sPageColorCode = "";
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
                string sSizeWidth = "";
                string sSizeHeight = "";
                string sBookDeadTime = "";
                string sBookDeadDay = "";
                string sMaterialTime = "";
                string sMaterialDays = "";
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
                string sCusSheet = "";
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
                AddContent("<BR><BR>" + HTML.Box("File", "<br>&nbsp;&nbsp;<input type='file' id='fileupload' name='pic' size='70'>&nbsp;<input type='BUTTON' class='Edit' value='Import'name='upload' onclick='javascript:CheckFileNew();'></br></br>"));
                #endregion

                #region Add Buttons

                string backURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                backURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=1650&Mode=1&CLk=T&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data Management";
                AddUrlButton("Back", "prevcircle.gif", backURL);
                #endregion

                #region Define the Hidden Fields
                AddContent(HTML.InputHidden("HIDDEN_LibraryPath", GetLibraryPath().Replace("Library", "")));
                AddContent(HTML.InputHidden("HIDDEN_FilePath", ""));
                AddContent(HTML.InputHidden("HIDDEN_Save", ""));
                AddContent(HTML.InputHidden("HIDDEN_FileName", ""));
                AddContent(HTML.InputHidden("HIDDEN_FilePathChrome", ""));
                AddContent(HTML.InputHidden("HIDDEN_browser", ""));
                #endregion

                if (!String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_Save")))
                {
                    if (Dispatch.ContentField("HIDDEN_Save") == "Save")
                    {
                        AddContent(Dispatch.ContentField("HIDDEN_FilePath"));
                        // DataSet ds = new DataSet();
                        // DataTable dt = new DataTable();
                        if (!String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_FileName")))
                        {

                            #region Save File in Rate Card folder
                            if (String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_FilePathChrome")))
                            {
                                SavedFilePath = SaveRateCardRLocation();
                            }
                            else
                            {
                                SavedFilePath = GetLibraryPath() + @"\RateCard\" + Dispatch.ContentField("HIDDEN_FileName");
                            }
                            #endregion
                            oXL = new Microsoft.Office.Interop.Excel.Application();
                            oXL.DisplayAlerts = false;
                            #region Read Excel File data
                            string extention = Path.GetExtension(SavedFilePath);
                            //   dt = objCRM.ConvertToDataTable(SavedFilePath);

                            mWorkBook = oXL.Workbooks.Open(SavedFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                            var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; data source=" + SavedFilePath + "; Extended Properties='Excel 12.0;HDR=NO;IMEX=1;'", SavedFilePath);
                            mWorkSheets = mWorkBook.Worksheets;
                            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item(1);
                            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;

                            #region Find EntityID
                            Record RecEntityID = FindRecord("Custom_Tables", "bord_name='ratescard' and bord_deleted is null");
                            if (!RecEntityID.Eof())
                            {
                                EntityID = RecEntityID.GetFieldAsString("Bord_TableId");
                            }
                            #endregion

                            sBaseCurrency = GetBaseCurrency();
                            //     AddContent(mWSheet1.UsedRange.Rows.Count.ToString() + "HEARTS ");
                            if (mWSheet1.UsedRange.Rows.Count > 0)
                            {

                                for (int i = 2; i < mWSheet1.UsedRange.Rows.Count + 1; i++)
                                {
                                    //     AddContent(mWSheet1.UsedRange.Rows.Count.ToString());
                                    //   AddContent((i + 2).ToString() + " SSLINE");
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
                                    //   AddContent(" " + i.ToString());
                                    try
                                    {
                                        if (String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Text).Trim())) break;
                                        if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))

                                        {
                                            AddContent("QUEEN " + i);
                                            sRateCardDescription = ((string)(mWSheet1.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value).Trim();
                                            //  AddContent("START " + sRateCardDescription);
                                            sPublication = ((string)(mWSheet1.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value).Trim().Replace("'", "''");
                                            Record objPublication = FindRecord("publicationS", "LOWER(LTRIM(RTRIM(pblc_Name)))='" + sPublication.ToLower() + "'");

                                            if (!objPublication.Eof())
                                            {
                                                sPublicationId = objPublication.GetFieldAsString("pblc_PublicationsID");

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value).Trim()))
                                                {
                                                    sRateCard = ((string)(mWSheet1.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value).Trim();
                                                    sRateCard = GetCaptionCode(sRateCard, "pnbr_commissiontype");
                                                    sRateCardDescription += " " + ((string)(mWSheet1.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value).Trim();
                                                }
                                                else
                                                    sRateCard = "";

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value).Trim()))

                                                {
                                                    Section = ((string)(mWSheet1.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value).Trim();

                                                    Record objSections = FindRecord("sections", "sctn_name='" + Section + "' and sctn_publicationID='" + sPublicationId + "'");
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
                                                    sRateCardDescription += " " + ((string)(mWSheet1.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value).Trim();
                                                }
                                                AddContent("CATT");
                                                //  AddContent((mWSheet1.Cells[1, 3] as Microsoft.Office.Interop.Excel.Range).Value2.ToString());
                                                //  AddContent((mWSheet1.Cells[2, 4] as Microsoft.Office.Interop.Excel.Range).Text.ToString());
                                                //  AddContent("BAT");

                                                if (!string.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                //  Microsoft.Office.Interop.Excel.Range r = mWSheet1.Cells[i, 4];

                                                {
                                                    //   AddContent("BAT");
                                                    sSubsection = ((string)(mWSheet1.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value).Trim();
                                                    //    AddContent("DOGGG " + sSubsection);
                                                    Record objSubSectionRec = FindRecord("SubSection", "suse_name='" + sSubsection + "'and suse_section = '" + sSectionId + "'");
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
                                                AddContent("RAT");
                                                sMonday = ((string)(mWSheet1.Cells[i, 6] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                //   AddContent("RATs");
                                                if (sMonday != "")
                                                {
                                                    //    AddContent("RATsfedcfmwep");
                                                    if (sMonday == "Y")
                                                        sAllday = sAllday + GetCaptionCode("Monday", "rate_day") + ",";


                                                    if ((mWSheet1.Cells[i, 7] as Microsoft.Office.Interop.Excel.Range).Value2 != null && sMonday == "Y")
                                                    {

                                                        sMondayRate = ((mWSheet1.Cells[i, 7] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                                                        //    AddContent(sMondayRate);
                                                        #region Check for Alphabets
                                                        bool result = decimal.TryParse(sMondayRate, out iMondayRate);
                                                        if (Math.Round(Convert.ToDouble(iMondayRate), 2) == 0 && Convert.ToString(Math.Round(Convert.ToDouble(sMondayRate), 2)) != "0")
                                                        {
                                                            if (Convert.ToString(Math.Round(Convert.ToDouble(sMondayRate), 2)) != "0.00")
                                                            {
                                                                isValid = "false";
                                                                //  AddContent("here - Mon");
                                                                sSuccessErrorMessage += Environment.NewLine + "Row no: " + iRowCurrent + " An Error occured during import process Invalid Number:  (Monday Rate).";
                                                                iFailedCount++;
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                    sMondayRate = "0.00";
                                                //   AddContent("mon");
                                                sTesuday = ((string)(mWSheet1.Cells[i, 8] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                if (sTesuday == "Y")
                                                {
                                                    sAllday = sAllday + GetCaptionCode("Tuesday", "rate_day") + ",";

                                                    if ((mWSheet1.Cells[i, 9] as Microsoft.Office.Interop.Excel.Range).Value2 != null)
                                                    {
                                                        sTuesdayRate = ((mWSheet1.Cells[i, 9] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                                                        #region Check for Alphabets
                                                        bool result = decimal.TryParse(sTuesdayRate, out iTuesdayRate);
                                                        //AddContent("sTuesdayRate" + Convert.ToString(Math.Round(Convert.ToDouble(sTuesdayRate), 2)));
                                                        //AddContent("iTuesdayRate" + Convert.ToString(Math.Round(Convert.ToDouble(iTuesdayRate), 2)));
                                                        if (Math.Round(Convert.ToDouble(iTuesdayRate), 2) == 0 && Convert.ToString(Math.Round(Convert.ToDouble(sTuesdayRate), 2)) != "0")
                                                        {
                                                            if (Convert.ToString(Math.Round(Convert.ToDouble(sTuesdayRate), 2)) != "0.00")
                                                            {
                                                                isValid = "false";
                                                                //      AddContent("Tue - ");
                                                                sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Tuesday Rate). iTuesdayRate=" + iTuesdayRate + ",sTuesdayRate=" + sTuesdayRate + "";
                                                                iFailedCount++;
                                                            }
                                                        }
                                                        #endregion
                                                    }
                                                }
                                                else
                                                    sTuesdayRate = "0.00";
                                                AddContent("FREAK");
                                                sWednesday = ((string)(mWSheet1.Cells[i, 10] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                if (sWednesday == "Y")
                                                {
                                                    sAllday = sAllday + GetCaptionCode("Wednesday", "rate_day") + ",";

                                                    if ((mWSheet1.Cells[i, 11] as Microsoft.Office.Interop.Excel.Range).Value2 != null)
                                                    {
                                                        sWednesdayRate = ((mWSheet1.Cells[i, 11] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                                                        #region Check for Alphabets
                                                        bool result = decimal.TryParse(sWednesdayRate, out iWednesdayRate);
                                                        if (Math.Round(Convert.ToDouble(iWednesdayRate), 2) == 0 && Convert.ToString(Math.Round(Convert.ToDouble(sWednesdayRate), 2)) != "0")
                                                        {
                                                            if (Convert.ToString(Math.Round(Convert.ToDouble(sWednesdayRate), 2)) != "0.00")
                                                            {
                                                                isValid = "false";
                                                                //    AddContent("wed - ");
                                                                sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Wednesday Rate).";
                                                                iFailedCount++;
                                                            }
                                                        }
                                                        #endregion
                                                    }
                                                }
                                                else
                                                    sWednesdayRate = "0.00";
                                                //    AddContent("BAT");
                                                sThursday = ((string)(mWSheet1.Cells[i, 12] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                if (sThursday == "Y")
                                                {
                                                    sAllday = sAllday + GetCaptionCode("Thursday", "rate_day") + ",";

                                                    if ((mWSheet1.Cells[i, 13] as Microsoft.Office.Interop.Excel.Range).Value2 != null)
                                                    {
                                                        sThursdayRate = ((mWSheet1.Cells[i, 13] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                                                        #region Check for Alphabets
                                                        bool result = decimal.TryParse(sThursdayRate, out iThrusdayRate);
                                                        if (Math.Round(Convert.ToDouble(iThrusdayRate), 2) == 0 && Convert.ToString(Math.Round(Convert.ToDouble(sThursdayRate), 2)) != "0")
                                                        {
                                                            if (Convert.ToString(Math.Round(Convert.ToDouble(sThursdayRate), 2)) != "0.00")
                                                            {
                                                                isValid = "false";
                                                                //   AddContent("thu - ");
                                                                sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Thrusday Rate).";
                                                                iFailedCount++;
                                                            }
                                                        }
                                                        #endregion
                                                    }
                                                }
                                                else
                                                    sThursdayRate = "0.00";
                                                AddContent("BATT");
                                                sFriday = ((string)(mWSheet1.Cells[i, 14] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                if (sFriday == "Y")
                                                {
                                                    //    AddContent("fricfmqsdkljvgncm");
                                                    sAllday = sAllday + GetCaptionCode("Friday", "rate_day") + ",";

                                                    if ((mWSheet1.Cells[i, 15] as Microsoft.Office.Interop.Excel.Range).Value2 != null)
                                                    {
                                                        //  AddContent("friRATEvcsdnwm");
                                                        sFridayRate = ((mWSheet1.Cells[i, 15] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                                                        #region Check for Alphabets
                                                        bool result = decimal.TryParse(sFridayRate, out iFridayRate);
                                                        if (Math.Round(Convert.ToDouble(iFridayRate), 2) == 0 && Convert.ToString(Math.Round(Convert.ToDouble(sFridayRate), 2)) != "0")
                                                        {
                                                            if (Convert.ToString(Math.Round(Convert.ToDouble(sFridayRate), 2)) != "0.00")
                                                            {
                                                                isValid = "false";
                                                                //     AddContent("here - ");
                                                                sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Friday Rate).";
                                                                iFailedCount++;
                                                            }
                                                        }
                                                        #endregion
                                                    }
                                                }
                                                else
                                                    sFridayRate = "0.00";
                                                AddContent("GFUIDNHSIPVFUON");
                                                //AddContent((mWSheet1.Cells[i, 17] as Microsoft.Office.Interop.Excel.Range).Value2.ToString());
                                                sSaturday = ((string)(mWSheet1.Cells[i, 16] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                //  AddContent("GFUIDNHSIPVFUON2");
                                                if (sSaturday == "Y")
                                                {
                                                    //    AddContent("satcdscbhwduibfweuidnhfcio");
                                                    sAllday = sAllday + GetCaptionCode("Saturday", "rate_day") + ",";
                                                }
                                                if (sAllday != "" && sSaturday == "Y")
                                                {
                                                    //   AddContent("SCREWDujbhwduiopcnweuipcbnwduip");
                                                    //   if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 17] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                    if ((mWSheet1.Cells[i, 17] as Microsoft.Office.Interop.Excel.Range).Value2 != null)
                                                    {
                                                        sSaturdayRate = ((mWSheet1.Cells[i, 17] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                                                        #region Check for Alphabets
                                                        bool result = decimal.TryParse(sSaturdayRate, out iSaturdayRate);
                                                        if (Math.Round(Convert.ToDouble(iSaturdayRate), 2) == 0 && Convert.ToString(Math.Round(Convert.ToDouble(sSaturdayRate), 2)) != "0")
                                                        {
                                                            if (Convert.ToString(Math.Round(Convert.ToDouble(sSaturdayRate), 2)) != "0.00")
                                                            {
                                                                isValid = "false";
                                                                sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Saturday Rate)).";
                                                                iFailedCount++;
                                                            }
                                                        }
                                                        #endregion
                                                    }
                                                }

                                                else
                                                {
                                                    sSaturdayRate = "0.00";
                                                    // AddContent("GFUIDNHSIPVFUON");
                                                }
                                                AddContent("FIRE");
                                                sSunday = ((string)(mWSheet1.Cells[i, 18] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                AddContent("FIRE2");
                                                if (sSunday != "")
                                                {
                                                    if (sSunday == "Y")
                                                        sAllday = sAllday + GetCaptionCode("Sunday", "rate_day") + ",";

                                                }
                                                AddContent("FIRE3");
                                                sAllday = sAllday.TrimEnd(',');

                                                if ((mWSheet1.Cells[i, 19] as Microsoft.Office.Interop.Excel.Range).Value2 != null)
                                                {
                                                    sSundayRate = ((mWSheet1.Cells[i, 19] as Microsoft.Office.Interop.Excel.Range).Value2).ToString();
                                                    #region Check for Alphabets
                                                    bool result = decimal.TryParse(sSundayRate, out iSundayRate);
                                                    // double goat = Convert.ToDouble(sSundayRate);
                                                    //sSundayRate = sSundayRate.Substring(1);
                                                    //    AddContent(sSundayRate);
                                                    if (Math.Round(Convert.ToDouble(iSundayRate), 2) == 0 && Convert.ToString(Math.Round(Convert.ToDouble(sSundayRate), 2)) != "0")
                                                    {
                                                        // AddContent("ALIVE2");
                                                        if (Convert.ToString(Math.Round(Convert.ToDouble(sSundayRate), 2)) != "0.00")
                                                        {
                                                            // isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Sunday Rate).";
                                                            iFailedCount++;
                                                        }
                                                    }
                                                    //  AddContent("ALIVE " + sSundayRate + "Next");
                                                    #endregion
                                                }
                                                else
                                                    sSundayRate = "0.00";
                                                AddContent("WATER");
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 32] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sClassType = ((string)(mWSheet1.Cells[i, 32] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                    //sClassTypeCode = GetCaptionCode(sClassType, "rate_classtype");

                                                    sRateCardDescription += " " + ((string)(mWSheet1.Cells[i, 32] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                }


                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 20] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sCusSheet = ((string)(mWSheet1.Cells[i, 20] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                    //sClassTypeCode = GetCaptionCode(sClassType, "rate_classtype");

                                                    sRateCardDescription += " " + ((string)(mWSheet1.Cells[i, 20] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                }

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 21] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sSize = ((string)(mWSheet1.Cells[i, 21] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                    sSizeCode = GetCaptionCode(sSize, "rate_size");

                                                    sRateCardDescription += " " + ((string)(mWSheet1.Cells[i, 21] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                }

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 22] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sStandardSize = ((string)(mWSheet1.Cells[i, 22] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                    InsertIntoCaption(sStandardSize, "rate_standardsize");
                                                    sStandardSizeCode = GetCaptionCode(sStandardSize, "rate_standardsize");
                                                    //      AddContent(sStandardSize);
                                                    sRateCardDescription += " " + ((string)(mWSheet1.Cells[i, 22] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 23] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sSizeHeight = ((string)(mWSheet1.Cells[i, 23] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                }

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 24] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sSizeWidth = ((string)(mWSheet1.Cells[i, 24] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 25] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sBookDeadTime = ((string)(mWSheet1.Cells[i, 25] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 26] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sBookDeadDay = ((string)(mWSheet1.Cells[i, 26] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 27] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sMaterialTime = ((string)(mWSheet1.Cells[i, 27] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 28] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sMaterialDays = ((string)(mWSheet1.Cells[i, 28] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                }

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    Record agen = FindRecord("Client", "client_name = '" + ((string)(mWSheet1.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Text).Trim() + "'");
                                                    if (!agen.Eof())
                                                    {
                                                        sClient = agen.GetFieldAsString("client_clientID");
                                                    }
                                                    else
                                                    {
                                                        throw new Exception("Client dows not exisit");
                                                    }

                                                    //sClient = dt.Rows[i]["Client"].ToString().Trim();

                                                }

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 29] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sHeight = ((string)(mWSheet1.Cells[i, 29] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sHeight, out parsedValue))
                                                    {
                                                        //AddContent("Height");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                    sHeight = "0";

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 30] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sWidth = ((string)(mWSheet1.Cells[i, 30] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sWidth, out parsedValue))
                                                    {
                                                        //AddContent("Width");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                    sWidth = "0";

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 31] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sPageColor = ((string)(mWSheet1.Cells[i, 31] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                    // sPageColorCode = GetCaptionCode(sPageColor, "rate_color");
                                                }
                                                if (sPageColor == "Colour" || sPageColor == "COLOUR") sPageColor = "Color";
                                                if (sPageColor == "Mono" || sPageColor == "MONO") sPageColor = "NoColor";

                                                //  AddContent("OUT ");
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 33] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {

                                                    sLoading = ((string)(mWSheet1.Cells[i, 33] as Microsoft.Office.Interop.Excel.Range).Text).Trim();
                                                    //   AddContent(sLoading);
                                                    #region Check for Alphabets
                                                    bool result = decimal.TryParse(sLoading, out iLoading);

                                                    if (Math.Round(Convert.ToDouble(iLoading), 2) == 0 && Convert.ToString(Math.Round(Convert.ToDouble(sLoading), 2)) != "0")
                                                    {
                                                        //    AddContent("IN");
                                                        if (Convert.ToString(Math.Round(Convert.ToDouble(sLoading), 2)) != "0")
                                                        {
                                                            isValid = "false";
                                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " An Error occured during import process Invalid Number:  (Loading).";
                                                            iFailedCount++;
                                                        }
                                                    }

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

                                                //  AddContent("GOAT" +dt.Rows[i]["Max Height"].ToString());
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 34] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    //  AddContent("in");
                                                    sRateMaxHeight = ((string)(mWSheet1.Cells[i, 34] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateMaxHeight, out parsedValue))
                                                    {
                                                        //AddContent("Max height");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                    sRateMaxHeight = "0";

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 35] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth1 = ((string)(mWSheet1.Cells[i, 35] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth1, out parsedValue))
                                                    {
                                                        //AddContent("width 1");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth1 = "0";
                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 36] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth2 = ((string)(mWSheet1.Cells[i, 36] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth2, out parsedValue))
                                                    {
                                                        //AddContent("width 2");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth2 = "0";
                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 37] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth3 = ((string)(mWSheet1.Cells[i, 37] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth3, out parsedValue))
                                                    {
                                                        //AddContent("width 3");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth3 = "0";
                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 38] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth4 = ((string)(mWSheet1.Cells[i, 38] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth4, out parsedValue))
                                                    {
                                                        //AddContent("width 4");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth4 = "0";
                                                }

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 39] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth5 = ((string)(mWSheet1.Cells[i, 39] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth5, out parsedValue))
                                                    {
                                                        //AddContent("width 5");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth5 = "0";
                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 40] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth6 = ((string)(mWSheet1.Cells[i, 40] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth6, out parsedValue))
                                                    {
                                                        //AddContent("width 6");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth6 = "0";
                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 41] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth7 = ((string)(mWSheet1.Cells[i, 41] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth7, out parsedValue))
                                                    {
                                                        //AddContent("width 7");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth7 = "0";
                                                }

                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 42] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth8 = ((string)(mWSheet1.Cells[i, 42] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth8, out parsedValue))
                                                    {
                                                        //AddContent("width 8");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth8 = "0";
                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 43] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth9 = ((string)(mWSheet1.Cells[i, 43] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth9, out parsedValue))
                                                    {
                                                        //AddContent("width 9");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth9 = "0";
                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 44] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth10 = ((string)(mWSheet1.Cells[i, 44] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth10, out parsedValue))
                                                    {
                                                        //AddContent("width 10");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth10 = "0";
                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 45] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth11 = ((string)(mWSheet1.Cells[i, 45] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth11, out parsedValue))
                                                    {
                                                        //AddContent("width 11");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth11 = "0";
                                                }
                                                if (!String.IsNullOrEmpty(((string)(mWSheet1.Cells[i, 46] as Microsoft.Office.Interop.Excel.Range).Text).Trim()))
                                                {
                                                    sRateWidth12 = ((string)(mWSheet1.Cells[i, 46] as Microsoft.Office.Interop.Excel.Range).Text).Trim();

                                                    decimal parsedValue;
                                                    if (!decimal.TryParse(sRateWidth12, out parsedValue))
                                                    {
                                                        //AddContent("width 12");
                                                        isValid = "false";
                                                    }
                                                }
                                                else
                                                {
                                                    sRateWidth12 = "0";
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

                                                // AddContent("Count - " + dt.Rows.Count.ToString());
                                                //  AddContent("BOSS" + isValid);
                                                if (mWSheet1.UsedRange.Rows.Count >= 1)
                                                {
                                                    //    AddContent("BOSSGood");
                                                    //  AddContent(isValid);
                                                    if (isValid == "true")
                                                    {
                                                        isDatavalid = isColumnValid;
                                                        //   AddContent("sPublicationId" + sPublicationId);
                                                        //    AddContent("sRateCardDescription" + sRateCardDescription);
                                                        Record objRateCardRec = null;
                                                        // AddContent("FILL" + sPageColorCode);
                                                        if (sStandardSize.Equals("Custom") || 1 == 1)
                                                        {
                                                            string where = ("rate_PublicationsID=" + sPublicationId + "and rate_section =  '" + sSectionId + "' and rate_name = '" + sRateCardDescription + "'and rate_maxheight = '" + sRateMaxHeight + "' and rate_color ='" + sPageColor + "' and rate_customsheet = '" + sCusSheet + "'and rate_customtype = '" + sClassType + "' and rate_BookingDeadlineTime = '" + sBookDeadTime + "' and rate_MaterialDeadlinDays = '" + sMaterialDays + "' and rate_MaterialDeadlineTime = '" + sMaterialTime + "' and rate_BookinDeadlineDays = '" + sBookDeadDay + "' and rate_day = '" + sAllday + "'");

                                                            if (sClient != "")
                                                            {
                                                                where += "and rate_client = '" + sClient + "'";
                                                            }
                                                            if (sSubSectionID != "")
                                                            {
                                                                where += "and rate_subsectionid = '" + sSubSectionID + "'";
                                                            }
                                                            objRateCardRec = FindRecord("Ratescard", where);
                                                            //AddContent("READ rate_PublicationsID = " + sPublicationId + " and rate_name = '" + sRateCardDescription + "'and rate_maxheight = '" + sRateMaxHeight + "' and rate_color = '" + sPageColor + "' and rate_customsheet = '"+ sClassType + "'and rate_customtype = '"+ sCusSheet + "'");
                                                        }
                                                        //else
                                                        //{
                                                        //    objRateCardRec = FindRecord("Ratescard", "rate_PublicationsID=" + sPublicationId + " and rate_name = '" + sRateCardDescription + "'");
                                                        //}

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
                                                        objRateCardRec.SetField("rate_SetSizesHeight", sSizeHeight);
                                                        objRateCardRec.SetField("rate_SetSizesWidth", sSizeWidth);
                                                        objRateCardRec.SetField("rate_BookingDeadlineTime", sBookDeadTime);
                                                        objRateCardRec.SetField("rate_BookinDeadlineDays", sBookDeadDay);
                                                        objRateCardRec.SetField("rate_MaterialDeadlineTime", sMaterialTime);
                                                        objRateCardRec.SetField("rate_MaterialDeadlinDays", sMaterialDays);

                                                        if (sMonday == "Y")
                                                            objRateCardRec.SetField("rate_MonRequired", sMonday);
                                                        if (sTesuday == "Y")
                                                            objRateCardRec.SetField("rate_TueRequired", sTesuday);
                                                        if (sWednesday == "Y")
                                                            objRateCardRec.SetField("rate_WedRequired", sWednesday);
                                                        if (sThursday == "Y")
                                                            objRateCardRec.SetField("rate_ThurRequired", sThursday);
                                                        if (sFriday == "Y")
                                                            objRateCardRec.SetField("rate_FriRequired", sFriday);
                                                        if (sSaturday == "Y")
                                                            objRateCardRec.SetField("rate_SatRequired", sSaturday);
                                                        if (sSunday == "Y")
                                                            objRateCardRec.SetField("rate_SunRequired", sSunday);
                                                        objRateCardRec.SetField("rate_client", sClient);

                                                        //AddContent("ROLL");
                                                        if (sHeight != "")
                                                        {
                                                            if (sHeight != "0")
                                                                objRateCardRec.SetField("rate_height", sHeight.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_height", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_height", "0.00");

                                                        if (sWidth != "")
                                                        {
                                                            if (sWidth != "0")
                                                                objRateCardRec.SetField("rate_width", sWidth.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_width", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_width", "0.00");

                                                        //'objRateCardRec.SetField("rate_category", sCategoryCode);
                                                        objRateCardRec.SetField("rate_color", sPageColor);
                                                        //  objRateCardRec.SetField("rate_classtype", sClassTypeCode);
                                                        //objRateCardRec.SetField("rate_classtype", sClassTypeCode);

                                                        if (sLoading != "")
                                                        {
                                                            if (iLoading != 0)
                                                                objRateCardRec.SetField("rate_loading", iLoading.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_loading", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_loading", "0.00");

                                                        //     AddContent("ABOUT TO ADD" + sMondayRate);
                                                        if (sMondayRate != "")
                                                        {
                                                            if (iMondayRate != 0)
                                                                objRateCardRec.SetField("rate_Monday", iMondayRate.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_Monday", "0.00");
                                                        }
                                                        else
                                                            objRateCardRec.SetField("rate_Monday", "0.00");
                                                        //    AddContent("JUST ADDED" );
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
                                                        //    AddContent("TEST" + sSundayRate);

                                                        if (sSundayRate != "")
                                                        {
                                                            //  AddContent(sSundayRate);
                                                            if (iSundayRate != 0)
                                                                objRateCardRec.SetField("rate_sunday", iSundayRate.ToString());
                                                            else
                                                                objRateCardRec.SetField("rate_sunday", "0.00");
                                                        }

                                                        else
                                                            objRateCardRec.SetField("rate_sunday", "0.00");

                                                        if (sClassType != "")
                                                        {
                                                            objRateCardRec.SetField("rate_customtype", sClassType);
                                                        }


                                                        if (sCusSheet != "")
                                                        {
                                                            objRateCardRec.SetField("rate_customsheet", sCusSheet);
                                                        }


                                                        objRateCardRec.SetField("rate_Monday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_tuesday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_wednesday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_thrusday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_friday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_saturday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_sunday_CID", sBaseCurrency);
                                                        objRateCardRec.SetField("rate_price_CID", sBaseCurrency);

                                                        objRateCardRec.SetField("rate_maxheight", sRateMaxHeight);
                                                        objRateCardRec.SetField("rate_width1", sRateWidth1);
                                                        objRateCardRec.SetField("rate_width2", sRateWidth2);
                                                        objRateCardRec.SetField("rate_width3", sRateWidth3);
                                                        objRateCardRec.SetField("rate_width4", sRateWidth4);
                                                        objRateCardRec.SetField("rate_width5", sRateWidth5);
                                                        objRateCardRec.SetField("rate_width6", sRateWidth6);
                                                        objRateCardRec.SetField("rate_width7", sRateWidth7);
                                                        objRateCardRec.SetField("rate_width8", sRateWidth8);
                                                        objRateCardRec.SetField("rate_width9", sRateWidth9);
                                                        objRateCardRec.SetField("rate_width10", sRateWidth10);
                                                        objRateCardRec.SetField("rate_width11", sRateWidth11);
                                                        objRateCardRec.SetField("rate_width12", sRateWidth12);
                                                        //    AddContent(sRateWidth1 + "CAT" + sRateMaxHeight);
                                                        objRateCardRec.SaveChanges();


                                                        sMondayRate = "0.00";
                                                        sTuesdayRate = "0.00";
                                                        sWednesdayRate = "0.00";
                                                        sThursdayRate = "0.00";
                                                        sFridayRate = "0.00";
                                                        sSaturdayRate = "0.00";
                                                        sSundayRate = "0.00";
                                                        iMondayRate = 0;
                                                        iTuesdayRate = 0;
                                                        iWednesdayRate = 0;
                                                        iThrusdayRate = 0;
                                                        iFridayRate = 0;
                                                        iSaturdayRate = 0;
                                                        iSundayRate = 0;
                                                        //     AddContent("LOOP END" + sMondayRate);



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
                                                    else if (mWSheet1.UsedRange.Rows.Count == 1)
                                                    {
                                                        //  AddContent("EMPTY");
                                                        Process[] processes = Process.GetProcessesByName("EXCEL");

                                                        foreach (var process in processes)
                                                        {
                                                            if (process.MainWindowTitle == "")
                                                                process.Kill();
                                                        }
                                                        mWSheet1 = null;
                                                        mWorkBook = null;

                                                        GC.WaitForPendingFinalizers();
                                                        GC.Collect();
                                                        GC.WaitForPendingFinalizers();
                                                        GC.Collect();
                                                        //' mWorkBook.Close();

                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                                                        GeneratelogFile("[" + System.DateTime.Now.ToString() + "] " + sSuccessErrorMessage);
                                                        string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                                                        sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&hasrows=Y&dotnetfunc=RunRateCardImportStatusPage&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName;
                                                        // Dispatch.Redirect(sURL);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                //  AddContent("DIS");
                                                Process[] processes = Process.GetProcessesByName("EXCEL");

                                                foreach (var process in processes)
                                                {
                                                    if (process.MainWindowTitle == "")
                                                        process.Kill();
                                                }
                                                mWSheet1 = null;
                                                mWorkBook = null;

                                                GC.WaitForPendingFinalizers();
                                                GC.Collect();
                                                GC.WaitForPendingFinalizers();
                                                GC.Collect();
                                                //' mWorkBook.Close();

                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                                                //AddContent("stuffed");
                                                sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " Publication " + sPublication + " does not exist in Sage CRM.";
                                                isDatavalid = "true";
                                                iRecordcount++;
                                            }
                                        }
                                        else
                                        {
                                            //  AddContent("BADWOLF");
                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " Publication column is empty in Excel Sheet.";
                                            //GeneratelogFile("[" + System.DateTime.Now.ToString() + "] " + sSuccessErrorMessage);                                            
                                        }
                                    }
                                    catch (Exception Ex)
                                    {
                                        //    AddContent("DAT");
                                        Process[] processes = Process.GetProcessesByName("EXCEL");

                                        foreach (var process in processes)
                                        {
                                            if (process.MainWindowTitle == "")
                                                process.Kill();
                                        }
                                        mWSheet1 = null;
                                        mWorkBook = null;

                                        GC.WaitForPendingFinalizers();
                                        GC.Collect();
                                        GC.WaitForPendingFinalizers();
                                        GC.Collect();
                                        //' mWorkBook.Close();

                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                                        isColumnValid = "false";
                                        AddContent(Ex.Message.ToString());
                                        string strURL = UrlDotNet(this.ThisDotNetDll, "RunRateCardImportStatusPage");
                                        strURL += "&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName + "&ValidColumn=ROW";
                                        //  Dispatch.Redirect(strURL);
                                        return;
                                    }
                                }
                            }
                            else
                            {
                                //  AddContent("ERROR");
                                Process[] processes = Process.GetProcessesByName("EXCEL");

                                foreach (var process in processes)
                                {
                                    if (process.MainWindowTitle == "")
                                        process.Kill();
                                }
                                mWSheet1 = null;
                                mWorkBook = null;

                                GC.WaitForPendingFinalizers();
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                GC.Collect();
                                //' mWorkBook.Close();

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                                GeneratelogFile("[" + System.DateTime.Now.ToString() + "]" + " No records found in the excel sheet.");
                                string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                                sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&hasrows=N&dotnetfunc=RunRateCardImportStatusPage&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName;
                                //    Dispatch.Redirect(sURL);
                            }
                        }
                        #endregion

                        if (isDatavalid == "true" && mWSheet1.UsedRange.Rows.Count > 0 && iDupRecord != mWSheet1.UsedRange.Rows.Count)
                        {

                            Process[] processes = Process.GetProcessesByName("EXCEL");

                            foreach (var process in processes)
                            {
                                if (process.MainWindowTitle == "")
                                    process.Kill();
                            }
                            mWSheet1 = null;
                            mWorkBook = null;

                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            //' mWorkBook.Close();
                            //    AddContent("LOLOLOL");
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                            //   AddContent("MADE");
                            RefreshMetata();
                            //  AddContent("IT");
                            GeneratelogFile("[" + System.DateTime.Now.ToString() + "] " + sSuccessErrorMessage);
                            string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                            sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&hasrows=Y&dotnetfunc=RunRateCardImportStatusPage&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName + "&AllDup=N";
                            //  AddContent("GET OUT");
                            Dispatch.Redirect(sURL);
                        }
                        if (iDupRecord == mWSheet1.UsedRange.Rows.Count && isDatavalid == "true")
                        {
                            // AddContent("VOUBNGF");
                            Process[] processes = Process.GetProcessesByName("EXCEL");

                            foreach (var process in processes)
                            {
                                if (process.MainWindowTitle == "")
                                    process.Kill();
                            }
                            mWSheet1 = null;
                            mWorkBook = null;

                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            //' mWorkBook.Close();

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
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
