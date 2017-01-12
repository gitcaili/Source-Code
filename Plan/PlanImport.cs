    using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;
using System.Data;
using System.IO;

namespace NZPACRM.Plan
{
    public class PlanImport : DataPageNew
    {
        CRMHelper objCRM = new CRMHelper();
        string LogfileName = "";
        string shttpURL = "";
        string sAgency = "";
        string sAgencyID = "";
        string sContactF = "";
        string sContactL = "";
        string sContactID = "";
        string sClient = "";
        string sDocVersion = "";
        string sAgencyCode = "";
        string sRef = "";
        string sBookName = "";
        string sBookingid = "";
        string sBilledBy = "";
        string sDesc = "";
        string sCostingVersion = "";
        string sCreatedBy = "";
        string sOpened = "";
        public PlanImport()
            : base("Booking", "book_bookingid", "")
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
                decimal iLoading = 0.00m;
                string sSuccessErrorMessage = "";
                int iFailedCount = 0;
                int InsertCount = 0;
                int iDupRecord = 0;
                string isDatavalid = "";
                string sBaseCurrency = "";
                
                #region Adding Html Form
                AddContent(HTML.Form());
                #endregion

                #region Get Template Block
                objCRM.GetTemplateBlock("Booking");
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
                            Record RecEntityID = FindRecord("Custom_Tables", "bord_name='Booking' and bord_deleted is null");
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
                                    string sSectionFlag = "";
                                    string sPublicationId = "";
                                    string isValid = "true";
                                    string isColumnValid = "true";
                                    int iRowCurrent = i + 2;
                                    string sExist = "T";                                    
                                    int iRecordcount = 0;
                                    string sAllday = "";
                                    
                                    try
                                    {
                                        if (!String.IsNullOrEmpty(dt.Rows[i]["Name"].ToString().Trim()))
                                        {                                            
                                            sBookName = dt.Rows[i]["Name"].ToString().Trim();                                            
                                            Record objBook = FindRecord("Booking", "LOWER(LTRIM(RTRIM(book_name)))='" + dt.Rows[i]["Name"].ToString().Trim().ToLower() + "'");
                                            
                                            if (objBook.Eof())
                                            {   
                                                sAgency = dt.Rows[i]["Agency"].ToString().Trim();
                                                Record objAgency = FindRecord("Company", "LOWER(LTRIM(RTRIM(comp_Name)))='" + dt.Rows[i]["Agency"].ToString().Trim().ToLower() + "'");

                                                if (!objAgency.Eof())
                                                {
                                                    sAgencyID = objAgency.GetFieldAsString("Comp_CompanyId");

                                                    sContactF = dt.Rows[i]["Contact First Name"].ToString().Trim();
                                                    sContactL = dt.Rows[i]["Contact Last Name"].ToString().Trim();
                                                    Record objPerson = FindRecord("Person", "Pers_CompanyId='" + sAgencyID + "' and LOWER(LTRIM(RTRIM(Pers_FirstName))) ='" + sContactF.Trim().ToLower() + "' and LOWER(LTRIM(RTRIM(Pers_LastName))) ='" + sContactL.Trim().ToLower() + "'");

                                                    if (!objPerson.Eof())
                                                    {
                                                        sContactID = objPerson.GetFieldAsString("Pers_PersonId");
                                                        isValid = "true";
                                                    }
                                                    else
                                                    {
                                                        sSuccessErrorMessage += Environment.NewLine + sBookName + " Plan is not imported as person is not exixts for selected company";
                                                        isValid = "false";
                                                        iFailedCount++;
                                                    }
                                                }
                                                else
                                                {
                                                    sSuccessErrorMessage += Environment.NewLine + sBookName + " Plan is not imported as Company is not exixts in Sage CRM";
                                                    isValid = "false";
                                                    iFailedCount++;
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Client"].ToString().Trim()))
                                                {
                                                    sClient = dt.Rows[i]["Client"].ToString().Trim();

                                                    Record recClient = FindRecord("Client", "LOWER(LTRIM(RTRIM(client_Name)))='" + sClient + "'");

                                                    if (!recClient.Eof())
                                                    {
                                                        sClient = recClient.GetFieldAsString("client_ClientID");
                                                    }
                                                    else
                                                    {
                                                        CreateClient();
                                                    }
                                                }

                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Agency Code"].ToString().Trim()))
                                                {
                                                    sAgencyCode = dt.Rows[i]["Agency Code"].ToString().Trim();
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Reference"].ToString().Trim()))
                                                {
                                                    sRef = dt.Rows[i]["Reference"].ToString().Trim();
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Billed by"].ToString().Trim()))
                                                {
                                                    sBilledBy = dt.Rows[i]["Billed by"].ToString().Trim();
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Description"].ToString().Trim()))
                                                {
                                                    sDesc = dt.Rows[i]["Description"].ToString().Trim();
                                                }
                                                if (!String.IsNullOrEmpty(dt.Rows[i]["Costing Version"].ToString().Trim()))
                                                {
                                                    sCostingVersion = dt.Rows[i]["Costing Version"].ToString().Trim();
                                                }

                                                if (dt.Rows.Count >= 1)
                                                {
                                                    if (isValid == "true")
                                                    {
                                                        isDatavalid = isColumnValid;

                                                        Record objRateCardRec = new Record("Booking");

                                                        objRateCardRec.SetField("book_agency", sAgencyID);
                                                        objRateCardRec.SetField("book_Contact", sContactID);
                                                        objRateCardRec.SetField("book_Client", sClient);
                                                        objRateCardRec.SetField("book_DocumentVersion", sDocVersion);
                                                        objRateCardRec.SetField("book_agencycode", sAgencyCode);
                                                        objRateCardRec.SetField("book_reference", sRef);
                                                        objRateCardRec.SetField("book_opened", sOpened);
                                                        objRateCardRec.SetField("book_CreatedBy", sCreatedBy);
                                                        objRateCardRec.SetField("book_Name", sBookName);
                                                        objRateCardRec.SetField("book_billedby", sBilledBy);
                                                        objRateCardRec.SetField("book_description", sDesc);
                                                        objRateCardRec.SetField("book_costingversion", sCostingVersion);

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
                                                sSuccessErrorMessage += Environment.NewLine + sBookName + " Plan is not imported as it's already exist in Sage CRM";
                                                isDatavalid = "true";
                                                iDupRecord++;
                                            }
                                        }
                                        else
                                        {
                                            sSuccessErrorMessage += Environment.NewLine + "Row no " + iRowCurrent + " Plan " + sAgency + " Doesnt exist in Sheet.";                                           
                                        }
                                    }
                                    catch (Exception Ex)
                                    {
                                        isColumnValid = "false";
                                        iFailedCount++;
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
                                sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&hasrows=N&dotnetfunc=RunPlanImportStatusPage&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName;
                                Dispatch.Redirect(sURL);
                            }
                        }
                        #endregion
                        
                        if (isDatavalid == "true" && dt.Rows.Count > 0 && iDupRecord != dt.Rows.Count)
                        {
                            RefreshMetata();
                            GeneratelogFile("[" + System.DateTime.Now.ToString() + "] " + sSuccessErrorMessage);
                            string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                            sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&hasrows=Y&dotnetfunc=RunPlanImportStatusPage&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName + "&AllDup=N";
                            
                            Dispatch.Redirect(sURL);
                        }
                        if (iDupRecord == dt.Rows.Count && isDatavalid == "true")
                        {
                            GeneratelogFile("[" + System.DateTime.Now.ToString() + "] " + sSuccessErrorMessage);
                            string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                            sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=NZPACRM&hasrows=Y&dotnetfunc=RunPlanImportStatusPage&inserted=" + InsertCount + "&Failed=" + iFailedCount + "&LogPath=" + LogfileName + "&AllDup=E";
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

        public void CreateClient()
        {
            Record recClient = new Record("Client");

            recClient.SetField("client_Name", sClient);
            recClient.SetField("client_CompanyId", sAgencyID);
            recClient.SetField("client_CreatedBy", CurrentUser.UserId);

            recClient.SaveChanges();
            sClient = recClient.RecordId.ToString();
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
            NewPath += "\\WWWRoot\\CustomPages\\NZPAImport\\ImportedFiles\\Plan\\";
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
                string ymd = System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString() + System.DateTime.Now.Millisecond.ToString() + theDate.ToString("yyyyMMdd") + "PlanImport.txt";
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
