using System;

using Sage.CRM.WebObject;

using Sage.CRM.Data;
using NZPACRM.Common;

using System.IO;

using System.Diagnostics;
namespace NZPACRM.Plan
{
    public class PlanImportSS : DataPageNew
    {
        Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        Microsoft.Office.Interop.Excel.Application oXL;
        CRMHelper objCRM = new CRMHelper();
        string LogfileName = "";
        string shttpURL = "";
        Boolean newplan = true;
        //   OleDbConnection oleExcelConnection = default(OleDbConnection);
        public PlanImportSS()
            : base("Booking", "book_BookingID", "")
        {
            string CurrUser = CurrentUser.UserId.ToString(); // the current user

            #region get Http from url
            try
            {
                string s = Dispatch.ServerVariable("HTTP_REFERER");
                char[] cSplit = { '/' };
                string[] sHTTP = s.Split(cSplit);

                if (!String.IsNullOrEmpty(sHTTP[0]))
                    shttpURL = sHTTP[0]; // checking to see if it is security or something?


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

        //public OleDbConnection oleExcelConnection = default(OleDbConnection);
        public override void BuildContents()
        {
            int bid2 = -1;
            int weekcurr = 1;
            try
            {
                //OleDbConnection conexel;
                //AddContent("<script type='text/javascript' src='../CustomPages/Booking/ClientFuncs.js'></script>");
                AddContent("<script type='text/javascript' src='../js/custom/ClientFuncs.js'></script>");
                string sSuccessErrorMessage = "";
                int iFailedCount = 0;
                int InsertCount = 0;
                int iDupRecord = 0;
                string isDatavalid = "";

                #region Adding Html Form
                AddContent(HTML.Form());
                #endregion

                string InstructionText = "";
                string sHTML = "";
                string sFileURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/";
                InstructionText += "<span style='margin: 1em 3em 10px 5.5em;'><B> 1.</B> Download the template to be used for Import Process by clicking on <B>Download Template</B> link.</span>";
                InstructionText += " <BR> ";
                InstructionText += "<span style='margin: 1em 3em 10px 5.5em;'><B> 2.</B> Copy and Paste all the data to be imported in the downloaded sheet and Save the same on your machine.</span>";
                InstructionText += " <BR> ";
                InstructionText += "<span style='margin: 1em 3em 10px 5.5em;'><B> 3.</B> Click on <B>Browse</B> button to get the excel file to be used for import process. (saved in above step)</span>";
                InstructionText += " <BR> ";
                InstructionText += "<span style='margin: 1em 3em 10px 5.5em;'><B> 4.</B> Once file is selected click on <B>Import button</B> to start import process.</span>";
                InstructionText += " <BR> ";
                InstructionText += "<span style='margin: 1em 3em 10px 5.5em;'><B> 5.</B> This process will import records in Sage CRM.</span>";
                InstructionText += " <BR> ";
                InstructionText += " <BR> ";

                sHTML += HTML.StartTable();
                sHTML += HTML.TableRow("");
                sHTML += HTML.TableData(InstructionText, "AdminHomeDescription");
                sHTML += "<BR>";
                sHTML += HTML.TableRow("");
                //sHTML += HTML.TableRow("");

                sHTML += HTML.TableData("<span style='margin: 1em 3em 10px 55em;font-Size:14px;'><a href='" + sFileURL + "NZPAImport/Templates/PlanTemplate.xlsx' class='PANEREPEAT'><u>Download Template</u></a></span>");
                sHTML += HTML.EndTable();
                sHTML += "<BR>";
                AddContent(HTML.Box("Import Plan Data", sHTML));

                AddContent("<BR><BR>" + HTML.Box("File", "<br>&nbsp;&nbsp;<input type='file' id='fileupload' name='pic' size='70'>&nbsp;<input type='BUTTON' class='Edit'value='Import'name='upload' onclick='javascipt:CheckFileNew();'></br></br>"));

                string backURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                backURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=1650&Mode=1&CLk=T&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data Management";
                AddUrlButton("Back", "prevcircle.gif", backURL);

                AddContent(HTML.InputHidden("HIDDEN_LibraryPath", GetLibraryPath().Replace("Library", "")));
                AddContent(HTML.InputHidden("HIDDEN_FilePath", ""));
                AddContent(HTML.InputHidden("HIDDEN_Save", ""));
                AddContent(HTML.InputHidden("HIDDEN_FileName", ""));
                AddContent(HTML.InputHidden("HIDDEN_FilePathChrome", ""));
                AddContent(HTML.InputHidden("HIDDEN_browser", ""));

                if (!String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_Save")))
                {
                    if (Dispatch.ContentField("HIDDEN_Save") == "Save")
                    {
                        AddContent(Dispatch.ContentField("HIDDEN_FilePath"));
                        string SavedFilePath = "";

                        if (String.IsNullOrEmpty(Dispatch.ContentField("HIDDEN_FilePathChrome")))
                        {
                            SavedFilePath = SaveRateCardRLocation();
                        }
                        else
                        {
                            SavedFilePath = GetLibraryPath() + @"\RateCard\" + Dispatch.ContentField("HIDDEN_FileName");
                        }

                        //  DataTable dt = new DataTable();
                        oXL = new Microsoft.Office.Interop.Excel.Application();
                        oXL.DisplayAlerts = false;

                        try
                        {
                            mWorkBook = oXL.Workbooks.Open(SavedFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                            var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; data source=" + SavedFilePath + "; Extended Properties='Excel 12.0;HDR=NO;IMEX=1;'", SavedFilePath);
                            mWorkSheets = mWorkBook.Worksheets;
                            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item(1);
                            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;

                            //   conexel = new OleDbConnection(connectionString);
                            //  OleDbCommand cmdExcel = new OleDbCommand();
                            //var adapter = new OleDbDataAdapter();
                            //cmdExcel.Connection = conexel;
                            // conexel.Open();
                            // cmdExcel.CommandText = "SELECT * FROM [sheet1$]";
                            // adapter.SelectCommand = cmdExcel;

                            // var adapter = new OleDbDataAdapter("SELECT * FROM [sheet1$]", connectionString);

                            //var ds = new DataSet();
                            //adapter.Fill(ds);

                            //   adapter.Fill(dt);

                            // conexel.Close();
                            AddContent("START");
                            for (int t = 4; t < 10; t++)
                            {
                                AddContent("t is " + t + " " +(string)(mWSheet1.Cells[t, 5] as Microsoft.Office.Interop.Excel.Range).Value);
                            }
                            AddContent("DOC");
                            mWSheet1.Rows.ClearFormats();
                           
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
                            string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                            sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=ImportCodenew&dotnetfunc=RunPlanImportStat&inserted=" + InsertCount + "&name=" + bid2;
                          //  Dispatch.Redirect(sURL);

                        }
                        catch (IndexOutOfRangeException e)
                        {
                            string sURL = shttpURL + "//" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do";
                            sURL += "?SID=" + Dispatch.EitherField("SID") + "&Act=432&dotnetdll=ImportCodenew&dotnetfunc=RunPlanImportStat&inserted=" + InsertCount + "&name=" + bid2;
                            //  Dispatch.Redirect(sURL);

                        }


                        catch (Exception Ex)
                        {

                            this.AddError(Ex.Message);

                        }

                        finally
                        {

                        }


                    }
                }
            }
            // }
            catch (Exception Ex)
            {
                this.AddError(Ex.Message);
            }



        }

        private string comissionconvert(string co)
        {
            if (co == null) return "Government";
            else if (co.Equals("Y")) return "Commission";
            else return "NonCommission";
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
            //string UploadfilePath = "C:\\Users\\Administrator\\Desktop\\Sia.xlsx";
            //AddContent(UploadfilePath);
            string LibPath = GetLibraryPath();
            string NewPath = LibPath.Replace("\\Library", "");
            FileName = Dispatch.ContentField("HIDDEN_FileName");
            NewPath += @"\WWWRoot\CustomPages\NZPAImport\ImportedFiles\Plans\";
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
                    // AddContent("HORSE");
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