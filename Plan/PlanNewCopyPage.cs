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
    class PlanNewCopyPage : Web
    {
        string sPlanId = "";
        CRMHelper objCRM = new CRMHelper();
        public PlanNewCopyPage()
        {
            //GetTabs("Booking", "Booking Summary");
        }
        public override void BuildContents()
        {
            try
            {
                #region PlanId
                if (!String.IsNullOrEmpty(Dispatch.EitherField("book_BookingID")))
                {
                    sPlanId = Dispatch.EitherField("book_BookingID");
                }
                #endregion

                #region MyRegion
                Record recPlan = FindRecord("Planbuilder", "pnbr_plan =" + sPlanId);
                if (!recPlan.Eof())
                {
                    List objList = new List("PlanBuilderList");
                    objList.PadBottom = false;
                    GridColCheckBox RowSelect = new GridColCheckBox("pnbr_select");
                    RowSelect.ReadOnly = false;
                    RowSelect.ShowHeading = true;
                    objList.Add(RowSelect);
                    objList.CheckBoxColumn = "Select/Deselect";
                    objList.ShowSelectUnselectButton = true;
                    objList.SelectUnselectButtonOnClickScript = "javascript:checkAll();";
                    if (sPlanId != "")
                    {
                        objList.Filter = "pnbr_plan =" + sPlanId + " and pnbr_Deleted is null";
                    }
                    AddContent(objList);
                #endregion
                }
                else
                {
                    string strMessage = "There is no Plan Builder Details attached to this Plan. Click on Copy Plan button to proceed further.";
                   // objCRM.GetStatusBlock("Booking", strMessage, "true", sPlanId.ToString());
                    AddInfo(strMessage);
                }
                string sUrl = "javascript:CopyPlan();";
                //AddUrlButton("Copy Plan", "save.gif", sUrl);
                AddUrlButton("Copy Plan Line Items", "save.gif", sUrl);
               
                string sPreviousUrl = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do?SID=" + Dispatch.EitherField("SID") + "&Act=432&Mode=1&CLk=T&Key0=58&Key58=" + Dispatch.EitherField("book_BookingID") + "&book_BookingID=" + Dispatch.EitherField("book_BookingID") + "&dotnetdll=NZPACRM&dotnetfunc=RunPlanSummaryPage&J=Booking Summary&J=Booking Summary&T=Booking";
                AddUrlButton("Cancel", "cancel.gif", sPreviousUrl);
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
        }
    }
}
