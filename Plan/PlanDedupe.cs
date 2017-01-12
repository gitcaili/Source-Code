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
   public class PlanDedupe:Web
    {
       public  PlanDedupe()
           :base()
       {

       }
           public override void BuildContents()
           {
               string sFlag = "";

               AddContent(HTML.Form());
               
               #region Define the Hidden Fields
               AddContent(HTML.InputHidden("HiddenMode", ""));
               #endregion

               EntryGroup objPlanBoxDedupe = new EntryGroup("BookingBoxDedupe", "booking");

               if (!String.IsNullOrEmpty(Dispatch.EitherField("HiddenMode")))
               {
                   if (objPlanBoxDedupe.Validate() == true)
                   {
                       CurrentUser.SessionWrite("PlanName", Dispatch.ContentField("book_name"));
                       string sURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do?" + "SID=" + Dispatch.EitherField("SID") + "&ACT=432&Key0=58&PlanName=dedupe&dotnetdll=NZPACRM&dotnetfunc=RunPlanConflictPage&T=New&J=Booking";                       
                       Dispatch.Redirect(sURL);
                   }
                   else
                   {
                       AddError("Validation Errors - Please correct the highlighted entries");
                   }
               }

               Record objmatchRules = FindRecord("matchRules", "MaRu_TableName=N'booking'");
                if(objmatchRules.Eof())
                {
                    string sURL = "http://" + Dispatch.Host + "/" + Dispatch.InstallName + "/eware.dll/Do?" + "SID=" + Dispatch.EitherField("SID") + "&ACT=432&Key0=58&PlanName=dedupe&dotnetdll=NZPACRM&dotnetfunc=RunPlanNewPage&T=New&J=Booking";                    
                    Dispatch.Redirect(sURL);
                }
               AddContent(HTML.InputHidden("HiddenName", ""));
               AddContent(objPlanBoxDedupe.GetHtmlInEditMode());
               AddSubmitButton("Enter Plan Details", "nextcircle.gif", "javascript:document.EntryForm.HiddenMode.value='dedupe';PlanDedupe();");
           }
    }
}
