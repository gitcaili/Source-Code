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
    class PlanAgencyCodeAjaxCallPage : Web
    {
        string sKey1 = "";
        string sCompanyAgencyCode = "";
        public PlanAgencyCodeAjaxCallPage()
        {

        }
        public override void BuildContents()
        {
            try
            {
                if (!String.IsNullOrEmpty(Dispatch.EitherField("Key1")))
                {
                    sKey1 = Dispatch.EitherField("Key1");
                }
                
                Record recAgency = FindRecord("company", "Comp_CompanyId =" + sKey1);
                if (!recAgency.Eof())
                {
                    //AddContent("sKey1 =" + sKey1 + " " + recAgency.GetFieldAsString("comp_agencycode"));
                    if (!String.IsNullOrEmpty(recAgency.GetFieldAsString("comp_agencycode")))
                        sCompanyAgencyCode = recAgency.GetFieldAsString("comp_agencycode");
                }
                AddContent("<returnmsg>" + sCompanyAgencyCode + "</returnmsg>");
                
            }
            catch (Exception ex)
            {
                this.AddError(ex.Message);
            }
        }
    }
}
