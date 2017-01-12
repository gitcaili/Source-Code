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
    public class PlanConflict:Web
    {
        public PlanConflict()
            : base()
        {
 
        }
        public override void BuildContents()
        {
            string sValue = "";
            List objResultsGrid = new List("PlanList");
            Record objMatchRules = FindRecord("matchrules", "MaRu_TableName=N'booking'");
            string sWhere = "";
            while (!objMatchRules.Eof())
            {
                sValue = CurrentUser.SessionRead("PlanName").ToString().Replace("'", "''");
                string sField = objMatchRules.GetFieldAsString("MaRu_FieldName");
                string sType = "";
                if (!String.IsNullOrEmpty(Dispatch.EitherField("PlanName")) && !String.IsNullOrEmpty(objMatchRules.GetFieldAsString("MaRu_FieldName")))
                {
                    sType = objMatchRules.GetFieldAsString("MaRu_MatchType");
                    if (sType.ToUpper() == "CONTAINS")
                        sWhere = sWhere + sField + " like N'%" + sValue + "%'";
                    else if (sType.ToUpper() == "DOESNTMATCH")
                        sWhere = sWhere + sField + " <> N'" + sValue + "'";
                    else if (sType.ToUpper() == "EXACT")
                        sWhere = sWhere + sField + " = N'" + sValue + "'";
                    else if (sType.ToUpper() == "PHONETIC")
                        sWhere = sWhere + "SUBSTRING(SOUNDEX(" + sField + "), 2, 3) = SUBSTRING(SOUNDEX('" + sValue + "'), 2, 3)";
                    else if (sType.ToUpper() == "STARTINGWITH")
                        sWhere = sWhere + sField + " like N'" + sValue + "%'";
                }
                objMatchRules.GoToNext();
            }
            Record objPlanRec = FindRecord("Booking", sWhere);
            if (sWhere == "" || objPlanRec.Eof())
            {
                Dispatch.Redirect(UrlDotNet(ThisDotNetDll, "RunPlanNewPage") + "&PlanName=dedupe");
            }
            objResultsGrid.Filter = sWhere;
            AddContent(objResultsGrid);

            AddUrlButton(Metadata.GetTranslation("GenCaptions", "DedupeIgnorebooking"), "nextcircle.gif", UrlDotNet(ThisDotNetDll, "RunPlanNewPage") + "&PlanName=" + sValue + "&From=ignore");
            AddUrlButton(Metadata.GetTranslation("GenCaptions", "DedupeBackbooking"), "prevcircle.gif", UrlDotNet(ThisDotNetDll, "RunPlanDedupePage"));
        }
    }
}
