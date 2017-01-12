using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using Sage.CRM.Controls;
using Sage.CRM.Data;
using NZPACRM.Common;
using System.Data;
using System.IO;

namespace NZPACRM.Sections
{
    public class SectionPageEdit: DataPageEdit
    {
        int iSectionID=0;
        public SectionPageEdit()
            : base("Sections", "sctn_Sctn_sectionid", "SectionsEntryScreen")
        {
           
            if (!String.IsNullOrEmpty(Dispatch.EitherField("T")))
            {
                if (Dispatch.EitherField("T").ToString().ToLower() != "sections")
                    GetTabs("Sections", "Summary");
            }
            else
            {
                GetTabs("Sections", "Summary");
            }

            #region Get current equipment id
            if (!String.IsNullOrEmpty(Dispatch.EitherField("sctn_sectionid")))
            {
                iSectionID = Convert.ToInt32(Dispatch.EitherField("sctn_sectionid"));
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key37")))
            {
                iSectionID = Convert.ToInt32(Dispatch.EitherField("Key37"));
            }
            #endregion

            SaveMethod = "RunSectionListPage&T=Sections&J=Summary";
            CancelMethod = "RunSectionListPage&T=Sections&J=Summary";
            DeleteMethod = "RunSectionPageDelete&sctn_Sctn_sectionid=" + iSectionID;
        }

        public override void  BuildContents()
        {
 	         base.BuildContents();
        }
    }
}
