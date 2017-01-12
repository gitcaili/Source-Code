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
    public class SectionPageNew : DataPageNew
    {
        public SectionPageNew()
            : base("Sections", "sctn_Sctn_sectionid", "SectionsEntryScreen")
        {
            SaveMethod = "RunSectionListPage";
            CancelMethod = "RunSectionListPage";
        }

        public override void BuildContents()
        {
            base.BuildContents();
        }

    }
}
