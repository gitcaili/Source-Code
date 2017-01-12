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
    public class SubSectionPageNew : DataPageNew
    {
        int iSectionID = 0;

        public SubSectionPageNew()
            : base("subSection", "suse_subsectionid", "SubSectionDetailBox")
        {
            #region Get current equipment id
            if (!String.IsNullOrEmpty(Dispatch.EitherField("sctn_Sctn_sectionid")))
            {
                iSectionID = Convert.ToInt32(Dispatch.EitherField("sctn_Sctn_sectionid"));
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key37")))
            {
                iSectionID = Convert.ToInt32(Dispatch.EitherField("Key37"));
            }
            #endregion

            this.SaveMethod = "RunSubSectionPage&sctn_Sctn_sectionid=" + iSectionID;
            this.CancelMethod = "RunSubSectionPage&sctn_Sctn_sectionid=" + iSectionID;
        }

        public override void BuildContents()
        {
            base.BuildContents();
        }
        public override bool Validate()
        {
            return base.Validate();
        }
    }
}
