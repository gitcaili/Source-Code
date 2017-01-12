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
    public class SubSectionPageDelete : DataPageDelete
    {
        string sSectionID = "";
        public SubSectionPageDelete()
            : base("suse_subsectionid", "suse_subsectionid", "SubSectionDetailBox")
        {
            int iSubSectionID = 0;
            #region Sub Section ID
            if (!String.IsNullOrEmpty(Dispatch.EitherField("suse_subsectionid")))
            {
                iSubSectionID = Convert.ToInt32(Dispatch.EitherField("suse_subsectionid"));
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key37")))
            {
                iSubSectionID = Convert.ToInt32(Dispatch.EitherField("Key37"));
            }
            #endregion

            #region Get current Section id

            Record objSectionRec = FindRecord("subSection", "suse_subsectionid=" + iSubSectionID);
            if (!objSectionRec.Eof())
            {
                sSectionID = objSectionRec.GetFieldAsString("suse_section");
            }
            #endregion
            this.CancelMethod = "RunSubSectionPage&sctn_Sctn_sectionid=" + sSectionID;
        }

        public override void BuildContents()
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

            #region Get current Section id
            if (!String.IsNullOrEmpty(Dispatch.EitherField("sctn_Sctn_sectionid")))
            {
                sSectionID = Dispatch.EitherField("sctn_Sctn_sectionid");
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key58")))
            {
                sSectionID =Dispatch.EitherField("Key58");
            }
            #endregion

            base.BuildContents();
        }

        public override void AfterSave(EntryGroup screen)
        {
            Dispatch.Redirect(UrlDotNet(ThisDotNetDll, "RunSubSectionPage") + "&sctn_Sctn_sectionid=" + sSectionID);

            base.AfterSave(screen);
        }
    }
}
