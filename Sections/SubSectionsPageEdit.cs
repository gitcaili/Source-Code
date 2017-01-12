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
    public class SubSectionsPageEdit : DataPageEdit
    {
        string iSubSectionID = "";
        string iSectionID = "";

        public SubSectionsPageEdit()
            : base("subSection", "suse_subsectionid", "SubSectionDetailBox")
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

            #region Sub Section ID
            if (!String.IsNullOrEmpty(Dispatch.EitherField("suse_subsectionid")))
            {
                iSubSectionID = Dispatch.EitherField("suse_subsectionid");
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key37")))
            {
                iSubSectionID = Dispatch.EitherField("Key37");
            }
            #endregion


            #region Get current Section id

            Record objSubSectoinRec = FindRecord("subSection", "suse_subsectionid=" + iSubSectionID);
            if (!objSubSectoinRec.Eof())
            {
                iSectionID = objSubSectoinRec.GetFieldAsString("suse_section");

            }
            #endregion

            this.SaveMethod = "RunSubSectionPage&sctn_Sctn_sectionid=" + iSectionID;
            this.CancelMethod = "RunSubSectionPage&sctn_Sctn_sectionid=" + iSectionID;
            this.DeleteMethod = "RunSubSectionPageDelete&sctn_Sctn_sectionid=" + iSectionID;
        }
    }
}
