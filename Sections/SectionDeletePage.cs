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
    public class SectionDeletePage : DataPageDelete
    {
        int iSectionsID = 0;
        public SectionDeletePage()
            : base("Sections", "sctn_Sctn_sectionid", "SectionsEntryScreen")
        {
            SaveMethod = "RunSectionListPage";
            CancelMethod = "RunSectionListPage";

            #region Get current equipment id
            base.OnLoad = "javascript:SetContextHyperLink();";
            if (!String.IsNullOrEmpty(Dispatch.EitherField("sctn_Sctn_sectionid")))
            {
                iSectionsID = Convert.ToInt32(Dispatch.EitherField("sctn_Sctn_sectionid"));
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key37")))
            {
                iSectionsID = Convert.ToInt32(Dispatch.EitherField("Key37"));
            }
            #endregion
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

            base.BuildContents();
        }

        public override void AfterSave(EntryGroup screen)
        {
            Record objSectionsRec = FindRecord("Sections", "sctn_Sctn_sectionid='" + iSectionsID + "'");
            while (!objSectionsRec.Eof())
            {
                objSectionsRec.SetField("prmt_Deleted", 1);
                objSectionsRec.GoToNext();
            }

            objSectionsRec.SaveChanges();
            base.AfterSave(screen);
        }
    }
}
