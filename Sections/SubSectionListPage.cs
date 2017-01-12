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
    public class SubSectionListPage : ListPage
    {
        int iSectionId = 0;
        EntryGroup objSectionsDetailsBox;
        public SubSectionListPage()
            : base("suse_subsectionid", "SubSectionsList", "SubSectionFilterBox")
        {
            base.OnLoad = "javascript:HideFilterScreen();";
            this.ResultsGrid.RowsPerScreen = 10;

            #region Get current Section id
            if (!String.IsNullOrEmpty(Dispatch.EitherField("sctn_Sctn_sectionid")))
            {
                iSectionId = Convert.ToInt32(Dispatch.EitherField("sctn_Sctn_sectionid"));
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key37")))
            {
                iSectionId = Convert.ToInt32(Dispatch.EitherField("Key37"));
            }
            #endregion
        }

        public override void BuildContents()
        {
            try
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
               
                Record objSectionRecord = FindRecord("Sections", "sctn_Sctn_sectionid=" + iSectionId);
                objSectionsDetailsBox = new EntryGroup("SectionsEntryScreen", "Summary");
                objSectionsDetailsBox.Fill(objSectionRecord);
                AddContent(objSectionsDetailsBox.GetHtmlInViewMode(objSectionRecord));
                this.ResultsGrid.Filter = "suse_Deleted is null and suse_section ='" + iSectionId + "'";
                base.BuildContents();
                AddUrlButton("Add Sub Sections", "NewLineItem.gif", UrlDotNet(ThisDotNetDll, "RunSubSectionPageNew") + "&Key37=" + iSectionId + "&J=Summary&Key58=" + iSectionId + "&T=subsection");
                AddUrlButton("Continue", "Continue.gif", UrlDotNet(ThisDotNetDll, "RunSectionListPage"));

            }
            catch (Exception Ex)
            {
                this.AddError(Ex.Message);
            }
        }

        public override void AddNewButton()
        {
            //base.AddNewButton();
        }
    }
}
