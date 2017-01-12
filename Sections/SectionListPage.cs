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
    public class SectionListPage : ListPage
    {
        string sURL = "";
        public SectionListPage()
            : base("Sections", "SectionsGrid", "SectionsFilterBox")
        {
            base.OnLoad = "javascript:HideSectionScreen();";
            //'base.OnLoad = "javascript:SetCloumnWdith();";
            this.ResultsGrid.RowsPerScreen = 10;
        }

        public override void BuildContents()
        {
            base.BuildContents();
            string sURL = UrlDotNet(ThisDotNetDll, "RunSectionPageNew");
            AddUrlButton("Add Sections", "NewDoc.gif", sURL);
            AddUrlButton("Back", "PrevCircle.gif", Url("1650") + "&MenuName=AdminDataManagement&BC=Admin,Admin,AdminDataManagement,Data Management");
           
            if (!String.IsNullOrEmpty(Dispatch.EitherField("T")))
            {
                if (Dispatch.EitherField("T").ToString().ToLower() != "sections")
                {
                    GetTabs("Sections", "Summary");
                }
                    
            }
            else
            {
                GetTabs("Sections", "Summary");
            }
        }
        
        public override void AddNewButton()
        {
            //base.AddNewButton();
        }
    }
}
