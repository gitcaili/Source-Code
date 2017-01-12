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
    public class SectionSubSectionPage : ListPage
    {
        string sURL = "";
        public SectionSubSectionPage()
            : base("Section", "SectionList", "SectionFilterBox")
        {
            base.OnLoad = "javascript:HideFilterButton();";
        }

        public override void BuildContents()
        {
            base.BuildContents();
            AddUrlButton("Add Equipments", "Equipment.gif", Url("343"));
            
        }

        public override void AddNewButton()
        {
            //base.AddNewButton();
        }
    }
}
