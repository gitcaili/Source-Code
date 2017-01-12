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
    public class PlanNotesPage : ListPage
    {
        string sPlanID = "";
        int iEntityId = 0;
        public PlanNotesPage()
            : base("Notes", "NoteList", "NoteFilterBox")
        {
            if (!String.IsNullOrEmpty(Dispatch.EitherField("book_bookingid")))
            {
                sPlanID = Dispatch.EitherField("book_bookingid").ToString();
            }
            else if (!String.IsNullOrEmpty(Dispatch.EitherField("Key58")))
            {
                sPlanID = Dispatch.EitherField("Key58").ToString();
            }
        }

        public override void BuildContents()
        {
            Record objEntityIdRec = FindRecord("custom_tables", "bord_name='Booking' and bord_deleted is null ");

            if (!objEntityIdRec.Eof())
                iEntityId = objEntityIdRec.GetFieldAsInt("Bord_tableid");

            AddTopContent(GetCustomEntityTopFrame("Booking"));
            this.ResultsGrid.Filter = "note_foreignid=" + sPlanID + "and note_foreigntableid=" + iEntityId;
            base.BuildContents();
        }
    
    }
}
