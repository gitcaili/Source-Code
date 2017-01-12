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
     public class PlanRedirector : Web
    {
         int iPlanId = 0;
         public PlanRedirector()
             : base()
         {
 
         }
         public override void BuildContents()
         {
             AddError("This Record has been deleted");
         }
    }
}
