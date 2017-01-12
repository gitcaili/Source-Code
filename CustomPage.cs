using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;

namespace NZPACRM
{
    public class MyCustomPage : Web
    {

        public override void BuildContents()
        {
            //Add your content here!      
            GetTabs("Client");

            AddContent("My Custom Page");

            AddContent("<BR>");

            //how to show translated values - maybe
            AddContent(Metadata.GetTranslation(Sage.Captions.sFam_GenMessages, "HelloWorld"));

            AddContent("<BR>");

            //how to check sys param values
            AddContent("The Base Currency is: " + Metadata.GetParam(Sage.ParamNames.BaseCurrency));

            AddContent("<BR>");

            //...etc
            //Dispatch.

        }



    }

}

