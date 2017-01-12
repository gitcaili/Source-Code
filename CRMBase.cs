using System;
using System.Collections.Generic;
using System.Text;
using Sage.CRM.WebObject;
using NZPACRM.Common;
using NZPACRM.Plan;
using NZPACRM.Sections;
using NZPACRM.RateCard;



/* ********************************
 * Sage CRM Customizations
 * Name: NZPACRM 
 * Created On: 8/27/2008 9:27:21 AM
 * Created By: Greytrix
 * 
 *********************************/

namespace NZPACRM
{
    //static class AppFactory is REQUIRED!
    public static class AppFactory
    {

        public static void RunRateImportSS(ref Web Aretval)
        {
            Aretval = new RateCardSS();
        }
        public static void RunAjaxSetStandardSize(ref Web AretVal)
        {
            AretVal = new AjaxSetStandardSize();
        }
        public static void RunPlanAgencyCodeAjaxCallPage(ref Web AretVal)
        {
            AretVal = new PlanAgencyCodeAjaxCallPage();
        }
        public static void RunClientPhoneEmail(ref Web AretVal)
        {
            AretVal = new ClientPhoneEmailPage();
        }
        public static void RunPublicationPhoneEmail(ref Web AretVal)
        {
            AretVal = new PublicationPhoneEmailPage();
        }
        public static void RunPublisherPhoneEmailPage(ref Web AretVal)
        {
            AretVal = new PublisherPhoneEmailPage();
        }
        public static void RunPublisherConfirmDelete(ref Web AretVal)
        {
            AretVal = new PublisherConfirmDelete();
        }
        public static void RunClientAddress(ref Web AretVal)
        {
            AretVal = new ClientAddress();
        }
        public static void RunClientAddressList(ref Web AretVal)
        {
            AretVal = new ClientAddressList();
        }
        public static void RunClientEditAddressPage(ref Web AretVal)
        {
            AretVal = new ClientAddressEdit();
        }
        public static void RunPlanSS(ref Web AretVal)
        {
            AretVal = new PlanImportSS();
        }
        public static void RunClientConfirmDelete(ref Web AretVal)
        {
            AretVal = new ClientConfirmDelete();
        } 
        public static void RunPublicationAddressList(ref Web AretVal)
        {
            AretVal = new PublicationAddressListPage();
        }
        public static void RunPublicationEditAddressPage(ref Web AretVal)
        {
            AretVal = new PublicationAddressEdit();
        }
        public static void RunPublicationConfirmDelete(ref Web AretVal)
        {
            AretVal = new PublicationConfirmDelete();
        }
        public static void RunPublicationAddressNew(ref Web AretVal)
        {
            AretVal = new PublicationAddressNew();
        }
        public static void RunPublisherAddressList(ref Web AretVal)
        {
            AretVal = new PublisherAddressList();
        }
        public static void RunPublisherEditAddressPage(ref Web AretVal)
        {
            AretVal = new PublisherAddressEdit();
        }
        public static void RunPublisherAddressNew(ref Web AretVal)
        {
            AretVal = new PublisherAddressNew();
        }
        public static void RunImportClientPage(ref Web AretVal)
        {
            AretVal = new ImportClientPage();
        }
        public static void RunImportPublicationsPage(ref Web AretVal)
        {
            AretVal = new ImportPublicationsPage();
        }
        public static void RunImportPublishersPage(ref Web AretVal)
        {
            AretVal = new ImportPublishersPage();
        }
        public static void RunClientRedirectPage(ref Web AretVal)
        {
            AretVal = new Redirect();
        }
        public static void RunPlanOtherImport(ref Web AretVal)
        {
            AretVal = new PlanOtherImport();
        }
        public static void RunPublicationRedirectPage(ref Web AretVal)
        {
            AretVal = new RedirectorPublicationList();
        }
        public static void RunPublishersRedirectPage(ref Web AretVal)
        {
            AretVal = new RedirectorPublishersList();
        }
        public static void RunRedirectClientPhoneEmailPage(ref Web AretVal)
        {
            AretVal = new RedirectClientPhoneEmailPage();
        }
        public static void RunRedirectPublicationPhoneEmailPage(ref Web AretVal)
        {
            AretVal = new RedirectPublicationPhoneEmailPage();
        }
        public static void RunRedirectPublishersPhoneEmailPage(ref Web AretVal)
        {
            AretVal = new RedirectPublishersPhoneEmailPage();
        }
        public static void RunClientImportCompletePage(ref Web AretVal)
        {
            AretVal = new ClientImportCompletePage();
        }
        public static void RunPublicationImportCompletePage(ref Web AretVal)
        {
            AretVal = new PublicationImportCompletePage();
        }
        public static void RunPublishersImportCompletePage(ref Web AretVal)
        {
            AretVal = new PublishersImportCompletePage();
        }

        public static void RunPlanDedupePage(ref Web AretVal)
        {
            AretVal = new PlanDedupe();
        }

        public static void RunPlanConflictPage(ref Web AretVal)
        {
            AretVal = new PlanConflict();
        }

        public static void RunPlanNewPage(ref Web AretVal)
        {
            AretVal = new PlanNewPage();
        }
        public static void RunPlanImportNew(ref Web AretVal)
        {
            AretVal = new PlanImportExcel();
        }

        public static void RunPlanSummaryPage(ref Web AretVal)
        {
            AretVal = new PlanSummaryPage();
        }

        public static void RunPlanDeletePage(ref Web AretVal)
        {
            AretVal = new PlanDeletePage();
        }

        public static void RunPlanSearchPage(ref Web AretVal)
        {
            AretVal = new PlanSearchPage();
        }

        public static void RunPlanCommunicationPage(ref Web AretVal)
        {
            AretVal = new PlanCommunication();
        }

        public static void RunPlanLibraryPage(ref Web AretVal)
        {
            AretVal = new PlanLibraryPage();
        }

        public static void RunPlanTrackingPage(ref Web AretVal)
        {
            AretVal = new PlanTracking();
        }

        public static void RunPlanSummaryPopUpPage(ref Web AretVal)
        {
            AretVal = new PlanSummaryPopPage();
        }

        public static void RunAjaxCall(ref Web AretVal)
        {
            AretVal = new AjaxCallPage();
        }
        public static void RunAjaxCustom(ref Web Aretval)
        {
            Aretval = new RateCardAjaxCallPageCustom();
        }

        public static void RunSectionListPage(ref Web AretVal)
        {
            AretVal = new SectionListPage();
        }

        public static void RunSectionPageNew(ref Web AretVal)
        {
            AretVal = new SectionPageNew();
        }

        public static void RunSectionPageEdit(ref Web AretVal)
        {
            AretVal = new SectionPageEdit();
        }

        public static void RunAjaxDefault(ref Web Aretval)
        {
            Aretval = new RateCardAjaxCallPageDefault();
        }

        public static void RunSectionPageDelete(ref Web AretVal)
        {
            AretVal = new SectionDeletePage();
        }
        public static void RunNabImport(ref Web AretVal)
        {
            AretVal = new importNabsfile();
        }
        public static void RunSubSectionPage(ref Web AretVal)
        {
            AretVal = new SubSectionListPage();
        }

        public static void RunSubSectionPageNew(ref Web AretVal)
        {
            AretVal = new SubSectionPageNew();
        }

        public static void RunSubSectionPageEdit(ref Web AretVal)
        {
            AretVal = new SubSectionsPageEdit();
        }

        public static void RunSubSectionPageDelete(ref Web AretVal)
        {
            AretVal = new SubSectionPageDelete();
        }

        public static void RunPlanPage(ref Web AretVal)
        {
            AretVal = new RateCardListPage();
        }

        public static void RunRateCardImportPage(ref Web AretVal)
        {
            AretVal = new ImportRateCard();
        }

        public static void RunRateCardImportStatusPage(ref Web AretVal)
        {
            AretVal = new RateCardImportStatus();
        }

        public static void RunRateCardEditPage(ref Web AretVal)
        {
            AretVal = new RateCardEditPage();
        }

        public static void RunRateCardDeletePage(ref Web AretVal)
        {
            AretVal = new RateCardPageDelete();
        }

        public static void RunRateCardRedirectorPage(ref Web AretVal)
        {
            AretVal = new PlanRedirector();
        }

        public static void RunPlanSendToSelfPage(ref Web AretVal)
        {
            AretVal = new PlanSendToSelfPage();
        }

        public static void RunPlanSendBooked(ref Web AretVal)
        {
            AretVal = new PlanSendWhenBooked();
        }

        public static void RunPlanSendToAgencyPage(ref Web AretVal)
        {
            AretVal = new PlanSendToAgencyPage();
        }

        public static void RunPlanCopyOldPage(ref Web AretVal)
        {
            AretVal = new PlanCopy_old();
        }
        public static void RunPlanCopyPage(ref Web AretVal)
        {
            AretVal = new PlanCopy();
        }
        public static void RunPlanClosePage(ref Web AretVal)
        {
            AretVal = new PlanClose();
        }

        public static void RunPlanImportPage(ref Web AretVal)
        {
            AretVal = new PlanImport();
        }
        public static void RunPlanImportStatusPage(ref Web AretVal)
        {
            AretVal = new PlanImportStatus();
        }
        public static void RunAjaxSetCommissionOnChangePage(ref Web AretVal)
        {
            AretVal = new AjaxSetCommissionOnChangePage();
        }
        public static void RunPlanNewCopyPage(ref Web AretVal)
        {
            AretVal = new PlanNewCopyPage();
        }//
        public static void RunRateCardAjaxCallPage(ref Web AretVal)
        {
            AretVal = new RateCardAjaxCallPage();
        }
        public static void RunPlanCopyRevisedAjaxCall(ref Web AretVal)
        {
            AretVal = new PlanCopyRevisedAjaxCall();
        }
        public static void RunPlanReffralPage(ref Web AretVal)
        {
            AretVal = new PlanReffralPage();
        }
        public static void RunRateCardManagementPage(ref Web AretVal)
        {
            AretVal = new RateCardManagementPage();
        }
        public static void RunInsertsAjaxPage(ref Web AretVal)
        {
            AretVal = new RateCardAjaxInserts();
        }
        public static void RunInsertsCostAjaxPage(ref Web AretVal)
        {
            AretVal = new RateCardAjaxInsertsCalculation();
        }
        public static void RunPlanBookingConfirmed(ref Web AretVal)
        {
            AretVal = new PlanBookingConfirmed();
        }
        public static void RunReactivatePlanPage(ref Web AretVal)
        {
            AretVal = new ReactivatePlanPage();
        }
    }
}