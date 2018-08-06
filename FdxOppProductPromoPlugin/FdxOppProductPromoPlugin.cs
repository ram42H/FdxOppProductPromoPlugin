﻿using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxOppProductPromoPlugin
{
    public class FdxOppProductPromoPlugin : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {

            #region Promo Plugin

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            IOrganizationService service = (IOrganizationService)factory.CreateOrganizationService(context.UserId);

            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            /* Variables section 
            Entity: Opportunity Product, Fields: Unadjusted MRR, Monthly Promo Value, Term
            Entity: Promo, Fields: Promo Category, units
            */
            Money unadjustedMRRCurr;
            decimal unadjustedMRR = 0, units = 0, monthlyPromoValue = 0;
            int termValue = 1, promoCategory = 0;
            string term = "";

            // The InputParameters collection contains all the data passed in the message request
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                //Obtain the target entity from the input parameters.
                Entity entity = (Entity)context.InputParameters["Target"];

                //Verify that the target entity represents an entity type you are expecting.
                //For example, an account. If not the plug-in was not registered correctly.
                if (entity.LogicalName != "opportunityproduct")
                    return;

                try
                {
                    #region Plug-in business logic

                    Entity oppProdEntity = new Entity();

                    oppProdEntity = service.Retrieve("opportunityproduct", entity.Id, new ColumnSet("fdx_unadjustedmrr", "fdx_term", "fdx_promo"));                    

                    unadjustedMRRCurr = (Money)oppProdEntity.Attributes["fdx_unadjustedmrr"];
                    unadjustedMRR = unadjustedMRRCurr.Value;                    

                    term = oppProdEntity.FormattedValues["fdx_term"].ToString();
                    termValue = Int32.Parse(term);

                    EntityReference promoName = (EntityReference)oppProdEntity.Attributes["fdx_promo"];

                    #region Update promo value based on promo selected on opportunity product

                    Entity updateOppProduct = new Entity
                    {
                        LogicalName = "opportunityproduct",
                        Id = entity.Id
                    };

                    if (promoName != null)
                    {
                        Entity promoEntity = service.Retrieve("fdx_promo", ((EntityReference)oppProdEntity.Attributes["fdx_promo"]).Id, new ColumnSet("fdx_category", "fdx_unit"));

                        tracingService.Trace("Promo GUID" + ((EntityReference)oppProdEntity.Attributes["fdx_promo"]).Id);

                        if (promoEntity.Contains("fdx_category"))
                        {
                            promoCategory = ((OptionSetValue)promoEntity.Attributes["fdx_category"]).Value;
                        }

                        if (promoEntity.Contains("fdx_unit"))
                        {
                            units = (decimal)promoEntity.Attributes["fdx_unit"];
                        }


                        switch (promoCategory)
                        {

                            //Calculation for promo type : Free time off
                            case 1:
                                monthlyPromoValue = (unadjustedMRR * units) / termValue;
                                updateOppProduct["fdx_monthlypromovalue"] = new Money(monthlyPromoValue);
                                break;

                            //Default value of monthly promo value field, N/A option also hits here
                            default:
                                monthlyPromoValue = 0;
                                updateOppProduct["fdx_monthlypromovalue"] = new Money(monthlyPromoValue);
                                break;

                        }

                        service.Update(updateOppProduct);

                        #endregion                        

                        #endregion

                    }

                }

                catch (Exception ex)
                {
                    tracingService.Trace("MyPlugin: {0}", ex.ToString());
                }


            }

            #endregion

        }

    }
}
