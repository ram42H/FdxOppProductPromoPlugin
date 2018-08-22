using Microsoft.Xrm.Sdk;
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
            Money unadjustedMRRCurr, flatOffAmountCurr;
            decimal unadjustedMRR = 0, units = 0, monthlyPromoValue = 0, flatOffAmount = 0;
            int termValue = 1, promoCategory = 0, percentageOff = 0;
            string term = null;
            EntityReference promoName = null;

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

                    if (oppProdEntity.Contains("fdx_unadjustedmrr"))
                    {
                        unadjustedMRRCurr = (Money)oppProdEntity.Attributes["fdx_unadjustedmrr"];
                        unadjustedMRR = unadjustedMRRCurr.Value;
                    }

                    if (oppProdEntity.Contains("fdx_term"))
                    {
                        term = oppProdEntity.FormattedValues["fdx_term"].ToString();
                        termValue = Int32.Parse(term);
                    }

                    if (oppProdEntity.Contains("fdx_promo"))
                    {
                        promoName = (EntityReference)oppProdEntity.Attributes["fdx_promo"];
                    }
                    
                    #region Update promo value based on promo selected on opportunity product

                    Entity updateOppProduct = new Entity
                    {
                        LogicalName = "opportunityproduct",
                        Id = entity.Id
                    };

                    if (promoName != null)
                    {
                        Entity promoEntity = service.Retrieve("fdx_promo", ((EntityReference)oppProdEntity.Attributes["fdx_promo"]).Id, new ColumnSet("fdx_category", "fdx_unit", "fdx_flatoffamount", "fdx_percentageoff")); 

                        tracingService.Trace("Promo GUID" + ((EntityReference)oppProdEntity.Attributes["fdx_promo"]).Id);

                        if (promoEntity.Contains("fdx_category"))
                        {
                            promoCategory = ((OptionSetValue)promoEntity.Attributes["fdx_category"]).Value;
                        }

                        if (promoEntity.Contains("fdx_unit"))
                        {
                            units = (decimal)promoEntity.Attributes["fdx_unit"];
                        }

                        if (promoEntity.Contains("fdx_flatoffamount"))
                        {
                            flatOffAmountCurr = (Money)promoEntity.Attributes["fdx_flatoffamount"];
                            flatOffAmount = flatOffAmountCurr.Value;
                        }

                        if (promoEntity.Contains("fdx_percentageoff"))
                        {
                            percentageOff = (int)promoEntity.Attributes["fdx_percentageoff"];
                        }


                        switch (promoCategory)
                        {

                            //Calculation for promo type : Free time off
                            case 1:
                                monthlyPromoValue = Math.Floor((unadjustedMRR * units) / termValue);                                
                                updateOppProduct["fdx_monthlypromovalue"] = new Money(monthlyPromoValue);
                                break;

                            //Calculation for promo type : Flat Dollar off
                            case 2:
                                monthlyPromoValue = Math.Floor((flatOffAmount * units) / termValue);
                                updateOppProduct["fdx_monthlypromovalue"] = new Money(monthlyPromoValue);
                                break;
                            
                            //Calculation for promo type : Percentage Off
                            case 3:
                                monthlyPromoValue = Math.Floor(((percentageOff * units) / (termValue * 100)) * unadjustedMRR);
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
