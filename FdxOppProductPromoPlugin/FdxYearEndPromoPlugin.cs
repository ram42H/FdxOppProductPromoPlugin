using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxOppProductPromoPlugin
{
    public class FdxYearEndPromoPlugin : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {

            #region Year End Promo Plugin

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            IOrganizationService service = (IOrganizationService)factory.CreateOrganizationService(context.UserId);

            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Variables Section             
            EntityCollection oppProductsList = null;
            EntityReference promoName = null;
            string oppProductListFetchXml = null, term = null;
            Money unadjustedMRRCurr;
            decimal unadjustedMRR = 0, monthlyPromoValue = 0, monthsTillEndDate = 0;
            int termValue = 1, promoCategory = 0, percentageOff = 0, averageDays = 0;
            DateTime endDate = DateTime.Now;
            DateTime actualCloseDate = DateTime.Now;
            TimeSpan timeDiffference;


            // The InputParameters collection contains all the data passed in the message request

            if (!(context.InputParameters.Contains("OpportunityClose") && context.InputParameters["OpportunityClose"] is Entity) && context.MessageName != "Win")
            {
                return;
            }

            //Obtain the target entity from the input parameters.
            Entity entity = (Entity)context.InputParameters["OpportunityClose"];                

            try
            {
                #region Plug-in business logic                    

                Guid opportunityId = ((EntityReference)entity.Attributes["opportunityid"]).Id;

                if (entity.Contains("actualend"))
                {
                    actualCloseDate = (DateTime)entity.Attributes["actualend"];
                }

                oppProductListFetchXml = this.fetchXmlString();

                //Retrieve all opportunity products with year end promo
                oppProductsList = service.RetrieveMultiple(new FetchExpression(string.Format(oppProductListFetchXml, opportunityId)));

                tracingService.Trace("Opportunity Products" + oppProductsList);

                //Update monthly promo value for each opportunity product with year end promo
                foreach (Entity oppProductList in oppProductsList.Entities)
                {

                    //Fields from opportunity entity
                    if (oppProductList.Contains("fdx_unadjustedmrr"))
                    {
                        unadjustedMRRCurr = (Money)oppProductList.Attributes["fdx_unadjustedmrr"];
                        unadjustedMRR = unadjustedMRRCurr.Value;
                    }

                    if (oppProductList.Contains("fdx_term"))
                    {
                        term = oppProductList.FormattedValues["fdx_term"].ToString();
                        termValue = Int32.Parse(term);
                    }

                    if (oppProductList.Contains("fdx_promo"))
                    {
                        promoName = (EntityReference)oppProductList.Attributes["fdx_promo"];
                    }

                    //Fields from promo entity
                    
                    Entity promoEntity = service.Retrieve("fdx_promo", ((EntityReference)oppProductList.Attributes["fdx_promo"]).Id, new ColumnSet("fdx_category", "fdx_percentageoff", "fdx_averagenumberofdaystoactivate", "fdx_enddate"));

                    tracingService.Trace("Promos" + promoEntity);

                    if (promoEntity.Contains("fdx_category"))
                    {
                        promoCategory = ((OptionSetValue)promoEntity.Attributes["fdx_category"]).Value;
                    }

                    if (promoEntity.Contains("fdx_percentageoff"))
                    {
                        percentageOff = (int)promoEntity.Attributes["fdx_percentageoff"];
                    }

                    if (promoEntity.Contains("fdx_averagenumberofdaystoactivate"))
                    {
                        averageDays = (int)promoEntity.Attributes["fdx_averagenumberofdaystoactivate"];
                    }

                    if (promoEntity.Contains("fdx_enddate"))
                    {
                        endDate = (DateTime)promoEntity.Attributes["fdx_enddate"];
                    }
                   
                    timeDiffference = endDate.Date - actualCloseDate.Date;
                    monthsTillEndDate = (((int)timeDiffference.TotalDays) - averageDays) / 31;

                    monthlyPromoValue = (unadjustedMRR * ((percentageOff * monthsTillEndDate) / (termValue * 100)));

                    Entity updateOppProduct = new Entity
                    {
                        LogicalName = "opportunityproduct",
                        Id = oppProductList.Id
                    };

                    updateOppProduct["fdx_monthlypromovalue"] = new Money(monthlyPromoValue);

                    service.Update(updateOppProduct);

                }
                 #endregion

            }

            catch (Exception ex)
            {
                tracingService.Trace("MyPlugin: {0}", ex.ToString());
            }
           
            #endregion
        }


        private string fetchXmlString()
        {
            string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='opportunityproduct'>
                        <attribute name='productid' />
                        <attribute name='productdescription' />
                        <attribute name='priceperunit' />
                        <attribute name='quantity' />
                        <attribute name='opportunityproductid' />
                        <attribute name='fdx_unadjustedmrr' />
                        <attribute name='fdx_term' />
                        <attribute name='fdx_promo' />
                        <attribute name='fdx_monthlypromovalue' />
                        <order attribute='productid' descending='false' />
                        <filter type='and'>
                            <condition attribute='opportunityid' operator='eq' uitype='opportunity' value='{0}' />
                        </filter>
                        <link-entity name='fdx_promo' from='fdx_promoid' to='fdx_promo' alias='an'>
                            <attribute name='fdx_unit' />
                            <attribute name='fdx_percentageoff' />
                            <attribute name='fdx_flatoffamount' />
                            <attribute name='fdx_enddate' />
                            <attribute name='fdx_averagenumberofdaystoactivate' />
                            <attribute name='fdx_category' />
                            <filter type='and'>
                            <condition attribute='fdx_category' operator='eq' value='4' />
                            </filter>
                        </link-entity>
                        </entity>
                    </fetch>"; 

            return fetchXml;

        }

    }
}
