using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.Threading;
using System.Management.Instrumentation;

namespace DynamicalEntitlementCreation
{
    public class Program : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //Obtain the tracing service
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            //Obtain the organization service reference which i will need for web service calls   
            IOrganizationServiceFactory serviceFactory =
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            var inputParams = context.InputParameters;
            try
            {

                if (inputParams.Contains("Target") && inputParams["Target"] != null && inputParams["Target"] is Entity)
                {
                    //Obtain the target entity from the input parameters 
                    Entity caseEntity = (Entity)context.InputParameters["Target"];
                    if (caseEntity.Attributes.Contains("customerid"))
                    {
                        Guid customerGuid = ((EntityReference)caseEntity["customerid"]).Id;
                   
                        var customerId = ((EntityReference)(caseEntity.Attributes["customerid"])).Id;
                        var entitlementId = GetEntitlement(service, customerId.ToString());
                        if (entitlementId != null)
                        {
                            Entity caseUpdate = new Entity("incident");
                            caseUpdate.Id = caseEntity.Id;
                            caseUpdate["entitlementid"] = new EntityReference("entitlement", entitlementId.Value);
                            service.Update(caseUpdate);
                        }
                        if (entitlementId == null)
                        {
                            Entity creationEntitlement = new Entity("entitlement");
                            //Guid entityGuid = creationEntitlement.Id;
                            
                            ////Guid contactGuid = (Guid)GetContact(service, customerId.ToString());
                            Guid contactGuid = customerGuid;
                            Entity contact = service.Retrieve("contact", contactGuid, new ColumnSet(true));
                            if (contact.Attributes.Contains("fullname"))
                            {
                                string customerName = Convert.ToString(contact.Attributes["fullname"]);
                               // Guid defaultSlaGuid = new Guid("497392FA-BDEA-EA11-A817-000D3A5694EB");     //default sla guid : Minimal Customer based SLA   *** 
                                
                                //string customerName = retrieveCustomerName(service, contactGuid);
                                creationEntitlement["name"] = "Case Volume  Tickets For Customer " + customerName;

                                creationEntitlement["customerid"] = new EntityReference("contact", contactGuid);
                                creationEntitlement["startdate"] = DateTime.Now;
                                //Customers will be entitled with addition of sixty days 
                                creationEntitlement["enddate"] = DateTime.Now.AddDays(60);
                                //creationEntitlement["slaid"] = new EntityReference("sla", defaultSlaGuid);                           ***
                                //creationEntitlement["totalterms"] = 10.00;      //10 service will be provided to the customer     ***
                                //will create the entitlement of the said customer 

                                Guid entGuid = service.Create(creationEntitlement);
                                Entity entitlement = new Entity("entitlement", entGuid);

                                entitlement["statecode"] = new OptionSetValue(1);     //Status - Activate{auto}
                                entitlement["statuscode"] = new OptionSetValue(1);   //Status reason -active
                                service.Update(entitlement);

                                //will simulate time flow , for giving svc.create() some time to actually reflect in cds behind 
                                //Thread.Sleep(5000);
                                //Assign new Entitlement to the case field ->entitlement 
                                Entity newCaseUpdate = new Entity("incident");
                                newCaseUpdate.Id = caseEntity.Id;
                                newCaseUpdate["entitlementid"] = new EntityReference("entitlement", entitlement.Id);
                                service.Update(newCaseUpdate);
                            }
                        }
                    }
                }
            }

            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("An Error occured in FollowUpPlugin.", ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("FollowUpPlugin: {0}", ex.ToString());
                throw;
            }
        }

        private Guid? GetEntitlement(IOrganizationService service, string customerId)
        {
            Guid? entitlementId = null;
            //Retrieve the Entitlement based on the Customer Id   
            string fetch2 = @"  
               <fetch mapping='logical'>  
                 <entity name='entitlement'>   
                    <attribute name='entitlementid'/>   
                    <attribute name='name'/>   
                    <filter type='and'>
      <condition attribute='customerid' operator='eq' value='" + customerId + "' /></filter> </entity></fetch> ";
            EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetch2));
            if (result.Entities.Any())
                entitlementId = result.Entities.FirstOrDefault().Id;
            return entitlementId;
        }
        //private Guid? GetContact(IOrganizationService service, string customerId)
        //{
        //    Guid? contactId = null;
        //    //Retrive the Contact based on the Customer Name
        //    string fetch3 = @"
        //                    <fetch mapping='logical'>   
        //                        <entity name='contact'>  
        //                             <attribute name='fullname'/>   

        //                        <filter>
        //<condition attribute='fullname' operator='eq' value='" + customerId + "' /></filter> </entity></fetch> ";
        //    EntityCollection contactResult = service.RetrieveMultiple(new FetchExpression(fetch3));
        //    if (contactResult.Entities.Any())
        //        contactId = contactResult.Entities.FirstOrDefault().Id;
        //    return contactId;
        //}

        //public string retrieveCustomerName(IOrganizationService service, Guid customerGuid)
        //{
        //    QueryExpression retrieveName = new QueryExpression()
        //    {
        //        EntityName="contact",
                
        //        ColumnSet=new ColumnSet("fullname")
        //    };
        //    retrieveName.Criteria = new FilterExpression();
        //    retrieveName.Criteria.AddCondition("contactid", ConditionOperator.Equal, customerGuid);
        //    EntityCollection result = service.RetrieveMultiple(retrieveName);
        //    string resultName = Convert.ToString(result.Atrribut)
        //    return "";
        //}
    }
}