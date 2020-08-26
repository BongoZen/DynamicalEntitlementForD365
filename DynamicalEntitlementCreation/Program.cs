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

namespace DynamicalEntitlementCreation
{
    public class Program: IPlugin
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
                        var customerId = ((EntityReference)(caseEntity.Attributes["customerid"])).Id;
                        var entitlementId = GetEntitlement(service, customerId.ToString());
                        if (entitlementId != null)
                        {
                            Entity caseUpdate = new Entity("incident");
                            caseUpdate.Id = caseEntity.Id;
                            caseUpdate["entitlementid"] = new EntityReference("entitlement", entitlementId.Value);
                            service.Update(caseUpdate);
                        }
                        if(entitlementId == null)
                        {
                            Entity creationEntitlement = new Entity("entitlement");
                            //Guid entityGuid = creationEntitlement.Id;
                            Entity contact = new Entity("contact");
                            Guid contactGuid = (Guid)GetContact(service, customerId.ToString()); 
                            creationEntitlement["name"] = "Case Volume  Tickets For Customer "+customerId.ToString();
                            creationEntitlement["customerid"] = new EntityReference("contact", contactGuid);
                            creationEntitlement["startdate"]=DateTime.Now;
                            //Customers will entitled with addition of sixty days 
                            creationEntitlement["enddate"] = DateTime.Now.AddDays(60);
                            //will create the entitlement of the said customer 
                            service.Create(creationEntitlement);
                            //will simulate time flow , for giving svc.create() some time to actually reflect in cds behind 
                            Thread.Sleep(5000);
                            //Assign new Entitlement to the case field ->entitlement 
                            Entity newCaseUpdate = new Entity("incident");
                            newCaseUpdate.Id = caseEntity.Id;
                            newCaseUpdate["entitlementid"] = new EntityReference("entitlement", creationEntitlement.Id);
                            service.Update(newCaseUpdate);

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
        private Guid? GetContact(IOrganizationService service, string customerId)
        {
            Guid? contactId = null;
            //Retrive the Contact based on the Customer Name
            string fetch3 = @"
                            <fetch mapping='logical'>   
                                <entity name='contact'>  
                                     <attribute name='fullname'/>   

                                <filter>
        <condition attribute='fullname' operator='eq' value='"+customerId+ "' /></filter> </entity></fetch> ";
            EntityCollection contactResult = service.RetrieveMultiple(new FetchExpression(fetch3));
            if (contactResult.Entities.Any())
                contactId = contactResult.Entities.FirstOrDefault().Id;
            return contactId;
        }
    }
}
