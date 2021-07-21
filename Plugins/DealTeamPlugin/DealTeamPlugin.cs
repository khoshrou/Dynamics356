using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System;
using System.Linq;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;

namespace DealTeamPlugin
{
    public class DealTeamPlugin : IPlugin
    {
        const string DealLogicalName = "tspe_deal";
        const string DealTeamLogicalName = "tspe_fee";
        const string DealTeamIDColumnName = "tspe_feeid";
        const string DealIDColumnName = "tspe_dealid";
        const string DealLookupColumnName = "tspe_dealid";
        const string DealMemberTypeColumnName = "xpd_dealmembertype";
        const string UserLookupColumnName = "xpd_user";
        const string DealTeamColumnName = "xpd_dealteam";
        const string UserInitialColumnName = "xpd_userinitial";

        const int InternalMember = 930580001;

        private string GetDealTeams(ITracingService tracingServic, IOrganizationService service, Guid dealId, Guid? dealTeamId = null)
        {
            var fetchXml = $@"
<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
  <entity name='{DealTeamLogicalName}'>
    <attribute name='{DealMemberTypeColumnName}' />
    <attribute name='{DealTeamIDColumnName}' />
    <order attribute='{DealMemberTypeColumnName}' descending='false' />
    <filter type='and'>
      <condition attribute='{DealMemberTypeColumnName}' operator='eq' value='{InternalMember}' />
      <condition attribute='{DealLookupColumnName}' operator='eq' value='{{{dealId}}}' />
    </filter>
    <link-entity name='systemuser' from='systemuserid' to='{UserLookupColumnName}' visible='false' link-type='outer' alias='usr'>
      <attribute name='firstname' />
      <attribute name='lastname' />
      <attribute name='{UserInitialColumnName}' />
    </link-entity>
  </entity>
</fetch>
";
            tracingServic.Trace("DealTeamPlugin:GetDealTeams Fetch XML: {0}", fetchXml);
            EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
            var text = "";
            if (result != null && result.Entities != null)
            {
                var key = "usr." + UserInitialColumnName;
                var list = new List<string>();
                tracingServic.Trace("DealTeamPlugin:GetDealTeams List Count: {0}, looking for key: {1}", result.Entities.Count, key);
                foreach (var e in result.Entities)
                {
                    tracingServic.Trace("DealTeamPlugin:GetDealTeams checking for delete, e.id={0}, deal team id={1}", e.Id, (dealTeamId != null && dealTeamId.HasValue) ? dealTeamId.Value.ToString() : "");
                    if (dealTeamId != null && dealTeamId.HasValue && e.Id == dealTeamId.Value)
                        continue;

                    if (e.Contains(key))
                    {
                        var v = ((AliasedValue)e.Attributes[key])?.Value?.ToString();
                        if (!list.Any(x => x == v))
                        {
                            text += (text.Length > 0 ? ", " : "") + v;
                            list.Add(v);
                        }
                    }
                }
                tracingServic.Trace("DealTeamPlugin:GetDealTeams Text: {0}", text);
            }
            return text;
        }

        private void ExecuteForEntity(ITracingService tracingService, IOrganizationService service, Guid dealteamid, bool delete = false)
        {
            var cols = new ColumnSet(new string[] { DealLookupColumnName });
            tracingService.Trace("DealTeamPlugin:ExecuteForEntity Before retriving dealteam");

            Entity dealTeamEntity = service.Retrieve(DealTeamLogicalName, dealteamid, cols);

            if (dealTeamEntity != null)
            {
                if (!dealTeamEntity.Contains(DealLookupColumnName))
                {
                    tracingService.Trace("DealTeamPlugin:ExecuteForEntity dealteam entity does not contain {0}", DealLookupColumnName);
                    return;
                }

                tracingService.Trace("DealTeamPlugin:ExecuteForEntity dealteam entity is OK");
                EntityReference dealRef = (EntityReference)dealTeamEntity[DealLookupColumnName];
                if (dealRef != null)
                {
                    tracingService.Trace("DealTeamPlugin:ExecuteForEntity dealRef is OK");
                    var dealTeamsText = GetDealTeams(tracingService, service, dealRef.Id, delete ? dealteamid : (Guid?)null);

                    tracingService.Trace("DealTeamPlugin:ExecuteForEntity seting up entity to update");
                    Entity deal = new Entity(DealLogicalName);
                    deal.Id = dealRef.Id;
                    deal.Attributes.Add(DealTeamColumnName, dealTeamsText);

                    tracingService.Trace("DealTeamPlugin:ExecuteForEntity tring to update");
                    service.Update(deal);
                    tracingService.Trace("DealTeamPlugin:ExecuteForEntity deal updated, id: {0}, text: {1}", dealRef.Id, dealTeamsText);
                }
                else
                {
                    tracingService.Trace("DealTeamPlugin:ExecuteForEntity dealRef is NULL");
                }
            }
            else
            {
                tracingService.Trace("DealTeamPlugin:ExecuteForEntity dealteam entity is NULL");
            }
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("DealTeamPlugin: Starting execution");

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                tracingService.Trace("DealTeamPlugin: Target is entity");
                // Obtain the target entity from the input parameters.  
                Entity entity = (Entity)context.InputParameters["Target"];

                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {

                    // check entity is not null
                    if (entity == null)
                    {
                        tracingService.Trace("DealTeamPlugin: Entity is null");
                        //throw new InvalidPluginExecutionException("DealTeamPlugin: Entity is null");
                        return;
                    }

                    // check its a dealteam entity
                    if (entity.LogicalName != DealTeamLogicalName)
                    {
                        tracingService.Trace("DealTeamPlugin: Entity logical name is not {0}, it is {1}", DealTeamLogicalName, entity.LogicalName);
                        //throw new InvalidPluginExecutionException("DealTeamPlugin: Entity logical name is not dealteam");
                        return;
                    }

                    // check if event is create or update
                    if (context.MessageName == "Create" || context.MessageName == "Update")
                    {
                        tracingService.Trace("DealTeamPlugin: Its a Create or Update, Message Name: {0}", context.MessageName);
                        if (entity.Contains(DealTeamIDColumnName))
                        {
                            
                            tracingService.Trace("DealTeamPlugin: {0} EXISTS.", DealTeamIDColumnName);
                            Guid dealteamid = entity.Id;// new Guid(entity[DealTeamIDColumnName].ToString());
                            ExecuteForEntity(tracingService, service, dealteamid);
                        }
                        else
                        {
                            tracingService.Trace("DealTeamPlugin: Entity missing dealteamid");
                            //throw new InvalidPluginExecutionException("DealTeamPlugin: Entity missing dealid");
                            return;
                        }

                        //service.Retrieve()
                        //create logic
                        //EntityReference entity = (EntityReference)entity["tspe_dealid"];
                    }
                    else
                    {
                        //update logic
                        tracingService.Trace("DealTeamPlugin: message name is not create or update, it is {0}", context.MessageName);
                        throw new InvalidPluginExecutionException("DealTeamPlugin: Entity is null");
                    }
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in DealTeamPlugin.", ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("DealTeamPlugin Exception: {0}", ex.ToString());
                    throw;
                }
            }
            else if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is EntityReference)
            {
                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                // Obtain the target entity from the input parameters.
                EntityReference entityRef = (EntityReference)context.InputParameters["Target"];
                ExecuteForEntity(tracingService, service, entityRef.Id, true);
            }
        }
    }
}
