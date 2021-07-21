using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SetDealTeam
{
    class Program
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

        static object Get(Entity e, string key)
        {
            if (e.Contains(key))
                return e[key];
            return null;
        }

        static string GetDealTeams(IOrganizationService service, Guid dealId)
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
            //Console.WriteLine("SetDealTeam:GetDealTeams Fetch XML: {0}", fetchXml);
            EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
            var text = "";
            if (result != null && result.Entities != null)
            {
                var key = "usr." + UserInitialColumnName;
                var list = new List<string>();
                //Console.WriteLine("SetDealTeam:GetDealTeams List Count: {0}, looking for key: {1}", result.Entities.Count, key);
                foreach (var e in result.Entities)
                {
                    //Console.WriteLine("SetDealTeam:GetDealTeams checking for delete, e.id={0}", e.Id);

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
                //Console.WriteLine("SetDealTeam:GetDealTeams Text: {0}", text);
            }
            return text;
        }


        static void Main(string[] args)
        {
            var conn = ConfigurationManager.ConnectionStrings["hannover"]?.ConnectionString;

            using (var devConnection = new CrmServiceClient(conn))
            {
                //Create the IOrganizationService:
                var svc = (IOrganizationService)devConnection.OrganizationWebProxyClient != null ?

                (IOrganizationService)devConnection.OrganizationWebProxyClient :

                (IOrganizationService)devConnection.OrganizationServiceProxy;


                WhoAmIRequest request = new WhoAmIRequest();

                WhoAmIResponse response = (WhoAmIResponse)svc.Execute(request);

                Console.WriteLine("Your UserId is {0}", response.UserId);

                var fetchXml = $@"
<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
  <entity name='{DealLogicalName}'>
    <attribute name='{DealTeamColumnName}' />
  </entity>
</fetch>
";
                EntityCollection result = svc.RetrieveMultiple(new FetchExpression(fetchXml));

                if (result != null && result.Entities != null)
                {
                    Console.WriteLine("Count: " + result.Entities.Count);
                    var k = 0;
                    foreach (var e in result.Entities)
                    {
                        var text = GetDealTeams(svc, e.Id);
                        if (string.IsNullOrWhiteSpace(text)) 
                            continue;

                        try
                        {
                            // update
                            Entity ue = new Entity(DealLogicalName);
                            ue.Id = e.Id;
                            ue.Attributes.Add(DealTeamColumnName, text);
                            svc.Update(ue);

                            Console.WriteLine($"{k} Updated Id = {e.Id}, DealTeam: {text}");
                            k++;
                            //if (k == 10) break;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"***Exception*** ID: {e.Id} {ex.Message}");
                        }
                    }
                    Console.WriteLine("Count(k): " + k);
                }

                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
        }
    }
}
