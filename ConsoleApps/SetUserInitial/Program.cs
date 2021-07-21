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

namespace SetUserInitial
{
    class Program
    {
        static object Get(Entity e, string key)
        {
            if (e.Contains(key))
                return e[key];
            return null;
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

                var fetchXml = @"
<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
  <entity name='systemuser'>
    <attribute name='lastname' />
    <attribute name='firstname' />
    <attribute name='xpd_userinitial' />
    <order attribute='xpd_userinitial' descending='true' />
  </entity>
</fetch>
";
                EntityCollection result = svc.RetrieveMultiple(new FetchExpression(fetchXml));

                if (result != null && result.Entities != null)
                {
                    var fkey = "firstname";
                    var lkey = "lastname";
                    var ikey = "xpd_userinitial";
                    Console.WriteLine("Count: " + result.Entities.Count);
                    var k = 0;
                    foreach (var e in result.Entities)
                    {
                        var firstname = Get(e, fkey)?.ToString() ?? "";
                        var lastname = Get(e, lkey)?.ToString() ?? "";
                        var initial = Get(e, ikey)?.ToString() ?? "";
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(firstname) && !string.IsNullOrWhiteSpace(lastname))
                            {
                                if (char.IsLetter(firstname[0]) && char.IsLetter(lastname[0]))
                                {
                                    initial = (firstname[0].ToString() + lastname[0].ToString()).ToUpper();

                                    // update
                                    Entity ue = new Entity("systemuser");
                                    ue.Id = e.Id;
                                    ue.Attributes.Add("xpd_userinitial", initial);
                                    svc.Update(ue);

                                    Console.WriteLine($"ID: {e.Id} FirstName: {firstname}  LastName: {lastname} Initial: {initial}");
                                    k++;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"***Exception*** ID: {e.Id} FirstName: {firstname}  LastName: {lastname} Initial: {initial}");
                        }
                    }
                    Console.WriteLine("Count(k): " + k);
                }

                Console.WriteLine("Press any key to exit.");
                Console.ReadLine();
            }
        }
    }
}
