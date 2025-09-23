using DataverseSchemaManager.Models;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;

namespace DataverseSchemaManager.Services
{
    public class DataverseService : IDisposable
    {
        private ServiceClient? _serviceClient;

        public bool Connect(string connectionString)
        {
            try
            {
                _serviceClient = new ServiceClient(connectionString);
                return _serviceClient.IsReady;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error connecting to Dataverse: {ex.Message}");
                return false;
            }
        }

        public void CheckSchemaExists(List<SchemaDefinition> schemas)
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            var groupedByTable = schemas.GroupBy(s => s.TableName.ToLower());

            foreach (var tableGroup in groupedByTable)
            {
                try
                {
                    var request = new RetrieveEntityRequest
                    {
                        LogicalName = tableGroup.Key,
                        EntityFilters = EntityFilters.Attributes
                    };

                    var response = (RetrieveEntityResponse)_serviceClient.Execute(request);
                    var entityMetadata = response.EntityMetadata;

                    var existingAttributes = entityMetadata.Attributes
                        .Where(a => a.LogicalName != null)
                        .Select(a => a.LogicalName.ToLower())
                        .ToHashSet();

                    foreach (var schema in tableGroup)
                    {
                        schema.ExistsInDataverse = existingAttributes.Contains(schema.ColumnName.ToLower());
                    }
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    if (ex.Detail.ErrorCode == -2147185403)
                    {
                        foreach (var schema in tableGroup)
                        {
                            schema.ExistsInDataverse = false;
                            schema.ErrorMessage = $"Table '{tableGroup.Key}' does not exist";
                        }
                    }
                    else
                    {
                        foreach (var schema in tableGroup)
                        {
                            schema.ErrorMessage = ex.Message;
                        }
                    }
                }
                catch (Exception ex)
                {
                    foreach (var schema in tableGroup)
                    {
                        schema.ErrorMessage = ex.Message;
                    }
                }
            }
        }

        public void CreateSchema(List<SchemaDefinition> schemas, string solutionName, string publisherPrefix)
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            var newSchemas = schemas.Where(s => !s.ExistsInDataverse).ToList();
            var groupedByTable = newSchemas.GroupBy(s => s.TableName.ToLower());

            foreach (var tableGroup in groupedByTable)
            {
                foreach (var schema in tableGroup)
                {
                    try
                    {
                        var attributeMetadata = CreateAttributeMetadata(schema, publisherPrefix);

                        var request = new CreateAttributeRequest
                        {
                            EntityName = tableGroup.Key,
                            Attribute = attributeMetadata,
                            SolutionUniqueName = solutionName
                        };

                        _serviceClient.Execute(request);
                        Console.WriteLine($"Created column '{schema.ColumnName}' in table '{schema.TableName}'");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error creating column '{schema.ColumnName}' in table '{schema.TableName}': {ex.Message}");
                        schema.ErrorMessage = ex.Message;
                    }
                }
            }

            try
            {
                var publishRequest = new Microsoft.Crm.Sdk.Messages.PublishAllXmlRequest();
                _serviceClient.Execute(publishRequest);
                Console.WriteLine("Published all customizations");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error publishing customizations: {ex.Message}");
            }
        }

        private AttributeMetadata CreateAttributeMetadata(SchemaDefinition schema, string publisherPrefix)
        {
            var logicalName = $"{publisherPrefix}_{schema.ColumnName.ToLower().Replace(" ", "_")}";
            var displayName = schema.ColumnName;

            AttributeMetadata attributeMetadata = schema.ColumnType.ToLower() switch
            {
                "text" or "string" => new StringAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MaxLength = 100,
                    Format = StringFormat.Text,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label($"Auto-generated column: {displayName}", 1033)
                },
                "number" or "int" or "integer" => new IntegerAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MinValue = int.MinValue,
                    MaxValue = int.MaxValue,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label($"Auto-generated column: {displayName}", 1033)
                },
                "decimal" or "money" or "currency" => new MoneyAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MinValue = -922337203685477.0000,
                    MaxValue = 922337203685477.0000,
                    Precision = 2,
                    PrecisionSource = 2,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label($"Auto-generated column: {displayName}", 1033)
                },
                "date" or "datetime" => new DateTimeAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    Format = DateTimeFormat.DateAndTime,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label($"Auto-generated column: {displayName}", 1033)
                },
                "boolean" or "bool" or "bit" => new BooleanAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    DefaultValue = false,
                    OptionSet = new BooleanOptionSetMetadata(
                        new OptionMetadata(new Label("Yes", 1033), 1),
                        new OptionMetadata(new Label("No", 1033), 0)
                    ),
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label($"Auto-generated column: {displayName}", 1033)
                },
                _ => throw new NotSupportedException($"Column type '{schema.ColumnType}' is not supported")
            };

            return attributeMetadata;
        }

        public List<Entity> GetSolutions()
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            var query = new QueryExpression("solution")
            {
                ColumnSet = new ColumnSet("uniquename", "friendlyname", "publisherid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("ismanaged", ConditionOperator.Equal, false),
                        new ConditionExpression("isvisible", ConditionOperator.Equal, true)
                    }
                }
            };

            var solutions = _serviceClient.RetrieveMultiple(query);
            return solutions.Entities.ToList();
        }

        public string? GetPublisherPrefix(Guid publisherId)
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            var publisher = _serviceClient.Retrieve("publisher", publisherId, new ColumnSet("customizationprefix"));
            return publisher.GetAttributeValue<string>("customizationprefix");
        }

        public void Dispose()
        {
            _serviceClient?.Dispose();
        }
    }
}