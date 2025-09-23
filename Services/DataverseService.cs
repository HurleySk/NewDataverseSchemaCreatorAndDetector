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

                if (!_serviceClient.IsReady)
                {
                    Console.WriteLine($"Connection failed. Error: {_serviceClient.LastError}");
                    if (_serviceClient.LastException != null)
                    {
                        Console.WriteLine($"Exception: {_serviceClient.LastException.Message}");
                        Console.WriteLine($"Stack: {_serviceClient.LastException.StackTrace}");
                    }
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error connecting to Dataverse: {ex.Message}");
                Console.WriteLine($"Stack: {ex.StackTrace}");
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
                        var columnExists = existingAttributes.Contains(schema.ColumnName.ToLower());
                        schema.TableExistsInDataverse = true;
                        schema.ColumnExistsInDataverse = columnExists;
                    }
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    if (ex.Detail.ErrorCode == -2147185403 || ex.Detail.ErrorCode == 31993685)
                    {
                        foreach (var schema in tableGroup)
                        {
                            schema.TableExistsInDataverse = false;
                            schema.ColumnExistsInDataverse = false;
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

            var newSchemas = schemas.Where(s => !s.ColumnExistsInDataverse).ToList();
            var groupedByTable = newSchemas.GroupBy(s => s.TableName.ToLower());

            foreach (var tableGroup in groupedByTable)
            {
                var tableName = tableGroup.Key;
                var hasTableError = tableGroup.First().ErrorMessage?.Contains("does not exist") == true;

                if (hasTableError)
                {
                    try
                    {
                        CreateTable(tableName, publisherPrefix, solutionName);
                        Console.WriteLine($"Created table '{tableName}'");

                        foreach (var schema in tableGroup)
                        {
                            schema.TableExistsInDataverse = true;
                            schema.ErrorMessage = null;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error creating table '{tableName}': {ex.Message}");
                        foreach (var schema in tableGroup)
                        {
                            schema.ErrorMessage = $"Failed to create table: {ex.Message}";
                        }
                        continue;
                    }
                }

                foreach (var schema in tableGroup)
                {
                    try
                    {
                        var attributeMetadata = CreateAttributeMetadata(schema, publisherPrefix);

                        var request = new CreateAttributeRequest
                        {
                            EntityName = tableName,
                            Attribute = attributeMetadata,
                            SolutionUniqueName = solutionName
                        };

                        _serviceClient.Execute(request);
                        schema.ColumnExistsInDataverse = true;
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
            var columnType = schema.ColumnType.Trim().ToLower();

            // Check for unsupported types first
            if (columnType.StartsWith("lookup"))
            {
                throw new NotSupportedException($"Lookup columns cannot be auto-created. Type: '{schema.ColumnType}'. Lookups require relationship metadata and must be created manually or with relationship details.");
            }

            if ((columnType == "choice" || columnType == "picklist") && string.IsNullOrEmpty(schema.ChoiceOptions))
            {
                throw new NotSupportedException($"Choice/Picklist columns require options to be specified in the Choice Options column. Type: '{schema.ColumnType}'. Format: 'Option1;Option2;Option3' or '1:Option1;2:Option2;3:Option3'");
            }

            if (columnType == "customer")
            {
                throw new NotSupportedException($"Customer columns cannot be auto-created. Type: '{schema.ColumnType}'. Customer type is a special polymorphic lookup and must be created manually.");
            }

            // Map supported types
            AttributeMetadata attributeMetadata = columnType switch
            {
                // String types
                "text" or "string" or "single line of text" => new StringAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MaxLength = 100,
                    Format = StringFormat.Text,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label($"Auto-generated column: {displayName}", 1033)
                },

                // Multi-line text
                "multiple lines of text" or "memo" or "multiline" => new MemoAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MaxLength = 2000,
                    Format = StringFormat.Text,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label($"Auto-generated column: {displayName}", 1033)
                },

                // Integer types
                "number" or "int" or "integer" or "whole number" => new IntegerAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MinValue = int.MinValue,
                    MaxValue = int.MaxValue,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label($"Auto-generated column: {displayName}", 1033)
                },

                // Decimal/Money types
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

                // Floating point
                "float" or "double" or "floating point number" => new DoubleAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MinValue = -100000000000.0,
                    MaxValue = 100000000000.0,
                    Precision = 2,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label($"Auto-generated column: {displayName}", 1033)
                },

                // DateTime types
                "date" or "datetime" or "date and time" => new DateTimeAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    Format = DateTimeFormat.DateAndTime,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label($"Auto-generated column: {displayName}", 1033)
                },

                // Date only
                "date only" or "dateonly" => new DateTimeAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    Format = DateTimeFormat.DateOnly,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label($"Auto-generated column: {displayName}", 1033)
                },

                // Boolean types
                "boolean" or "bool" or "bit" or "yes/no" or "y/n" => new BooleanAttributeMetadata
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

                // Choice/Picklist types
                "choice" or "picklist" => CreatePicklistAttribute(schema, logicalName, displayName),

                _ => throw new NotSupportedException($"Column type '{schema.ColumnType}' is not recognized or supported. Supported types: single line of text, multiple lines of text, whole number, floating point number, money, date only, date and time, yes/no, choice/picklist (with options)")
            };

            return attributeMetadata;
        }

        private PicklistAttributeMetadata CreatePicklistAttribute(SchemaDefinition schema, string logicalName, string displayName)
        {
            var options = new List<OptionMetadata>();
            var choiceOptionsText = schema.ChoiceOptions ?? string.Empty;
            var optionParts = choiceOptionsText.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            int defaultValue = 1;
            foreach (var optionPart in optionParts)
            {
                var trimmedOption = optionPart.Trim();
                if (string.IsNullOrEmpty(trimmedOption)) continue;

                int optionValue;
                string optionLabel;

                if (trimmedOption.Contains(':'))
                {
                    var parts = trimmedOption.Split(':', 2);
                    if (int.TryParse(parts[0].Trim(), out optionValue))
                    {
                        optionLabel = parts[1].Trim();
                    }
                    else
                    {
                        optionValue = defaultValue++;
                        optionLabel = trimmedOption;
                    }
                }
                else
                {
                    optionValue = defaultValue++;
                    optionLabel = trimmedOption;
                }

                options.Add(new OptionMetadata(new Label(optionLabel, 1033), optionValue));
            }

            if (options.Count == 0)
            {
                throw new InvalidOperationException($"No valid options found in ChoiceOptions for column '{schema.ColumnName}'. Format: 'Option1;Option2;Option3' or '1:Option1;2:Option2;3:Option3'");
            }

            var optionSet = new OptionSetMetadata
            {
                IsGlobal = false,
                OptionSetType = OptionSetType.Picklist
            };

            foreach (var option in options)
            {
                optionSet.Options.Add(option);
            }

            return new PicklistAttributeMetadata
            {
                SchemaName = logicalName,
                RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                DisplayName = new Label(displayName, 1033),
                Description = new Label($"Auto-generated choice column: {displayName}", 1033),
                OptionSet = optionSet
            };
        }

        private void CreateTable(string tableName, string publisherPrefix, string solutionName)
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            var schemaName = $"{publisherPrefix}_{tableName.ToLower().Replace(" ", "_")}";
            var displayName = tableName;

            var entityMetadata = new EntityMetadata
            {
                SchemaName = schemaName,
                DisplayName = new Label(displayName, 1033),
                DisplayCollectionName = new Label($"{displayName}s", 1033),
                Description = new Label($"Auto-generated table: {displayName}", 1033),
                OwnershipType = OwnershipTypes.UserOwned,
                IsActivity = false,
                HasNotes = false,
                HasActivities = false
            };

            var primaryAttribute = new StringAttributeMetadata
            {
                SchemaName = $"{publisherPrefix}_name",
                RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                MaxLength = 100,
                Format = StringFormat.Text,
                DisplayName = new Label("Name", 1033),
                Description = new Label("Primary name field", 1033)
            };

            var request = new Microsoft.Xrm.Sdk.Messages.CreateEntityRequest
            {
                Entity = entityMetadata,
                PrimaryAttribute = primaryAttribute,
                SolutionUniqueName = solutionName
            };

            _serviceClient.Execute(request);
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