using DataverseSchemaManager.Constants;
using DataverseSchemaManager.Interfaces;
using DataverseSchemaManager.Models;
using Microsoft.Extensions.Logging;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Polly;
using Polly.Retry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using System.Threading.Tasks;

namespace DataverseSchemaManager.Services
{
    /// <summary>
    /// Provides operations for interacting with Microsoft Dataverse.
    /// </summary>
    public class DataverseService : IDataverseService
    {
        private ServiceClient? _serviceClient;
        private readonly ILogger<DataverseService> _logger;
        private readonly AsyncRetryPolicy _retryPolicy;

        public DataverseService(ILogger<DataverseService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            // Configure retry policy for transient failures
            _retryPolicy = Policy
                .Handle<FaultException<OrganizationServiceFault>>(ex =>
                    ex.Detail.ErrorCode == DataverseConstants.ErrorCodes.ApiLimitExceeded)
                .WaitAndRetryAsync(
                    DataverseConstants.Api.MaxRetryAttempts,
                    retryAttempt => TimeSpan.FromMilliseconds(
                        DataverseConstants.Api.RetryBaseDelayMs * Math.Pow(2, retryAttempt - 1)),
                    onRetry: (exception, timeSpan, retryCount, context) =>
                    {
                        _logger.LogWarning(
                            "API throttling detected. Retry {RetryCount} of {MaxRetries}. Waiting {DelayMs}ms before retry. Error: {Error}",
                            retryCount,
                            DataverseConstants.Api.MaxRetryAttempts,
                            timeSpan.TotalMilliseconds,
                            exception.Message);
                    });
        }

        /// <inheritdoc/>
        public async Task<bool> ConnectAsync(string connectionString, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogInformation("Attempting to connect to Dataverse...");

                await Task.Run(() =>
                {
                    _serviceClient = new ServiceClient(connectionString);
                }, cancellationToken);

                if (_serviceClient == null || !_serviceClient.IsReady)
                {
                    _logger.LogError(
                        "Connection failed. Error: {Error}",
                        _serviceClient?.LastError ?? "ServiceClient is null");

                    if (_serviceClient?.LastException != null)
                    {
                        _logger.LogError(
                            _serviceClient.LastException,
                            "Connection exception details");
                    }

                    return false;
                }

                _logger.LogInformation("Successfully connected to Dataverse environment: {Url}",
                    _serviceClient.ConnectedOrgFriendlyName ?? "Unknown");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error connecting to Dataverse");
                return false;
            }
        }

        /// <inheritdoc/>
        public async Task CheckSchemaExistsAsync(List<SchemaDefinition> schemas, CancellationToken cancellationToken = default)
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            _logger.LogInformation("Checking existence of {Count} schema definitions across {TableCount} tables",
                schemas.Count,
                schemas.Select(s => s.TableLogicalName.ToLower()).Distinct().Count());

            var groupedByTable = schemas.GroupBy(s => s.TableLogicalName.ToLower());

            foreach (var tableGroup in groupedByTable)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await CheckTableSchemaAsync(tableGroup.ToList(), tableGroup.Key, cancellationToken);
            }

            _logger.LogInformation("Schema check completed");
        }

        private async Task CheckTableSchemaAsync(List<SchemaDefinition> tableSchemas, string tableName, CancellationToken cancellationToken)
        {
            try
            {
                _logger.LogDebug("Checking table '{TableName}' with {ColumnCount} columns", tableName, tableSchemas.Count);

                var request = new RetrieveEntityRequest
                {
                    LogicalName = tableName,
                    EntityFilters = EntityFilters.Attributes
                };

                var response = await _retryPolicy.ExecuteAsync(async () =>
                {
                    return await Task.Run(() =>
                        (RetrieveEntityResponse)_serviceClient!.Execute(request),
                        cancellationToken);
                });

                var entityMetadata = response.EntityMetadata;
                var existingAttributes = entityMetadata.Attributes
                    .Where(a => a.LogicalName != null)
                    .Select(a => a.LogicalName!.ToLower())
                    .ToHashSet();

                foreach (var schema in tableSchemas)
                {
                    var columnExists = existingAttributes.Contains(schema.LogicalName.ToLower());
                    schema.TableExistsInDataverse = true;
                    schema.ColumnExistsInDataverse = columnExists;

                    if (columnExists)
                    {
                        _logger.LogDebug("Column '{Table}.{Column}' exists", tableName, schema.LogicalName);
                    }
                    else
                    {
                        _logger.LogDebug("Column '{Table}.{Column}' does not exist (will be created)", tableName, schema.LogicalName);
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                if (ex.Detail.ErrorCode == DataverseConstants.ErrorCodes.EntityDoesNotExist ||
                    ex.Detail.ErrorCode == DataverseConstants.ErrorCodes.EntityNotFound)
                {
                    _logger.LogWarning("Table '{TableName}' does not exist (will be created)", tableName);

                    foreach (var schema in tableSchemas)
                    {
                        schema.TableExistsInDataverse = false;
                        schema.ColumnExistsInDataverse = false;
                    }
                }
                else
                {
                    _logger.LogError(ex, "Error checking table '{TableName}': {Error}", tableName, ex.Message);

                    foreach (var schema in tableSchemas)
                    {
                        schema.ErrorMessage = ex.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error checking table '{TableName}'", tableName);

                foreach (var schema in tableSchemas)
                {
                    schema.ErrorMessage = ex.Message;
                }
            }
        }

        /// <inheritdoc/>
        public async Task CreateSchemaAsync(List<SchemaDefinition> schemas, string solutionName, string publisherPrefix, CancellationToken cancellationToken = default)
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            var newSchemas = schemas.Where(s => !s.ColumnExistsInDataverse).ToList();
            _logger.LogInformation("Creating {Count} new schemas in solution '{Solution}'", newSchemas.Count, solutionName);

            var groupedByTable = newSchemas.GroupBy(s => s.TableName.ToLower());

            foreach (var tableGroup in groupedByTable)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var tableName = tableGroup.Key;
                var hasTableError = tableGroup.First().ErrorMessage?.Contains("does not exist") == true;

                if (hasTableError)
                {
                    try
                    {
                        var representativeSchema = tableGroup.First();
                        await CreateTableAsync(representativeSchema, publisherPrefix, solutionName, cancellationToken);
                        _logger.LogInformation("Created table '{TableName}'", tableName);

                        foreach (var schema in tableGroup)
                        {
                            schema.TableExistsInDataverse = true;
                            schema.ErrorMessage = null;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error creating table '{TableName}'", tableName);

                        foreach (var schema in tableGroup)
                        {
                            schema.ErrorMessage = $"Failed to create table: {ex.Message}";
                        }
                        continue;
                    }
                }

                foreach (var schema in tableGroup)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    try
                    {
                        await CreateColumnAsync(schema, tableName, publisherPrefix, solutionName, cancellationToken);
                        schema.ColumnExistsInDataverse = true;
                        _logger.LogInformation("Created column '{Table}.{Column}' (Type: {Type})",
                            schema.TableName, schema.ColumnName, schema.ColumnType);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error creating column '{Table}.{Column}'", schema.TableName, schema.ColumnName);
                        schema.ErrorMessage = ex.Message;
                    }
                }
            }

            try
            {
                await PublishAllCustomizationsAsync(cancellationToken);
                _logger.LogInformation("Published all customizations");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error publishing customizations");
            }
        }

        private async Task CreateColumnAsync(SchemaDefinition schema, string tableName, string publisherPrefix, string solutionName, CancellationToken cancellationToken)
        {
            var attributeMetadata = CreateAttributeMetadata(schema, publisherPrefix);
            var columnType = schema.ColumnType.Trim().ToLower();

            // For lookup and customer columns, we need to create the relationship
            if (columnType.StartsWith("lookup"))
            {
                await CreateLookupColumnWithRelationshipAsync(schema, tableName, attributeMetadata as LookupAttributeMetadata, publisherPrefix, solutionName, cancellationToken);
            }
            else if (columnType.StartsWith("customer"))
            {
                await CreateCustomerColumnWithRelationshipAsync(schema, tableName, attributeMetadata as LookupAttributeMetadata, publisherPrefix, solutionName, cancellationToken);
            }
            else
            {
                // Standard column creation
                var request = new CreateAttributeRequest
                {
                    EntityName = tableName,
                    Attribute = attributeMetadata,
                    SolutionUniqueName = solutionName
                };

                await _retryPolicy.ExecuteAsync(async () =>
                {
                    return await Task.Run(() =>
                        _serviceClient!.Execute(request),
                        cancellationToken);
                });
            }
        }

        private async Task CreateLookupColumnWithRelationshipAsync(SchemaDefinition schema, string sourceTable, LookupAttributeMetadata? lookupAttribute, string publisherPrefix, string solutionName, CancellationToken cancellationToken)
        {
            if (lookupAttribute == null || string.IsNullOrWhiteSpace(schema.LookupTargetTable))
            {
                throw new InvalidOperationException("Invalid lookup configuration");
            }

            var targetTable = schema.LookupTargetTable.ToLower();
            var relationshipName = !string.IsNullOrWhiteSpace(schema.LookupRelationshipName)
                ? $"{publisherPrefix}_{schema.LookupRelationshipName.ToLower().Replace(" ", "_")}"
                : $"{publisherPrefix}_{sourceTable}_{schema.LogicalName}";

            var relationship = new OneToManyRelationshipMetadata
            {
                SchemaName = relationshipName,
                ReferencedEntity = targetTable,
                ReferencingEntity = sourceTable.ToLower()
            };

            var request = new CreateOneToManyRequest
            {
                OneToManyRelationship = relationship,
                Lookup = lookupAttribute,
                SolutionUniqueName = solutionName
            };

            await _retryPolicy.ExecuteAsync(async () =>
            {
                return await Task.Run(() =>
                    _serviceClient!.Execute(request),
                    cancellationToken);
            });

            _logger.LogInformation("Created lookup relationship '{RelationshipName}' from '{Source}' to '{Target}'",
                relationshipName, sourceTable, targetTable);
        }

        private async Task CreateCustomerColumnWithRelationshipAsync(SchemaDefinition schema, string sourceTable, LookupAttributeMetadata? customerAttribute, string publisherPrefix, string solutionName, CancellationToken cancellationToken)
        {
            if (customerAttribute == null || string.IsNullOrWhiteSpace(schema.CustomerTargetTables))
            {
                throw new InvalidOperationException("Invalid customer configuration");
            }

            var targetTables = schema.CustomerTargetTables.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            if (targetTables.Length == 0)
            {
                throw new InvalidOperationException($"Customer column '{schema.ColumnName}' has no valid target tables");
            }

            // Create customer lookup with multiple target entities
            // Note: Customer lookups typically target 'account' and 'contact'
            var relationshipName = $"{publisherPrefix}_{sourceTable}_{schema.LogicalName}";

            // For customer columns, we create a polymorphic lookup
            // This requires creating relationships to each target table
            foreach (var targetTable in targetTables)
            {
                var relationship = new OneToManyRelationshipMetadata
                {
                    SchemaName = $"{relationshipName}_{targetTable.ToLower()}",
                    ReferencedEntity = targetTable.ToLower(),
                    ReferencingEntity = sourceTable.ToLower()
                };

                var request = new CreateOneToManyRequest
                {
                    OneToManyRelationship = relationship,
                    Lookup = customerAttribute,
                    SolutionUniqueName = solutionName
                };

                await _retryPolicy.ExecuteAsync(async () =>
                {
                    return await Task.Run(() =>
                        _serviceClient!.Execute(request),
                        cancellationToken);
                });

                _logger.LogInformation("Created customer relationship '{RelationshipName}' from '{Source}' to '{Target}'",
                    relationship.SchemaName, sourceTable, targetTable);
            }
        }

        private AttributeMetadata CreateAttributeMetadata(SchemaDefinition schema, string publisherPrefix)
        {
            // Use required logical name from Excel (no auto-generation)
            var logicalNamePart = schema.LogicalName.ToLower().Replace(" ", "_");
            var logicalName = $"{publisherPrefix}_{logicalNamePart}";

            var displayName = schema.ColumnName;
            var columnType = schema.ColumnType.Trim().ToLower();
            var requiredLevel = ParseRequiredLevel(schema.Required);

            _logger.LogDebug("Creating attribute metadata for '{Column}' with type '{Type}' and logical name '{LogicalName}'",
                schema.ColumnName, columnType, logicalName);

            // Handle lookup types
            if (columnType.StartsWith("lookup"))
            {
                if (string.IsNullOrWhiteSpace(schema.LookupTargetTable))
                {
                    throw new InvalidOperationException($"Lookup column '{schema.ColumnName}' requires LookupTargetTable to be specified in Excel.");
                }
                return CreateLookupAttribute(schema, logicalName, displayName, publisherPrefix, requiredLevel);
            }

            // Handle customer types
            if (columnType.StartsWith("customer"))
            {
                if (string.IsNullOrWhiteSpace(schema.CustomerTargetTables))
                {
                    throw new InvalidOperationException($"Customer column '{schema.ColumnName}' requires CustomerTargetTables to be specified in Excel (comma-separated, e.g., 'account,contact').");
                }
                return CreateCustomerAttribute(schema, logicalName, displayName, publisherPrefix, requiredLevel);
            }

            // Handle choice types
            if ((columnType == "choice" || columnType == "picklist") && string.IsNullOrEmpty(schema.ChoiceOptions))
            {
                throw new InvalidOperationException($"Choice/Picklist column '{schema.ColumnName}' requires options to be specified in the Choice Options column. Format: 'Option1;Option2;Option3' or '1:Option1;2:Option2;3:Option3'");
            }

            // Map supported types
            AttributeMetadata attributeMetadata = columnType switch
            {
                // String types
                "text" or "string" or "single line of text" => new StringAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredLevel),
                    MaxLength = DataverseConstants.AttributeDefaults.TextMaxLength,
                    Format = StringFormat.Text,
                    DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode)
                },

                // Multi-line text
                "multiple lines of text" or "memo" or "multiline" => new MemoAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredLevel),
                    MaxLength = DataverseConstants.AttributeDefaults.MemoMaxLength,
                    Format = StringFormat.Text,
                    DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode)
                },

                // Integer types
                "number" or "int" or "integer" or "whole number" => new IntegerAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredLevel),
                    MinValue = int.MinValue,
                    MaxValue = int.MaxValue,
                    DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode)
                },

                // Decimal/Money types
                "decimal" or "money" or "currency" => new MoneyAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredLevel),
                    MinValue = DataverseConstants.AttributeDefaults.MoneyMinValue,
                    MaxValue = DataverseConstants.AttributeDefaults.MoneyMaxValue,
                    Precision = DataverseConstants.AttributeDefaults.DefaultPrecision,
                    PrecisionSource = DataverseConstants.AttributeDefaults.MoneyPrecisionSource,
                    DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode)
                },

                // Floating point
                "float" or "double" or "floating point number" => new DoubleAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredLevel),
                    MinValue = DataverseConstants.AttributeDefaults.DoubleMinValue,
                    MaxValue = DataverseConstants.AttributeDefaults.DoubleMaxValue,
                    Precision = DataverseConstants.AttributeDefaults.DefaultPrecision,
                    DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode)
                },

                // DateTime types
                "date" or "datetime" or "date and time" => new DateTimeAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredLevel),
                    Format = DateTimeFormat.DateAndTime,
                    DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode)
                },

                // Date only
                "date only" or "dateonly" => new DateTimeAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredLevel),
                    Format = DateTimeFormat.DateOnly,
                    DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode)
                },

                // Boolean types
                "boolean" or "bool" or "bit" or "yes/no" or "y/n" => new BooleanAttributeMetadata
                {
                    SchemaName = logicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredLevel),
                    DefaultValue = false,
                    OptionSet = new BooleanOptionSetMetadata(
                        new OptionMetadata(new Label("Yes", DataverseConstants.Localization.DefaultLanguageCode), 1),
                        new OptionMetadata(new Label("No", DataverseConstants.Localization.DefaultLanguageCode), 0)
                    ),
                    DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode)
                },

                // Choice/Picklist types
                "choice" or "picklist" => CreatePicklistAttribute(schema, logicalName, displayName, requiredLevel),

                _ => throw new NotSupportedException($"Column type '{schema.ColumnType}' is not recognized or supported. Supported types: single line of text, multiple lines of text, whole number, floating point number, money, date only, date and time, yes/no, choice/picklist (with options)")
            };

            // Set description only if user provided one
            if (!string.IsNullOrWhiteSpace(schema.Description))
            {
                attributeMetadata.Description = new Label(schema.Description, DataverseConstants.Localization.DefaultLanguageCode);
            }

            return attributeMetadata;
        }

        private AttributeRequiredLevel ParseRequiredLevel(string? requiredValue)
        {
            if (string.IsNullOrWhiteSpace(requiredValue))
            {
                return AttributeRequiredLevel.None;
            }

            return requiredValue.Trim().ToLower() switch
            {
                "none" or "optional" => AttributeRequiredLevel.None,
                "required" or "business required" => AttributeRequiredLevel.ApplicationRequired,
                "recommended" or "business recommended" => AttributeRequiredLevel.Recommended,
                "system required" => AttributeRequiredLevel.SystemRequired,
                _ => AttributeRequiredLevel.None
            };
        }

        private PicklistAttributeMetadata CreatePicklistAttribute(SchemaDefinition schema, string logicalName, string displayName, AttributeRequiredLevel requiredLevel)
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

                options.Add(new OptionMetadata(new Label(optionLabel, DataverseConstants.Localization.DefaultLanguageCode), optionValue));
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
                RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredLevel),
                DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode),
                OptionSet = optionSet
            };
        }

        private LookupAttributeMetadata CreateLookupAttribute(SchemaDefinition schema, string logicalName, string displayName, string publisherPrefix, AttributeRequiredLevel requiredLevel)
        {
            if (string.IsNullOrWhiteSpace(schema.LookupTargetTable))
            {
                throw new ArgumentException("LookupTargetTable is required for lookup columns", nameof(schema));
            }

            var lookupAttribute = new LookupAttributeMetadata
            {
                SchemaName = logicalName,
                RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredLevel),
                DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode)
            };

            _logger.LogInformation("Created lookup attribute '{LogicalName}' targeting table '{TargetTable}'",
                logicalName, schema.LookupTargetTable);

            return lookupAttribute;
        }

        private LookupAttributeMetadata CreateCustomerAttribute(SchemaDefinition schema, string logicalName, string displayName, string publisherPrefix, AttributeRequiredLevel requiredLevel)
        {
            if (string.IsNullOrWhiteSpace(schema.CustomerTargetTables))
            {
                throw new ArgumentException("CustomerTargetTables is required for customer columns", nameof(schema));
            }

            // Customer is a special lookup that can reference multiple entity types (account, contact, etc.)
            var customerAttribute = new LookupAttributeMetadata
            {
                SchemaName = logicalName,
                RequiredLevel = new AttributeRequiredLevelManagedProperty(requiredLevel),
                DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode)
            };

            _logger.LogInformation("Created customer attribute '{LogicalName}' targeting tables '{TargetTables}'",
                logicalName, schema.CustomerTargetTables);

            return customerAttribute;
        }

        private async Task CreateTableAsync(SchemaDefinition schema, string publisherPrefix, string solutionName, CancellationToken cancellationToken)
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            var tableName = schema.TableName;

            // Use required table logical name from Excel (no auto-generation)
            var logicalNamePart = schema.TableLogicalName.ToLower().Replace(" ", "_");
            var schemaName = $"{publisherPrefix}_{logicalNamePart}";

            var displayName = tableName;

            // Use explicit display collection name if provided, otherwise auto-generate with "s"
            var displayCollectionName = !string.IsNullOrWhiteSpace(schema.TableDisplayCollectionName)
                ? schema.TableDisplayCollectionName
                : $"{displayName}s";

            _logger.LogInformation("Creating new table '{TableName}' with schema name '{SchemaName}' and collection name '{CollectionName}'",
                displayName, schemaName, displayCollectionName);

            var entityMetadata = new EntityMetadata
            {
                SchemaName = schemaName,
                DisplayName = new Label(displayName, DataverseConstants.Localization.DefaultLanguageCode),
                DisplayCollectionName = new Label(displayCollectionName, DataverseConstants.Localization.DefaultLanguageCode),
                OwnershipType = OwnershipTypes.UserOwned,
                IsActivity = false,
                HasNotes = false,
                HasActivities = false
            };

            var primaryAttribute = new StringAttributeMetadata
            {
                SchemaName = $"{publisherPrefix}_name",
                RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                MaxLength = DataverseConstants.AttributeDefaults.TextMaxLength,
                Format = StringFormat.Text,
                DisplayName = new Label("Name", DataverseConstants.Localization.DefaultLanguageCode)
            };

            var request = new Microsoft.Xrm.Sdk.Messages.CreateEntityRequest
            {
                Entity = entityMetadata,
                PrimaryAttribute = primaryAttribute,
                SolutionUniqueName = solutionName
            };

            await _retryPolicy.ExecuteAsync(async () =>
            {
                return await Task.Run(() =>
                    _serviceClient!.Execute(request),
                    cancellationToken);
            });
        }

        /// <inheritdoc/>
        public async Task<List<Entity>> GetSolutionsAsync(CancellationToken cancellationToken = default)
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            _logger.LogDebug("Retrieving unmanaged solutions from Dataverse");

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

            var solutions = await Task.Run(() =>
                _serviceClient.RetrieveMultiple(query),
                cancellationToken);

            _logger.LogInformation("Retrieved {Count} unmanaged solutions", solutions.Entities.Count);

            return solutions.Entities.ToList();
        }

        /// <inheritdoc/>
        public async Task<string?> GetPublisherPrefixAsync(Guid publisherId, CancellationToken cancellationToken = default)
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            _logger.LogDebug("Retrieving publisher prefix for publisher ID: {PublisherId}", publisherId);

            var publisher = await Task.Run(() =>
                _serviceClient.Retrieve("publisher", publisherId, new ColumnSet("customizationprefix")),
                cancellationToken);

            var prefix = publisher.GetAttributeValue<string>("customizationprefix");

            _logger.LogDebug("Publisher prefix: {Prefix}", prefix);

            return prefix;
        }

        private async Task PublishAllCustomizationsAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Publishing all customizations...");

            var publishRequest = new Microsoft.Crm.Sdk.Messages.PublishAllXmlRequest();

            await _retryPolicy.ExecuteAsync(async () =>
            {
                return await Task.Run(() =>
                    _serviceClient!.Execute(publishRequest),
                    cancellationToken);
            });
        }

        /// <summary>
        /// Validates a logical name against Dataverse naming rules.
        /// </summary>
        public (bool IsValid, string? ErrorMessage) ValidateLogicalName(string logicalName)
        {
            if (string.IsNullOrWhiteSpace(logicalName))
            {
                return (false, "Logical name cannot be empty");
            }

            var trimmed = logicalName.Trim();

            // Must contain only lowercase alphanumeric and underscores
            if (!System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^[a-z0-9_]+$"))
            {
                return (false, $"Logical name '{trimmed}' contains invalid characters. Only lowercase letters, numbers, and underscores allowed.");
            }

            // Cannot start with underscore
            if (trimmed.StartsWith("_"))
            {
                return (false, $"Logical name '{trimmed}' cannot start with underscore");
            }

            // Cannot end with underscore
            if (trimmed.EndsWith("_"))
            {
                return (false, $"Logical name '{trimmed}' cannot end with underscore");
            }

            // Check length (Dataverse limit is typically 50 for the logical part)
            if (trimmed.Length > 50)
            {
                return (false, $"Logical name '{trimmed}' exceeds maximum length of 50 characters");
            }

            return (true, null);
        }

        /// <summary>
        /// Checks if a column schema name already exists in a table.
        /// </summary>
        public async Task<bool> ColumnExistsAsync(string tableName, string columnSchemaName, CancellationToken cancellationToken = default)
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            try
            {
                var request = new RetrieveEntityRequest
                {
                    LogicalName = tableName.ToLower(),
                    EntityFilters = EntityFilters.Attributes
                };

                var response = await Task.Run(() =>
                    (RetrieveEntityResponse)_serviceClient.Execute(request),
                    cancellationToken);

                return response.EntityMetadata.Attributes
                    .Any(a => a.LogicalName?.Equals(columnSchemaName.ToLower(), StringComparison.OrdinalIgnoreCase) == true);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                if (ex.Detail.ErrorCode == DataverseConstants.ErrorCodes.EntityDoesNotExist ||
                    ex.Detail.ErrorCode == DataverseConstants.ErrorCodes.EntityNotFound)
                {
                    // Table doesn't exist, so column doesn't exist
                    return false;
                }
                throw;
            }
        }

        /// <summary>
        /// Checks if a table schema name already exists.
        /// </summary>
        public async Task<bool> TableExistsAsync(string tableSchemaName, CancellationToken cancellationToken = default)
        {
            if (_serviceClient == null || !_serviceClient.IsReady)
            {
                throw new InvalidOperationException("Not connected to Dataverse");
            }

            try
            {
                var request = new RetrieveEntityRequest
                {
                    LogicalName = tableSchemaName.ToLower(),
                    EntityFilters = EntityFilters.Entity
                };

                await Task.Run(() =>
                    _serviceClient.Execute(request),
                    cancellationToken);

                return true;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                if (ex.Detail.ErrorCode == DataverseConstants.ErrorCodes.EntityDoesNotExist ||
                    ex.Detail.ErrorCode == DataverseConstants.ErrorCodes.EntityNotFound)
                {
                    return false;
                }
                throw;
            }
        }

        /// <summary>
        /// Validates that a logical name doesn't already start with a prefix.
        /// </summary>
        public (bool IsValid, string? ErrorMessage) ValidatePrefixMismatch(string? explicitLogicalName, string publisherPrefix)
        {
            if (string.IsNullOrWhiteSpace(explicitLogicalName))
            {
                return (true, null); // No explicit name, will use auto-generated
            }

            var trimmed = explicitLogicalName.Trim().ToLower();

            // Check if it starts with a different prefix (containing underscore suggests it might have a prefix)
            if (trimmed.Contains("_"))
            {
                var parts = trimmed.Split('_', 2);
                var potentialPrefix = parts[0];

                // If the potential prefix doesn't match the publisher prefix, warn user
                if (!potentialPrefix.Equals(publisherPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    return (false, $"Logical name '{trimmed}' appears to contain prefix '{potentialPrefix}' which doesn't match publisher prefix '{publisherPrefix}'. Remove the prefix from the logical name - it will be added automatically.");
                }
            }

            return (true, null);
        }

        public void Dispose()
        {
            _logger.LogDebug("Disposing DataverseService");
            _serviceClient?.Dispose();
        }
    }
}
