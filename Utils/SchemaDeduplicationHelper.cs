using DataverseSchemaManager.Models;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DataverseSchemaManager.Utils
{
    /// <summary>
    /// Provides deduplication logic for schema definitions to handle duplicate entries in input files.
    /// </summary>
    public static class SchemaDeduplicationHelper
    {
        /// <summary>
        /// Deduplicates a list of schema definitions based on table logical name and column logical name.
        /// First occurrence wins. Logs warnings for duplicates with conflicting data.
        /// </summary>
        /// <param name="schemas">The list of schemas to deduplicate</param>
        /// <param name="logger">Logger for reporting duplicates and conflicts</param>
        /// <param name="sourceContext">Context string for logging (e.g., "Excel", "CSV", "creation")</param>
        /// <returns>Deduplicated list of schemas</returns>
        public static List<SchemaDefinition> DeduplicateSchemas(
            List<SchemaDefinition> schemas,
            ILogger logger,
            string sourceContext = "input")
        {
            if (schemas == null || schemas.Count == 0)
            {
                return schemas ?? new List<SchemaDefinition>();
            }

            var seen = new Dictionary<string, (SchemaDefinition schema, int index)>();
            var result = new List<SchemaDefinition>();
            int duplicateCount = 0;
            int conflictCount = 0;

            for (int i = 0; i < schemas.Count; i++)
            {
                var schema = schemas[i];

                // Skip schemas with missing required fields
                if (string.IsNullOrWhiteSpace(schema.TableLogicalName) ||
                    string.IsNullOrWhiteSpace(schema.LogicalName))
                {
                    continue;
                }

                var key = CreateKey(schema.TableLogicalName, schema.LogicalName);

                if (!seen.ContainsKey(key))
                {
                    seen[key] = (schema, i);
                    result.Add(schema);
                }
                else
                {
                    duplicateCount++;
                    var (firstSchema, firstIndex) = seen[key];

                    // Detect conflicting data
                    var conflicts = DetectConflicts(firstSchema, schema);

                    if (conflicts.Any())
                    {
                        conflictCount++;
                        logger.LogWarning(
                            "Duplicate schema with conflicting data found in {Source} at index {CurrentIndex}: " +
                            "Table='{Table}', Column='{Column}' (first seen at index {FirstIndex}). " +
                            "Keeping first occurrence. Conflicts: {Conflicts}",
                            sourceContext,
                            i + 1,
                            schema.TableLogicalName,
                            schema.LogicalName,
                            firstIndex + 1,
                            string.Join("; ", conflicts));
                    }
                    else
                    {
                        logger.LogInformation(
                            "Exact duplicate schema found in {Source} at index {CurrentIndex}: " +
                            "Table='{Table}', Column='{Column}' (first seen at index {FirstIndex}). Skipping duplicate.",
                            sourceContext,
                            i + 1,
                            schema.TableLogicalName,
                            schema.LogicalName,
                            firstIndex + 1);
                    }
                }
            }

            if (duplicateCount > 0)
            {
                logger.LogInformation(
                    "Deduplication complete for {Source}: Removed {DuplicateCount} duplicate(s) " +
                    "({ConflictCount} with conflicting data, {ExactCount} exact match(es)). " +
                    "Kept {ResultCount} unique schema(s).",
                    sourceContext,
                    duplicateCount,
                    conflictCount,
                    duplicateCount - conflictCount,
                    result.Count);
            }

            return result;
        }

        /// <summary>
        /// Creates a unique key for schema identification based on table and column logical names.
        /// </summary>
        private static string CreateKey(string tableLogicalName, string logicalName)
        {
            return $"{tableLogicalName.ToLower().Trim()}|{logicalName.ToLower().Trim()}";
        }

        /// <summary>
        /// Detects conflicting data between two schemas that have the same table+column logical names.
        /// </summary>
        private static List<string> DetectConflicts(SchemaDefinition first, SchemaDefinition second)
        {
            var conflicts = new List<string>();

            // Check ColumnType
            if (HasConflict(first.ColumnType, second.ColumnType))
            {
                conflicts.Add($"ColumnType: '{first.ColumnType}' vs '{second.ColumnType}'");
            }

            // Check TableName (display)
            if (HasConflict(first.TableName, second.TableName))
            {
                conflicts.Add($"TableName: '{first.TableName}' vs '{second.TableName}'");
            }

            // Check ColumnName (display)
            if (HasConflict(first.ColumnName, second.ColumnName))
            {
                conflicts.Add($"ColumnName: '{first.ColumnName}' vs '{second.ColumnName}'");
            }

            // Check ChoiceOptions
            if (HasConflict(first.ChoiceOptions, second.ChoiceOptions))
            {
                conflicts.Add($"ChoiceOptions: '{first.ChoiceOptions}' vs '{second.ChoiceOptions}'");
            }

            // Check LookupTargetTable
            if (HasConflict(first.LookupTargetTable, second.LookupTargetTable))
            {
                conflicts.Add($"LookupTargetTable: '{first.LookupTargetTable}' vs '{second.LookupTargetTable}'");
            }

            // Check CustomerTargetTables
            if (HasConflict(first.CustomerTargetTables, second.CustomerTargetTables))
            {
                conflicts.Add($"CustomerTargetTables: '{first.CustomerTargetTables}' vs '{second.CustomerTargetTables}'");
            }

            // Check Required level
            if (HasConflict(first.Required, second.Required))
            {
                conflicts.Add($"Required: '{first.Required}' vs '{second.Required}'");
            }

            // Check Description (informational, not critical)
            if (HasConflict(first.Description, second.Description))
            {
                conflicts.Add($"Description: '{TruncateForDisplay(first.Description)}' vs '{TruncateForDisplay(second.Description)}'");
            }

            return conflicts;
        }

        /// <summary>
        /// Checks if two string values conflict (both non-empty but different).
        /// </summary>
        private static bool HasConflict(string? value1, string? value2)
        {
            // No conflict if either is null/empty
            if (string.IsNullOrWhiteSpace(value1) || string.IsNullOrWhiteSpace(value2))
            {
                return false;
            }

            // Conflict if both have values but they're different (case-insensitive)
            return !value1.Trim().Equals(value2.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Truncates a string for display in logs to avoid excessive output.
        /// </summary>
        private static string TruncateForDisplay(string? value, int maxLength = 50)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            if (value.Length <= maxLength)
            {
                return value;
            }

            return value.Substring(0, maxLength) + "...";
        }
    }
}
