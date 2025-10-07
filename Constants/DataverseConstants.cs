namespace DataverseSchemaManager.Constants
{
    /// <summary>
    /// Contains constant values used throughout the Dataverse schema management application.
    /// </summary>
    public static class DataverseConstants
    {
        /// <summary>
        /// Error codes related to Dataverse operations.
        /// </summary>
        public static class ErrorCodes
        {
            /// <summary>
            /// Error code indicating that an entity (table) does not exist in Dataverse.
            /// </summary>
            public const int EntityDoesNotExist = -2147185403;

            /// <summary>
            /// Alternative error code for entity not found.
            /// </summary>
            public const int EntityNotFound = 31993685;

            /// <summary>
            /// Error code for entity not found (ObjectDoesNotExist).
            /// </summary>
            public const int ObjectDoesNotExist = 32277441;

            /// <summary>
            /// Error code indicating API request limit has been exceeded (throttling).
            /// </summary>
            public const int ApiLimitExceeded = -2147204784;
        }

        /// <summary>
        /// Language and localization constants.
        /// </summary>
        public static class Localization
        {
            /// <summary>
            /// Default language code for English (United States).
            /// </summary>
            public const int EnglishUS = 1033;

            /// <summary>
            /// Default language code used for all labels and descriptions.
            /// </summary>
            public const int DefaultLanguageCode = EnglishUS;
        }

        /// <summary>
        /// Default attribute metadata values.
        /// </summary>
        public static class AttributeDefaults
        {
            /// <summary>
            /// Default maximum length for single line text fields.
            /// </summary>
            public const int TextMaxLength = 100;

            /// <summary>
            /// Default maximum length for multi-line text (memo) fields.
            /// </summary>
            public const int MemoMaxLength = 2000;

            /// <summary>
            /// Default precision for decimal and floating point numbers.
            /// </summary>
            public const int DefaultPrecision = 2;

            /// <summary>
            /// Minimum value for money/currency fields.
            /// </summary>
            public const double MoneyMinValue = -922337203685477.0000;

            /// <summary>
            /// Maximum value for money/currency fields.
            /// </summary>
            public const double MoneyMaxValue = 922337203685477.0000;

            /// <summary>
            /// Minimum value for floating point number fields.
            /// </summary>
            public const double DoubleMinValue = -100000000000.0;

            /// <summary>
            /// Maximum value for floating point number fields.
            /// </summary>
            public const double DoubleMaxValue = 100000000000.0;

            /// <summary>
            /// Precision source for money fields (2 = use currency precision).
            /// </summary>
            public const int MoneyPrecisionSource = 2;
        }

        /// <summary>
        /// Excel-related constants.
        /// </summary>
        public static class Excel
        {
            /// <summary>
            /// Default header row index (1-based).
            /// </summary>
            public const int DefaultHeaderRow = 1;

            /// <summary>
            /// Default data start row index (1-based).
            /// </summary>
            public const int DefaultDataStartRow = 2;

            /// <summary>
            /// Maximum retry attempts when Excel file is locked.
            /// </summary>
            public const int MaxFileAccessRetries = 5;

            /// <summary>
            /// Initial retry delay in milliseconds when Excel file is locked.
            /// </summary>
            public const int InitialRetryDelayMs = 500;
        }

        /// <summary>
        /// API and performance-related constants.
        /// </summary>
        public static class Api
        {
            /// <summary>
            /// Maximum number of requests to batch together.
            /// </summary>
            public const int MaxBatchSize = 50;

            /// <summary>
            /// Default timeout for Dataverse operations in seconds.
            /// </summary>
            public const int DefaultTimeoutSeconds = 120;

            /// <summary>
            /// Number of retry attempts for transient failures.
            /// </summary>
            public const int MaxRetryAttempts = 3;

            /// <summary>
            /// Base delay for exponential backoff in milliseconds.
            /// </summary>
            public const int RetryBaseDelayMs = 1000;
        }

        /// <summary>
        /// Application messages and templates.
        /// </summary>
        public static class Messages
        {
            public const string AutoGeneratedColumnDescription = "Auto-generated column: {0}";
            public const string AutoGeneratedTableDescription = "Auto-generated table: {0}";
            public const string AutoGeneratedChoiceDescription = "Auto-generated choice column: {0}";
            public const string PrimaryNameFieldDescription = "Primary name field";
        }

        /// <summary>
        /// File and path constants.
        /// </summary>
        public static class Files
        {
            public const string DefaultConfigFile = "appsettings.json";
            public const string DefaultOutputCsvName = "new_schemas.csv";
            public const string SampleExcelFileName = "sample_schema.xlsx";
            public const string LogDirectory = "logs";
        }
    }
}
