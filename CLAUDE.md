# Project Summary

## Objective
Create a console app for the detection, documentation, and detection of new Dataverse schema.

## Requirements

1. Console app can be configured to consume a given Excel file. If none configured, ask user.
2. Console app should be configured to look for three specific columns by name which will contain a list of schema to be checked for. First column will contain table name, second column column name, and third column column type.
3. App will then connect to Dataverse environment and scan for schema to see if it is created.
4. User should then have the option to output a list of new schema to a csv file, create the new schema, or both.
5. If create schema selected, user should be asked which solution should contain the schema and schema should be created there. If publisher prefix does not match, flag with error.