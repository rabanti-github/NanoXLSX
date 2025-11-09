/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Registry
{
    /// <summary>
    /// Static class, holding UUIDs to be used for registering packages, containing plug-ins or identifiers for data entities
    /// </summary>
    /// \remark <remarks>The UUID strings shall never be altered. New UUIDs may be added. Obsolete may be completely removed.</remarks>
    public static class PlugInUUID
    {

        #region writerUUIDs
        /// <summary>
        /// UUID for the password writer, when a workbook is saved
        /// </summary>
        public const string PasswordWriter = "8106E566-60D6-45DB-BF87-33AB3882C019";
        /// <summary>
        /// UUID for the workbook writer, when a workbook is saved
        /// </summary>
        public const string WorkbookWriter = "D4272E3A-AC56-4524-9B9F-7B1448DF536B";
        /// <summary>
        /// UUID for the worksheet writer, when a workbook is saved
        /// </summary>
        public const string WorksheetWriter = "51F952E9-A914-4F12-B1CC-2F6C1F3637D7";
        /// <summary>
        /// UUID for the style writer, when a workbook is saved
        /// </summary>
        public const string StyleWriter = "009D7028-E8D9-4BB6-B5C7-F6D5EA2BA01F";
        /// <summary>
        /// UUID for the shared strings writer, when a workbook is saved
        /// </summary>
        public const string SharedStringsWriter = "731BF436-E28D-4136-BEF4-394D2CC65E01";
        /// <summary>
        /// UUID for the matadata writer (app data), when a workbook is saved
        /// </summary>
        public const string MetadataAppWriter = "49910428-CACB-475A-B39D-833D384DADE8";
        /// <summary>
        /// UUID for the metadata writer (core data), when a workbook is saved
        /// </summary>
        public const string MetadataCoreWriter = "19C28EEF-D80E-4A22-9B30-26376C7512FE";
        /// <summary>
        /// UUID for the theme writer, when a workbook is saved
        /// </summary>
        public const string ThemeWriter = "62E3A926-08F3-4343-ACCE-2A42096C3235";
        #endregion

        #region generalWriterQueueUUIDs
        /// <summary>
        /// UUID for the prepending queue. Plug-ins can register to this queue to be executed before the regular XLSX writers
        /// </summary>
        public const string WriterPrependingQueue = "772C4BF6-ED81-4127-80C7-C99D2B5C5EEC";

        /// <summary>
        /// UUID for the prepending queue that holds plug-ins for registering additional package parts for the XLSX building process
        /// </summary>
        public const string WriterPackageRegistryQueue = "C0CE40AC-14D5-4403-A5A3-018C6057A80E";

        /// <summary>
        /// UUID for the appending queue. Plug-ins can register to this queue to be executed after the regular XLSX writers
        /// </summary>
        public const string WriterAppendingQueue = "04F73656-C355-40A9-9E68-CB21329F3E53";
        #endregion


        #region inlineQueueWriterUUIDs
        /// <summary>
        /// UUID for in-line queued writers, appended right after the execution of the workbook writer, when a workbook is saved
        /// </summary>
        public const string WorkbookInlineWriter = "E69CEC04-A5CD-4DC2-9517-88F895C5CB1E";
        /// <summary>
        /// UUID for in-line queued writers, appended right after the execution of the worksheet writer, when a workbook is saved
        /// </summary>
        public const string WorksheetInlineWriter = "E0F6C065-00F8-4A67-AFAF-F358342845BC";
        /// <summary>
        /// UUID for in-line queued writers, appended right after the execution of the style writer, when a workbook is saved
        /// </summary>
        public const string StyleInlineWriter = "E9358F10-DD9B-4C5B-9BBB-DC32D5EB0DBB";
        /// <summary>
        /// UUID for in-line queued writers, appended right after the execution of the shared strings writer, when a workbook is saved
        /// </summary>
        public const string SharedStringsInlineWriter = "1E87131E-E6BA-4292-B4E5-55B73233D3F5";
        /// <summary>
        /// UUID for in-line queued writers, appended right after the execution of the matadata writer (app data), when a workbook is saved
        /// </summary>
        public const string MetadataAppInlineWriter = "AB45D7E1-7FF9-43D9-B482-91D677A7D614";
        /// <summary>
        /// UUID for in-line queued writers, appended right after the execution of the metadata writer (core data), when a workbook is saved
        /// </summary>
        public const string MetadataCoreInlineWriter = "85AC02E3-1F92-4921-BC69-39B3F328ABCD";
        /// <summary>
        /// UUID for in-line queued writers, appended right after the execution of the theme writer, when a workbook is saved
        /// </summary>
        public const string ThemeInlineWriter = "4CB6FD0E-AB69-40E9-B048-06B0E00C892D";
        #endregion

        #region readerUUIDs
        /// <summary>
        /// UUID for the password reader, when a workbook is loaded
        /// </summary>
        public const string PasswordReader = "1090EEDC-27AB-4A90-AAAB-E9B02C086082";
        /// <summary>
        /// UUID for the workbook reader, when a workbook is loaded
        /// </summary>
        public const string WorkbookReader = "B8C3405A-081C-453B-9C88-6A4BD7F5359B";
        /// <summary>
        /// UUID for the worksheet reader, when a workbook is loaded
        /// </summary>
        public const string WorksheetReader = "1DE75D75-5BF9-48EA-9387-DCF5459EC401";
        /// <summary>
        /// UUID for the style reader, when a workbook is loaded
        /// </summary>
        public const string StyleReader = "67AAB19A-4BF1-41B4-BC86-8C5BB5BB91F6";
        /// <summary>
        /// UUID for the shared strings reader, when a workbook is loaded
        /// </summary>
        public const string SharedStringsReader = "FF9BC0E6-59BF-4A16-B289-3F2AFD568438";
        /// <summary>
        /// UUID for the matadata reader (app data), when a workbook is loaded
        /// </summary>
        public const string MetadataAppReader = "28C59145-7BB8-416F-BAC9-0130DD8557F9";
        /// <summary>
        /// UUID for the metadata reader (core data), when a workbook is loaded
        /// </summary>
        public const string MetadataCoreReader = "B53F0F3E-71FF-43F0-B60C-C3478DE65788";
        /// <summary>
        /// UUID for the theme reader, when a workbook is loaded
        /// </summary>
        public const string ThemeReader = "B4733D00-B596-4440-8E33-A803289848BC";
        /// <summary>
        /// UUID for the relationship reader, when a workbook is loaded
        /// </summary>
        public const string RelationshipReader = "DB9AF89B-6181-4F94-A666-5AB70840EDDF";
        #endregion

        #region generalReaderQueueUUIDs
        /// <summary>
        /// UUID for the prepending queue. Plug-ins can register to this queue to be executed before the regular XLSX readers
        /// </summary>
        public const string ReaderPrependingQueue = "658A903B-512D-490C-A99B-40C0B0947CBF";

        /// <summary>
        /// UUID for the prepending queue that holds plug-ins for registering additional package parts for the XLSX reading process (e.g. additional XML files to be parsed)
        /// </summary>
        public const string ReaderPackageRegistryQueue = "1DD50B15-6EB8-451B-A6A8-C9265A8EF55C";

        /// <summary>
        /// UUID for the appending queue. Plug-ins can register to this queue to be executed after the regular XLSX readers
        /// </summary>
        public const string ReaderAppendingQueue = "69EE822E-910E-4E6B-BC5B-8F27629933AF";
        #endregion

        #region inlineQueueReaderUUIDs
        /// <summary>
        /// UUID for in-line queued readers, appended right after the execution of the workbook reader, when a workbook is loaded
        /// </summary>
        public const string WorkbookInlineReader = "33782BED-FCBA-4BE1-911A-5327C64B9580";
        /// <summary>
        /// UUID for in-line queued reader, appended right after the execution of the worksheet reader, when a workbook is loaded
        /// </summary>
        public const string WorksheetInlineReader = "20BE8320-9B90-41D2-8580-E1FE05DDC881";
        /// <summary>
        /// UUID for in-line queued reader, appended right after the execution of the style reader, when a workbook is loaded
        /// </summary>
        public const string StyleInlineReader = "9AC00387-E677-4F1C-88D6-558DAE6FF764";
        /// <summary>
        /// UUID for in-line queued reader, appended right after the execution of the shared strings reader, when a workbook is loaded
        /// </summary>
        public const string SharedStringsInlineReader = "3730F89E-CD7C-4BD8-B6AC-A18D803ADB2B";
        /// <summary>
        /// UUID for in-line queued reader, appended right after the execution of the matadata reader (app data), when a workbook is loaded
        /// </summary>
        public const string MetadataAppInlineReader = "789AFD19-31C5-409A-86C6-7CF5CC49B9C1";
        /// <summary>
        /// UUID for in-line queued reader, appended right after the execution of the metadata reader (core data), when a workbook is loaded
        /// </summary>
        public const string MetadataCoreInlineReader = "64A26388-EAD1-4435-AC07-A7FF18DCEEB7";
        /// <summary>
        /// UUID for in-line queued reader, appended right after the execution of the theme reader, when a workbook is loaded
        /// </summary>
        public const string ThemeInlineReader = "4B44E8A8-4560-44EB-8B24-5E11FDC04971";
        /// <summary>
        /// UUID for in-line queued reader, appended right after the execution of the relationship reader, when a workbook is loaded
        /// </summary>
        public const string RelationshipInlineReader = "E474D078-FBBC-49BE-B0B8-6086C07023DA";
        #endregion

        #region entityUUIDs
        /// <summary>
        /// UUID for the worksheet definitions section, on reading a workbook
        /// </summary>
        public const string WorksheetDefinitionEntity = "40CF0799-E4E7-4EA7-925F-BB6C9E8F588A";
        /// <summary>
        /// UUID for the selected worksheet, on reading a workbook
        /// </summary>
        public const string SelectedWorksheetEntity = "DD9B5E9B-2276-484D-B36A-B1F5EB6EE08A";
        /// <summary>
        /// UUID for the worksheet relationship section, on reading a workbook
        /// </summary>
        public const string RelationshipEntity = "F2DECC2C-544A-4B22-8C6E-386464586E60";
        /// <summary>
        /// UUID for the styles, on reading a workbook
        /// </summary>
        public const string StyleEntity = "638F9F5A-334A-49A1-BE07-1F2F3BFB70C4";
        #endregion

    }
}
