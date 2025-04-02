/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Registry
{
    /// <summary>
    /// Static class, holding UUIDs to be used for registering packages and containing plug-ins
    /// </summary>
    /// \remark <remarks>The UUID strings shall never be altered. New UUIDs may be added. Obsolete may be completely removed.</remarks>
    public static class PlugInUUID
    {
        
        #region writerUUIDs
        /// <summary>
        /// UUID for the password writer, when a workbook is saved
        /// </summary>
        public const string PASSWORD_WRITER = "8106E566-60D6-45DB-BF87-33AB3882C019";
        /// <summary>
        /// UUID for the workbook writer, when a workbook is saved
        /// </summary>
        public const string WORKBOOK_WRITER = "D4272E3A-AC56-4524-9B9F-7B1448DF536B";
        /// <summary>
        /// UUID for the worksheet writer, when a workbook is saved
        /// </summary>
        public const string WORKSHEET_WRITER = "51F952E9-A914-4F12-B1CC-2F6C1F3637D7";
        /// <summary>
        /// UUID for the style writer, when a workbook is saved
        /// </summary>
        public const string STYLE_WRITER = "009D7028-E8D9-4BB6-B5C7-F6D5EA2BA01F";
        /// <summary>
        /// UUID for the shared strings writer, when a workbook is saved
        /// </summary>
        public const string SHARED_STRING_WRITER = "731BF436-E28D-4136-BEF4-394D2CC65E01";
        /// <summary>
        /// UUID for the matadata writer (app data), when a workbook is saved
        /// </summary>
        public const string METADATA_APP_WRITER = "49910428-CACB-475A-B39D-833D384DADE8";
        /// <summary>
        /// UUID for the metadata writer (core data), when a workbook is saved
        /// </summary>
        public const string METADATA_CORE_WRITER = "19C28EEF-D80E-4A22-9B30-26376C7512FE";
        /// <summary>
        /// UUID for the theme writer, when a workbook is saved
        /// </summary>
        public const string THEME_WRITER = "62E3A926-08F3-4343-ACCE-2A42096C3235";
        #endregion

        #region generalWriterQueueUUIDs
        /// <summary>
        /// UUID for the prepending queue. Plug-ins can register to this queue to be executed before the regular XLSX writers
        /// </summary>
        public const string WRITER_PREPENDING_QUEUE = "772C4BF6-ED81-4127-80C7-C99D2B5C5EEC";

        /// <summary>
        /// UUID for the prepending queue that holds plug-ins for registering additional package parts for the XLSX building process
        /// </summary>
        public const string WRITER_PACKAGE_REGISTRY_QUEUE = "C0CE40AC-14D5-4403-A5A3-018C6057A80E";

        /// <summary>
        /// UUID for the appending queue. Plug-ins can register to this queue to be executed after the regular XLSX writers
        /// </summary>
        public const string WRITER_APPENDING_QUEUE = "04F73656-C355-40A9-9E68-CB21329F3E53";
        #endregion


         #region inlineQueueWriterUUIDs
        /// <summary>
        /// UUID for inline queued writers, appended right after the execution of the workbook writer, when a workbook is saved
        /// </summary>
        public const string WORKBOOK_INLINE_WRITER = "E69CEC04-A5CD-4DC2-9517-88F895C5CB1E";
        /// <summary>
        /// UUID for inline queued writers, appended right after the execution of the worksheet writer, when a workbook is saved
        /// </summary>
        public const string WORKSHEET_INLINE_WRITER = "E0F6C065-00F8-4A67-AFAF-F358342845BC";
        /// <summary>
        /// UUID for inline queued writers, appended right after the execution of the style writer, when a workbook is saved
        /// </summary>
        public const string STYLE_INLINE_WRITER = "E9358F10-DD9B-4C5B-9BBB-DC32D5EB0DBB";
        /// <summary>
        /// UUID for inline queued writers, appended right after the execution of the shared strings writer, when a workbook is saved
        /// </summary>
        public const string SHARED_QUEUE_STRING_WRITER = "1E87131E-E6BA-4292-B4E5-55B73233D3F5";
        /// <summary>
        /// UUID for inline queued writers, appended right after the execution of the matadata writer (app data), when a workbook is saved
        /// </summary>
        public const string METADATA_APP_INLINE_WRITER = "AB45D7E1-7FF9-43D9-B482-91D677A7D614";
        /// <summary>
        /// UUID for inline queued writers, appended right after the execution of the metadata writer (core data), when a workbook is saved
        /// </summary>
        public const string METADATA_CORE_INLINE_WRITER = "85AC02E3-1F92-4921-BC69-39B3F328ABCD";
        /// <summary>
        /// UUID for inline queued writers, appended right after the execution of the theme writer, when a workbook is saved
        /// </summary>
        public const string THEME_INLINE_WRITER = "4CB6FD0E-AB69-40E9-B048-06B0E00C892D";
        #endregion

    }
}
