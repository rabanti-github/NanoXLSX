/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Interfaces.Writer
{
    internal interface IWriteContext
    {
        /// <summary>
        /// Gets the workbook handled by the current writing operation.
        /// </summary>
        Workbook Workbook { get; }

        /// <summary>
        /// Gets the processing data that contains data outside of the workbook (used for preparation etc.)
        /// </summary>
        IWriterProcessingData WriterProcessingData { get; }

        /// <summary>
        /// Marks a writing feature as successfully prepared for the current
        /// writing operation.
        /// </summary>
        /// <param name="featureUuid">
        /// UUID of the prepared writing feature.
        /// </param>
        void MarkFeatureAsPrepared(string featureUuid);

        /// <summary>
        /// Determines whether a writing feature was successfully prepared
        /// for the current writing operation.
        /// </summary>
        /// <param name="featureUuid">
        /// UUID of the writing feature.
        /// </param>
        /// <returns>
        /// True if the feature was prepared; otherwise false.
        /// </returns>
        bool IsFeaturePrepared(string featureUuid);
    }
}
