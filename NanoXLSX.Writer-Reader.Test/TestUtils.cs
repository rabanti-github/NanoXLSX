using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Xml;
using NanoXLSX.Extensions;
using NanoXLSX.Styles;
using Xunit;
using static NanoXLSX.Cell;

namespace NanoXLSX.Test.Writer_Reader.Utils
{
    [ExcludeFromCodeCoverage]
    public class TestUtils
    {
        private const string ASSEMBLY_RESOURCE_NAMESPACE = "NanoXLSX.Test.Writer_Reader"; // Change this on refactoring

        public static void AssertEquals<T>(T value1, T value2, T inequalValue, Address cellAddress)
        {
            Cell cell1 = new Cell(value1, CellType.DEFAULT, cellAddress);
            Cell cell2 = new Cell(value2, CellType.DEFAULT, cellAddress);
            Cell cell3 = new Cell(inequalValue, CellType.DEFAULT, cellAddress);
            Assert.True(cell1.Equals(cell2));
            Assert.False(cell1.Equals(cell3));
        }
        public static void AssertCellRange(string expectedAddresses, List<Address> addresses)
        {
            string[] addressStrings = SplitValues(expectedAddresses);
            List<Address> expected = new List<Address>();
            foreach (string address in addressStrings)
            {
                expected.Add(new Address(address));
            }
            Assert.Equal(expected.Count, addresses.Count);
            for (int i = 0; i < expected.Count; i++)
            {
                Assert.Equal(expected[i], addresses[i]);
            }
        }

        public static void AssertZipEntry(Stream zipStream, string pathInZip, string expectedContent)
        {
            using (var zip = new ZipArchive(zipStream, ZipArchiveMode.Read, leaveOpen: true))
            {
                var entry = zip.GetEntry(pathInZip);
                Assert.NotNull(entry);

                using (var reader = new StreamReader(entry.Open()))
                {
                    string content = reader.ReadToEnd();
                    Assert.Contains(expectedContent, content);
                }
            }
        }

        public static List<string> SplitValuesAsList(string valueString)
        {
            return new List<string>(SplitValues(valueString));
        }

        public static string[] SplitValues(string valueString)
        {
            if (valueString == null || valueString == "")
            {
                return new string[0];
            }
            return valueString.Split(new char[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
        }

        public static Stream GetResource(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return null;
            }
            Assembly assembly = Assembly.GetExecutingAssembly();
            string resourceName = $"{ASSEMBLY_RESOURCE_NAMESPACE}.Resources.{path}";
            try
            {
                return assembly.GetManifestResourceStream(resourceName);
            }
            catch
            {
                return null;
            }
        }

        public static object CreateInstance(Type sourceType, string sourceValue)
        {
            if (sourceType == typeof(decimal))
            {
                return decimal.Parse(sourceValue);
            }
            else if (sourceType == typeof(double))
            {
                return double.Parse(sourceValue);
            }
            else if (sourceType == typeof(int))
            {
                double d = double.Parse(sourceValue);
                return (int)d;
            }
            else if (sourceType == typeof(string))
            {
                return sourceValue.ToString(CultureInfo.InvariantCulture);
            }
            throw new ArgumentException("Not implemented source type: " + sourceType);
        }

        public static Cell SaveAndReadStyledCell(object value, Style style, string targetCellAddress)
        {
            return SaveAndReadStyledCell(value, value, style, targetCellAddress);
        }
        public static Cell SaveAndReadStyledCell(object givenValue, object expectedValue, Style style, string targetCellAddress)
        {
            Workbook workbook = new Workbook(false);
            workbook.AddWorksheet("sheet1");
            workbook.CurrentWorksheet.AddCell(givenValue, targetCellAddress, style);
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            Workbook givenWorkbook = WorkbookReader.Load(stream);
            Cell cell = givenWorkbook.CurrentWorksheet.Cells[targetCellAddress];
            Assert.Equal(expectedValue, cell.Value);
            return cell;
        }

        public static Workbook WriteAndReadWorkbook(Workbook workbook)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveAsStream(stream, true);
                stream.Position = 0;
                Workbook readWorkbook = WorkbookReader.Load(stream);
                return readWorkbook;
            }
        }

        public static void AssertExistingFile(string expectedPath, bool deleteAfterAssertion)
        {
            FileInfo fi = new FileInfo(expectedPath);
            Assert.True(fi.Exists);
            if (deleteAfterAssertion)
            {
                try
                {
                    fi.Delete();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Could not delete " + expectedPath + ": " + ex.Message);
                }
            }
        }

        public static string GetRandomName()
        {
            string path = Path.GetTempFileName();
            FileInfo fi = new FileInfo(path);
            if (fi.Exists)
            {
                fi.Delete();
            }
            return path.Replace(".tmp", ".xlsx");
        }

        /// <summary>
        /// Reads the first inner text value of the given node name from an XML stream.
        /// </summary>
        /// <param name="stream">The XML content as a readable stream.</param>
        /// <param name="nodeName">The name of the XML node to read.</param>
        /// <returns>The first inner text value of the node, or null if not found.</returns>
        public static string ReadFirstNodeValue(Stream stream, string nodeName)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }
            if (string.IsNullOrEmpty(nodeName))
            {
                throw new ArgumentException("Node name must be specified.", nameof(nodeName));
            }
            using (var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true,
                DtdProcessing = DtdProcessing.Ignore
            }))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == nodeName)
                    {
                        // ReadElementContentAsString moves past the end element automatically
                        return reader.ReadElementContentAsString();
                    }
                }
            }

            return null; // not found
        }

        /// <summary>
        /// Reads the defined attribute of the first occurring node with the given name from an XML stream.
        /// </summary>
        /// <param name="stream">The XML content as a readable stream.</param>
        /// <param name="nodeName">The name of the XML node to read.</param>
        /// <param name="attributeName">The name of the attribute to read.</param>"
        /// <returns>The value of the defined attribute from the first node occurrence, or null if not found.</returns>
        public static string ReadFirstAttributeValue(Stream stream, string nodeName, string attributeName)
        {
            List<string> values = ReadAllAttributeValues((MemoryStream)stream, nodeName, attributeName);
            if (values != null && values.Count > 0)
            {
                return values[0];
            }
            return null;
        }

        /// <summary>
        /// Reads the defined attributes of the any occurring node with the given name from an XML stream.
        /// </summary>
        /// <param name="stream">The XML content as a readable stream.</param>
        /// <param name="nodeName">The name of the XML node to read.</param>
        /// <param name="attributeName">The name of the attribute to read.</param>"
        /// <returns>A list of the values of the defined attribute from any node occurrence, or null if not found.</returns>
        public static List<string> ReadAllAttributeValues(MemoryStream stream, string nodeName, string attributeName)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }
            if (string.IsNullOrEmpty(nodeName) || string.IsNullOrEmpty(attributeName))
            {
                throw new ArgumentException("Node and attribute name must be specified.", nameof(nodeName));
            }
            List<string> values = new List<string>();
            stream.Position = 0;
            using (var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true,
                DtdProcessing = DtdProcessing.Ignore
            }))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == nodeName)
                    {
                        values.Add(reader.GetAttribute(attributeName));
                    }
                }
            }
            if (values.Count > 0)
            {
                return values;
            }
            else
            {
                return null; // not found
            }
        }
    }
}
