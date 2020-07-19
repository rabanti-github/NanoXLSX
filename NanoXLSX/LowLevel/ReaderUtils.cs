using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace NanoXLSX.LowLevel
{
    public static class ReaderUtils
    {
        /// <summary>
        /// Gets the xml attribute of the passed XML node by its name
        /// </summary>
        /// <param name="targetName">Name of the target attribute</param>
        /// <param name="node">XML node that contains the attribute</param>
        /// <param name="fallbackValue">Optional fallback value if the attribute was not found. Default is null</param>
        /// <returns>Attribute value as string or default value if not found (can be null)</returns>
        public static string GetAttribute(string targetName, XmlNode node, string fallbackValue = null)
        {
            if (node.Attributes == null || node.Attributes.Count == 0)
            {
                return fallbackValue;
            }

            foreach (XmlAttribute attribute in node.Attributes)
            {
                if (attribute.Name == targetName)
                {
                    return attribute.Value;
                }
            }

            return fallbackValue;
        }
    }
}
