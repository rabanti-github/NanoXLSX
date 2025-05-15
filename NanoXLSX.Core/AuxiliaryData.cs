/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using NanoXLSX.Utils;

namespace NanoXLSX
{
    /// <summary>
    /// Class holding non-core data, e.g. from a loader plug-in. The data may be processed by plug-ins
    /// </summary>
    internal class AuxiliaryData
    {
        /// <summary>
        /// Default entity ID if nothing is defined
        /// </summary>
        private const string DEFAULT_ENTITY_ID = "";

        // Structure: PlugInId => EntityId => SubEntityId => value
        private readonly Dictionary<string, Dictionary<string, Dictionary<string, object>>> data
            = new Dictionary<string, Dictionary<string, Dictionary<string, object>>>();

        /// <summary>
        /// Registers or updates a value for a given plug-in, object.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="valueId">ID of the value, represented as number (e.g. index)</param>
        /// <param name="value">Generic value</param>
        /// \remark <remarks>The entity ID - which is neglected in this method - is automatically set to the value <see cref="DEFAULT_ENTITY_ID"/> (empty)</remarks>
        public void SetData(string plugInId, int valueId, object value)
        {
            string id = ParserUtils.ToString(valueId);
            SetData(plugInId, id, value);
        }

        /// <summary>
        /// Registers or updates a value for a given plug-in, object.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="valueId">ID of the value (e.g. cell address in a worksheet)</param>
        /// <param name="value">Generic value</param>
        /// \remark <remarks>The entity ID - which is neglected in this method - is automatically set to the value <see cref="DEFAULT_ENTITY_ID"/> (empty)</remarks>
        public void SetData(string plugInId, string valueId, object value)
        {
            SetData(plugInId, DEFAULT_ENTITY_ID, valueId, value);
        }

        /// <summary>
        /// Registers or updates a value for a given plug-in, entity, and object.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="entityId">ID of the entity (e.g. a worksheet ID)</param>
        /// <param name="valueId">ID of the value, represented as number (e.g. index)</param>
        /// <param name="value">Generic value</param>
        public void SetData(string plugInId, string entityId, int valueId, object value)
        {
            string id = ParserUtils.ToString(valueId);
            SetData(plugInId, entityId, id, value);
        }

        /// <summary>
        /// Registers or updates a value for a given plug-in, entity, and object.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="entityId">ID of the entity (e.g. a worksheet ID)</param>
        /// <param name="valueId">ID of the value (e.g. cell address in a worksheet)</param>
        /// <param name="value">Generic value</param>
        public void SetData(string plugInId, string entityId, string valueId, object value)
        {
            if (!data.TryGetValue(plugInId, out var pluginData))
            {
                pluginData = new Dictionary<string, Dictionary<string, object>>();
                data[plugInId] = pluginData;
            }

            if (!pluginData.TryGetValue(entityId, out var entityData))
            {
                entityData = new Dictionary<string, object>();
                pluginData[entityId] = entityData;
            }

            entityData[valueId] = value;
        }

        /// <summary>
        /// Retrieves the value if it exists, or returns null.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="valueId">ID of the value, represented as number (e.g. index)</param>
        /// <returns>Generic value or null, if not found</returns>
        public object GetData(string plugInId, int valueId)
        {
            string id = ParserUtils.ToString(valueId);
            return GetData(plugInId, id);
        }

        /// <summary>
        /// Retrieves the value if it exists, or returns null.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="valueId">ID of the value (e.g. cell address in a worksheet)</param>
        /// <returns>Generic value or null, if not found</returns>
        public object GetData(string plugInId, string valueId)
        {
            return GetData(plugInId, DEFAULT_ENTITY_ID, valueId);
        }

        /// <summary>
        /// Retrieves the value if it exists, or returns null.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="entityId">ID of the entity (e.g. a worksheet ID)</param>
        /// <param name="valueId">ID of the value, represented as number (e.g. index)</param>
        /// <returns>Generic value or null, if not found</returns>
        public object GetData(string plugInId, int entityId, string valueId)
        {
            string id = ParserUtils.ToString(entityId);
            return GetData(plugInId, id, valueId);
        }

        /// <summary>
        /// Retrieves the value if it exists, or returns null.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="entityId">ID of the entity (e.g. a worksheet ID)</param>
        /// <param name="valueId">ID of the value (e.g. cell address in a worksheet)</param>
        /// <returns>Generic value or null, if not found</returns>
        public object GetData(string plugInId, string entityId, string valueId)
        {
            if (data.TryGetValue(plugInId, out var pluginData) &&
                pluginData.TryGetValue(entityId, out var entityData) &&
                entityData.TryGetValue(valueId, out var value))
            {
                return value;
            }
            return null;
        }

        /// <summary>
        /// Retrieves the typed value if it exists, or returns default (possibly null).
        /// </summary>
        /// <typeparam name="T">Target type of the value</typeparam>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="valueId">ID of the value, represented as number (e.g. index)</param>
        /// <returns>Typed value or default (possibly null), if not found</returns>
        public T GetData<T>(string plugInId, int valueId)
        {
            string id = ParserUtils.ToString(valueId);
            return GetData<T>(plugInId, id);
        }

        /// <summary>
        /// Retrieves the typed value if it exists, or returns default (possibly null).
        /// </summary>
        /// <typeparam name="T">Target type of the value</typeparam>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="valueId">ID of the value (e.g. cell address in a worksheet)</param>
        /// <returns>Typed value or default (possibly null), if not found</returns>
        public T GetData<T>(string plugInId, string valueId)
        {
            return GetData<T>(plugInId, DEFAULT_ENTITY_ID, valueId);
        }

        /// <summary>
        /// Retrieves the typed value if it exists, or returns default (possibly null).
        /// </summary>
        /// <typeparam name="T">Target type of the value</typeparam>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="entityId">ID of the entity (e.g. a worksheet ID)</param>
        /// <param name="valueId">ID of the value, represented as number (e.g. index)</param>
        /// <returns>Typed value or default (possibly null), if not found</returns>
        public T GetData<T>(string plugInId, string entityId, int valueId)
        {
            string id = ParserUtils.ToString(valueId);
            return GetData<T>(plugInId, entityId, id);
        }

        /// <summary>
        /// Retrieves the typed value if it exists, or returns default (possibly null).
        /// </summary>
        /// <typeparam name="T">Target type of the value</typeparam>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="entityId">ID of the entity (e.g. a worksheet ID)</param>
        /// <param name="valueId">ID of the value (e.g. cell address in a worksheet)</param>
        /// <returns>Typed value or default (possibly null), if not found</returns>
        public T GetData<T>(string plugInId, string entityId, string valueId)
        {
            object value = GetData(plugInId, entityId, valueId);
            return value is T ? (T)value : default;
        }

        public List<T> GetDataList<T>(string plugInId)
        {
            return GetDataList<T>(plugInId, DEFAULT_ENTITY_ID);
        }

        public List<T> GetDataList<T>(string plugInId, string entityId)
        {
            List<T> result = new List<T>();
            if (data.TryGetValue(plugInId, out Dictionary<string, Dictionary<string, object>> pluginData) &&
                pluginData.TryGetValue(entityId, out Dictionary<string, object> entityData))
            {
                foreach (KeyValuePair<string, object> kvp in entityData)
                {
                    if (kvp.Value is T value)
                    {
                        result.Add(value);
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Clears all elements from the data collection, removing any stored items.
        /// </summary>
        public void Clear()
        {
            data.Clear();
        }
    }
}
