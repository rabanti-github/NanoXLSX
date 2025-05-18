/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Linq;
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
        private readonly Dictionary<string, Dictionary<string, Dictionary<string, DataEntry>>> data
            = new Dictionary<string, Dictionary<string, Dictionary<string, DataEntry>>>();

        /// <summary>
        /// Registers or updates a value for a given plug-in, object.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="valueId">ID of the value, represented as number (e.g. index)</param>
        /// <param name="value">Generic value</param>
        /// <param name="persistent">Optional parameter that indicates whether the set data is persistent (true) or temporary (false = default)</param>
        /// \remark <remarks>The entity ID - which is neglected in this method - is automatically set to the value <see cref="DEFAULT_ENTITY_ID"/> (empty)</remarks>
        public void SetData(string plugInId, int valueId, object value, bool persistent = false)
        {
            string id = ParserUtils.ToString(valueId);
            SetData(plugInId, id, value, persistent);
        }

        /// <summary>
        /// Registers or updates a value for a given plug-in, object.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="valueId">ID of the value (e.g. cell address in a worksheet)</param>
        /// <param name="value">Generic value</param>
        /// <param name="persistent">Optional parameter that indicates whether the set data is persistent (true) or temporary (false = default)</param>
        /// \remark <remarks>The entity ID - which is neglected in this method - is automatically set to the value <see cref="DEFAULT_ENTITY_ID"/> (empty)</remarks>
        public void SetData(string plugInId, string valueId, object value, bool persistent = false)
        {
            SetData(plugInId, DEFAULT_ENTITY_ID, valueId, value, persistent);
        }

        /// <summary>
        /// Registers or updates a value for a given plug-in, entity, and object.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="entityId">ID of the entity (e.g. a worksheet ID)</param>
        /// <param name="valueId">ID of the value, represented as number (e.g. index)</param>
        /// <param name="value">Generic value</param>
        /// <param name="persistent">Optional parameter that indicates whether the set data is persistent (true) or temporary (false = default)</param>
        public void SetData(string plugInId, string entityId, int valueId, object value, bool persistent = false)
        {
            string id = ParserUtils.ToString(valueId);
            SetData(plugInId, entityId, id, value, persistent);
        }

        /// <summary>
        /// Registers or updates a value for a given plug-in, entity, and object.
        /// </summary>
        /// <param name="plugInId">Plug-in ID / UUID or any kind of general identification</param>
        /// <param name="entityId">ID of the entity (e.g. a worksheet ID)</param>
        /// <param name="valueId">ID of the value (e.g. cell address in a worksheet)</param>
        /// <param name="value">Generic value</param>
        /// <param name="persistent">Optional parameter that indicates whether the set data is persistent (true) or temporary (false = default)</param>
        public void SetData(string plugInId, string entityId, string valueId, object value, bool persistent = false)
        {
            if (!data.TryGetValue(plugInId, out Dictionary<string, Dictionary<string, DataEntry>> pluginData))
            {
                pluginData = new Dictionary<string, Dictionary<string, DataEntry>>();
                data[plugInId] = pluginData;
            }

            if (!pluginData.TryGetValue(entityId, out Dictionary<string, DataEntry> entityData))
            {
                entityData = new Dictionary<string, DataEntry>();
                pluginData[entityId] = entityData;
            }

            entityData[valueId] = new DataEntry(value, persistent);
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
            if (data.TryGetValue(plugInId, out Dictionary<string, Dictionary<string, DataEntry>> pluginData) &&
                pluginData.TryGetValue(entityId, out Dictionary<string, DataEntry> entityData) &&
                entityData.TryGetValue(valueId, out DataEntry entry))
            {
                return entry.Value;
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
            if (data.TryGetValue(plugInId, out Dictionary<string, Dictionary<string, DataEntry>> pluginData) &&
                pluginData.TryGetValue(entityId, out Dictionary<string, DataEntry> entityData))
            {
                foreach (KeyValuePair<string, DataEntry> entry in entityData)
                {
                    if (entry.Value.Value is T value)
                    {
                        result.Add(value);
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Clears only temporary (non-persistent) entries.
        /// </summary>
        public void ClearTemporaryData()
        {
            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, DataEntry>>> pluginPair in data.ToList())
            {
                Dictionary<string, Dictionary<string, DataEntry>> pluginData = pluginPair.Value;
                foreach (KeyValuePair<string, Dictionary<string, DataEntry>> entityPair in pluginData.ToList())
                {
                    Dictionary<string, DataEntry> entityData = entityPair.Value;
                    List<string> keysToRemove = new List<string>();
                    foreach (KeyValuePair<string, DataEntry> kvp in entityData)
                    {
                        if (!kvp.Value.Persistent)
                        {
                            keysToRemove.Add(kvp.Key);
                        }
                    }
                    foreach (string key in keysToRemove)
                    {
                        entityData.Remove(key);
                    }
                    if (entityData.Count == 0)
                    {
                        pluginData.Remove(entityPair.Key);
                    }
                }
                if (pluginData.Count == 0)
                {
                    data.Remove(pluginPair.Key);
                }
            }
        }

        /// <summary>
        /// Clears all elements from the data collection, removing any stored items.
        /// </summary>
        public void ClearData()
        {
            data.Clear();
        }

        /// <summary>
        /// Sub class representing a data entry
        /// </summary>
        private class DataEntry
        {
            /// <summary>
            /// Value object
            /// </summary>
            public object Value { get; }
            /// <summary>
            /// If true, the data is persistent (e.g. from an extended reader plug-ins), otherwise it is temporary
            /// </summary>
            public bool Persistent { get; }

            /// <summary>
            /// Default constructor
            /// </summary>
            /// <param name="value">Value object</param>
            /// <param name="persistent">Persistent parameter</param>
            public DataEntry(object value, bool persistent)
            {
                Value = value;
                Persistent = persistent;
            }
        }

    }
}
