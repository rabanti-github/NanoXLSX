/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Linq;
using NanoXLSX.Interfaces;

namespace NanoXLSX.Registry
{
    /// <summary>
    /// Class to register plug-in classes that extends the functionality of NanoXLSX (Core or any other package) without being defined as fixed dependencies
    /// </summary>
    public static class PluginLoader
    {
        private static bool initialized = false;

        private static Dictionary<string, PluginInstance> pluginClasses = new Dictionary<string, PluginInstance>();
        private static Dictionary<string, List<PluginInstance>> queuePluginClasses = new Dictionary<string, List<PluginInstance>>();

        /// <summary>
        /// Initializes the plug-in loader process. If already initialized, the method returns without action
        /// </summary>
        public static void Initialize()
        {

            if (initialized)
            {
                return;
            }
            LoadReferencedAssemblies();
            initialized = true;

        }

        /// <summary>
        /// Method to load all currently referenced assemblies and its classes
        /// </summary>
        private static void LoadReferencedAssemblies()
        {
            var loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();
            var loadedPaths = loadedAssemblies
                .Where(a => !a.IsDynamic && !string.IsNullOrEmpty(a.Location)) // Exclude dynamic assemblies
                .Select(a => a.Location)
                .ToArray();
            var referencedPaths = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.dll")
                .Where(path => !loadedPaths.Contains(path, StringComparer.InvariantCultureIgnoreCase));

            foreach (var path in referencedPaths)
            {
                try
                {
                    Assembly assembly = Assembly.LoadFrom(path);
                    RegisterPlugins(assembly);
                }
                catch
                {
                    // Log or handle exceptions if necessary
                }
            }
        }

        /// <summary>
        /// Method to analyze and register plug-ins from a referenced assembly
        /// </summary>
        /// <param name="assembly">Assembly to analyze</param>
        private static void RegisterPlugins(Assembly assembly)
        {
            IEnumerable<Type> replacingPluginTypes = GetAssemblyPluginsByType(assembly, typeof(NanoXlsxPluginAttribute));
            IEnumerable<Type> queuePluginTypes = GetAssemblyPluginsByType(assembly, typeof(NanoXlsxQueuePluginAttribute));
            HandleReplacingPlugins(replacingPluginTypes);
            HandleQueuePlugins(queuePluginTypes);
        }

        /// <summary>
        /// Method to get an enumeration of all class types from an assembly, matching the specified plug-in attribute type
        /// </summary>
        /// <param name="assembly">Assembly to analyze</param>
        /// <param name="attributeType">Plug-in attribute type, declared on classes of the assembly</param>
        /// <returns>IEnumerable of class types, matching the criteria</returns>
        private static IEnumerable<Type> GetAssemblyPluginsByType(Assembly assembly, Type attributeType)
        {
            IEnumerable<Type> pluginTypes = assembly.GetTypes()
                .Where(type => type.IsClass &&
                               !type.IsAbstract &&
                               typeof(IPlugin).IsAssignableFrom(type) && // Ensure the type implements IPlugin
                               type.GetCustomAttribute(attributeType) != null);
            return pluginTypes;
        }

        /// <summary>
        /// Method to handle (register) an enumeration of plug-in class types as replacing plug-ins
        /// </summary>
        /// <param name="pluginTypes">IEnumerable of plug-in class types to handle</param>
        private static void HandleReplacingPlugins(IEnumerable<Type> pluginTypes)
        {
            foreach (Type pluginType in pluginTypes)
            {
                NanoXlsxPluginAttribute attribute = pluginType.GetCustomAttribute<NanoXlsxPluginAttribute>();
                if (attribute != null)
                {
                    if (pluginClasses.ContainsKey(attribute.PluginUUID))
                    {
                        if (attribute.PluginOrder <= pluginClasses[attribute.PluginUUID].Order)
                        {
                            // Skip duplicates with lower order numbers
                            pluginClasses[attribute.PluginUUID] = new PluginInstance(attribute.PluginUUID, attribute.PluginOrder, pluginType);
                        }
                    }
                    else if (!pluginClasses.ContainsKey(attribute.PluginUUID))
                    {
                        pluginClasses.Add(attribute.PluginUUID, new PluginInstance(attribute.PluginUUID, attribute.PluginOrder, pluginType));
                    }
                }
            }
        }

        /// <summary>
        /// Method to handle (register) an enumeration of plug-in class types as queuing plug-ins
        /// </summary>
        /// <param name="queuePluginTypes">IEnumerable of plug-in class types to handle</param>
        private static void HandleQueuePlugins(IEnumerable<Type> queuePluginTypes)
        {
            foreach (Type pluginType in queuePluginTypes)
            {
                NanoXlsxQueuePluginAttribute attribute = pluginType.GetCustomAttribute<NanoXlsxQueuePluginAttribute>();
                if (attribute != null)
                {
                    if (!queuePluginClasses.ContainsKey(attribute.QueueUUID))
                    {
                        queuePluginClasses.Add(attribute.QueueUUID, new List<PluginInstance>());
                    }
                    queuePluginClasses[attribute.QueueUUID].Add(new PluginInstance(attribute.PluginUUID, attribute.PluginOrder, pluginType));
                }
            }
            // Sort each list based on PluginOrder (ascending)
            foreach (KeyValuePair<string, List<PluginInstance>> entry in queuePluginClasses)
            {
                entry.Value.Sort((a, b) => a.Order.CompareTo(b.Order));
            }
        }

        /// <summary>
        /// Method to get a replacing plug-in class instance. If not found, the fall back instance will be returned.
        /// If the fall back is returned, this means normally that no plug-in was loaded to replace the requested instance.
        /// </summary>
        /// <typeparam name="T">Plug-in type</typeparam>
        /// <param name="pluginUUID">Plug-in type</param>
        /// <param name="fallBackInstance">Fall back instance if no plug-in with the defined UUID was registered</param>
        /// <returns>Found instance or fall back</returns>
        internal static T GetPlugin<T>(string pluginUUID, T fallBackInstance)
        {
            if (pluginClasses.ContainsKey(pluginUUID))
            {
                PluginInstance plugin = pluginClasses[pluginUUID];
                return (T)Activator.CreateInstance(plugin.Type);
            }
            else
            {
                return fallBackInstance;
            }
        }

        /// <summary>
        /// Method to get the next instance of a queue plug-in. To get the first one of a queue, the parameter 'lastPluginUUID' is set as null.
        /// If no further plug-in can be found in the queue, null will be returned as instance
        /// </summary>
        /// <typeparam name="T">Plug-in type</typeparam>
        /// <param name="queueUUID">UUID of the queue</param>
        /// <param name="lastPluginUUID">UUID of the last plug-in instance, to determine the next one</param>
        /// <param name="currentPluginUUID">Out parameter that return the UUID of the determined, next plug-in instance</param>
        /// <returns>Plug-in instance or null, if the end of the queue was reached</returns>
        internal static T GetNextQueuePlugin<T>(string queueUUID, string lastPluginUUID, out string currentPluginUUID)
        {
            if (queuePluginClasses.ContainsKey(queueUUID) && queuePluginClasses[queueUUID].Count > 0)
            {
                PluginInstance plugin = null;
                List<PluginInstance> pluginList = queuePluginClasses[queueUUID];
                if (lastPluginUUID == null)
                {
                    plugin = pluginList[0];
                }
                else
                {
                    // Find the next plug-in after the currentPluginUUID
                    int index = pluginList.FindIndex(p => p.UUID == lastPluginUUID);
                    if (index >= 0 && index + 1 < pluginList.Count)
                    {
                        plugin = pluginList[index + 1]; // Get the next plug-in in the list
                    }
                }
                if (plugin == null)
                {
                    currentPluginUUID = null;
                    return default(T);
                }
                currentPluginUUID = plugin.UUID;
                return (T)Activator.CreateInstance(plugin.Type);
            }
            else
            {
                currentPluginUUID = null;
                return default(T);
            }
        }

        /// <summary>
        /// Helper class to hold plug-in information
        /// </summary>
        private sealed class PluginInstance
        {
            /// <summary>
            /// UUID of the plug-in
            /// </summary>
            public string UUID { get; private set; }
            /// <summary>
            /// Order number of the plug-in, used to select the valid one (if replacing) or the queue order
            /// </summary>
            public int Order { get; private set; }
            /// <summary>
            /// Class type of the plug-in
            /// </summary>
            /// \remark <remarks>All types, representing a NanoXLSX plug-in must have an empty constructor that can be uses for initialization</remarks>
            public Type Type { get; private set; }

            /// <summary>
            /// Constructor with parameters
            /// </summary>
            /// <param name="uuid">UUID of the plug-in</param>
            /// <param name="order">Order number</param>
            /// <param name="type">Class type</param>
            internal PluginInstance(string uuid, int order, Type type)
            {
                this.UUID = uuid;
                this.Order = order;
                this.Type = type;
            }
        }
    }
}
