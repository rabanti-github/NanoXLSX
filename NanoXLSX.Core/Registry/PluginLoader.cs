/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;
using NanoXLSX.Interfaces;
using NanoXLSX.Registry.Attributes;

namespace NanoXLSX.Registry
{
    /// <summary>
    /// Class to register plug-in classes that extends the functionality of NanoXLSX (Core or any other package) without being defined as fixed dependencies
    /// </summary>
    public static class PlugInLoader
    {
        private static bool initialized;
        private static readonly object _lock = new object();

        private static readonly Dictionary<string, PlugInInstance> plugInClasses = new Dictionary<string, PlugInInstance>();
        private static readonly Dictionary<string, List<PlugInInstance>> queuePlugInClasses = new Dictionary<string, List<PlugInInstance>>();

        /// <summary>
        /// Initializes the plug-in loader process. If already initialized, the method returns without action
        /// </summary>
        /// <returns>True if initialized in this call, otherwise false (if already initialized)</returns>
        public static bool Initialize()
        {

            if (initialized)
            {
                return false;
            }
            lock (_lock)
            {
                LoadReferencedAssemblies();
                initialized = true;
                return initialized;
            }
        }

        /// <summary>
        /// Method to load all currently referenced assemblies and its classes
        /// </summary>
        [ExcludeFromCodeCoverage] // Indirectly tested by InjectPlugins
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
                    Assembly assembly = Assembly.Load(path);
                    RegisterPlugIns(assembly); // Here, several classes are determined that has Attributes, describing them as plug-ins
                }
                catch
                {
                    // Log or handle exceptions if necessary
                }
            }
        }

        /// <summary>
        /// Method to manually inject plug-in classes.
        /// This method is mainly used by unit test. However, it could also be used by Plug-ins to inject additional "virtual" plug-ins
        /// </summary>
        /// <param name="pluginTypes">List of classes that are annotated with the attributes <see cref="NanoXlsxPlugInAttribute"/> and <see cref="NanoXlsxQueuePlugInAttribute"/></param>
        internal static void InjectPlugins(List<Type> pluginTypes)
        {
            lock (_lock)
            {
                // Collect all types decorated with the NanoXlsxPlugInAttribute.
                IEnumerable<Type> replacingPluginTypes = pluginTypes
                .Where(t => t.GetCustomAttribute<NanoXlsxPlugInAttribute>() != null);

                // Collect all types decorated with the NanoXlsxQueuePlugInAttribute.
                IEnumerable<Type> queuePluginTypes = pluginTypes
                    .Where(t => t.GetCustomAttributes<NanoXlsxQueuePlugInAttribute>().Any());

                // Pass the collected types to the appropriate handlers.
                HandleReplacingPlugIns(replacingPluginTypes);
                HandleQueuePlugIns(queuePluginTypes);
                initialized = true;
            }
        }

        /// <summary>
        /// Method to dispose all loaded plugins.
        /// This method is mainly used by unit test. However, it could also be used by Plug-ins to unload "virtual" plug-ins
        /// </summary>
        internal static void DisposePlugins()
        {
            lock (_lock)
            {
                plugInClasses.Clear();
                queuePlugInClasses.Clear();
                initialized = false;
            }
        }

        /// <summary>
        /// Method to analyze and register plug-ins from a referenced assembly
        /// </summary>
        /// <param name="assembly">Assembly to analyze</param>
        [ExcludeFromCodeCoverage] // Indirectly tested by InjectPlugins
        private static void RegisterPlugIns(Assembly assembly)
        {
            IEnumerable<Type> replacingPlugInTypes = GetAssemblyPlugInsByType(assembly, typeof(NanoXlsxPlugInAttribute));
            IEnumerable<Type> queuePlugInTypes = GetAssemblyPlugInsByType(assembly, typeof(NanoXlsxQueuePlugInAttribute));
            HandleReplacingPlugIns(replacingPlugInTypes);
            HandleQueuePlugIns(queuePlugInTypes);
        }


        /// <summary>
        /// Method to get an enumeration of all class types from an assembly, matching the specified plug-in attribute type
        /// </summary>
        /// <param name="assembly">Assembly to analyze</param>
        /// <param name="attributeType">Plug-in attribute type, declared on classes of the assembly</param>
        /// <returns>IEnumerable of class types, matching the criteria</returns>
        [ExcludeFromCodeCoverage] // Indirectly tested by InjectPlugins
        private static IEnumerable<Type> GetAssemblyPlugInsByType(Assembly assembly, Type attributeType)
        {
            List<Type> plugInTypes = new List<Type>();
            Type plugInInterface = typeof(IPlugIn);
            Type[] allTypes = assembly.GetTypes();

            for (int i = 0; i < allTypes.Length; i++)
            {
                Type type = allTypes[i];
                if (type.IsClass && !type.IsAbstract &&
                    plugInInterface.IsAssignableFrom(type) &&
                    type.GetCustomAttribute(attributeType) != null)
                {
                    plugInTypes.Add(type);
                }
            }
            return plugInTypes;
        }

        /// <summary>
        /// Method to handle (register) an enumeration of plug-in class types as replacing plug-ins
        /// </summary>
        /// <param name="plugInTypes">IEnumerable of plug-in class types to handle</param>
        [ExcludeFromCodeCoverage] // Indirectly tested by InjectPlugins
        private static void HandleReplacingPlugIns(IEnumerable<Type> plugInTypes)
        {
            foreach (Type plugInType in plugInTypes)
            {
                IEnumerable<NanoXlsxPlugInAttribute> attributes = plugInType.GetCustomAttributes<NanoXlsxPlugInAttribute>();
                foreach (NanoXlsxPlugInAttribute attribute in attributes)
                {
                    if (plugInClasses.ContainsKey(attribute.PlugInUUID))
                    {
                        if (attribute.PlugInOrder <= plugInClasses[attribute.PlugInUUID].Order)
                        {
                            // Skip duplicates with lower order numbers
                            plugInClasses[attribute.PlugInUUID] = new PlugInInstance(attribute.PlugInUUID, attribute.PlugInOrder, plugInType);
                        }
                    }
                    else if (!plugInClasses.ContainsKey(attribute.PlugInUUID))
                    {
                        plugInClasses.Add(attribute.PlugInUUID, new PlugInInstance(attribute.PlugInUUID, attribute.PlugInOrder, plugInType));
                    }
                }
            }
        }

        /// <summary>
        /// Method to handle (register) an enumeration of plug-in class types as queuing plug-ins
        /// </summary>
        /// <param name="queuePlugInTypes">IEnumerable of plug-in class types to handle</param>
        private static void HandleQueuePlugIns(IEnumerable<Type> queuePlugInTypes)
        {
            foreach (Type plugInType in queuePlugInTypes)
            {
                IEnumerable<NanoXlsxQueuePlugInAttribute> attributes = plugInType.GetCustomAttributes<NanoXlsxQueuePlugInAttribute>();
                foreach (var attribute in attributes)
                {
                    if (!queuePlugInClasses.TryGetValue(attribute.QueueUUID, out var value))
                    {
                        value = new List<PlugInInstance>();
                        queuePlugInClasses.Add(attribute.QueueUUID, value);
                    }

                    value.Add(new PlugInInstance(attribute.PlugInUUID, attribute.PlugInOrder, plugInType));
                }
            }
            // Sort each list based on PlugInOrder (ascending)
            foreach (KeyValuePair<string, List<PlugInInstance>> entry in queuePlugInClasses)
            {
                entry.Value.Sort((a, b) => a.Order.CompareTo(b.Order));
            }
        }

        /// <summary>
        /// Method to get a replacing plug-in class instance. If not found, the fall back instance will be returned.
        /// If the fall back is returned, this means normally that no plug-in was loaded to replace the requested instance.
        /// </summary>
        /// <typeparam name="T">Plug-in type</typeparam>
        /// <param name="plugInUUID">Plug-in type</param>
        /// <param name="fallBackInstance">Fall back instance if no plug-in with the defined UUID was registered</param>
        /// <returns>Found instance or fall back</returns>
        /// \remark <remarks>This method is not intended to manage preserved plugin instances. When called, always a new instance of the plugin will be created</remarks>
        internal static T GetPlugIn<T>(string plugInUUID, T fallBackInstance)
        {
            if (plugInClasses.TryGetValue(plugInUUID, out var plugIn))
            {
                return (T)Activator.CreateInstance(plugIn.Type);
            }
            else
            {
                return fallBackInstance;
            }
        }

        /// <summary>
        /// Method to get the next instance of a queue plug-in. To get the first one of a queue, the parameter 'lastPlugInUUID' is set as null.
        /// If no further plug-in can be found in the queue, null will be returned as instance
        /// </summary>
        /// <typeparam name="T">Plug-in type</typeparam>
        /// <param name="queueUUID">UUID of the queue</param>
        /// <param name="lastPlugInUUID">UUID of the last plug-in instance, to determine the next one</param>
        /// <param name="currentPlugInUUID">Out parameter that return the UUID of the determined, next plug-in instance</param>
        /// <returns>Plug-in instance or null, if the end of the queue was reached</returns>
        /// /// \remark <remarks>This method is not intended to manage preserved plugin instances. When called, always a new instance of the plugin will be created</remarks>
        internal static T GetNextQueuePlugIn<T>(string queueUUID, string lastPlugInUUID, out string currentPlugInUUID)
        {
            if (queuePlugInClasses.TryGetValue(queueUUID, out var plugInList) && plugInList.Count > 0)
            {
                PlugInInstance plugIn = null;
                if (lastPlugInUUID == null)
                {
                    plugIn = plugInList[0];
                }
                else
                {
                    // Find the next plug-in after the currentPlugInUUID
                    int index = plugInList.FindIndex(p => p.UUID == lastPlugInUUID);
                    if (index >= 0 && index + 1 < plugInList.Count)
                    {
                        plugIn = plugInList[index + 1]; // Get the next plug-in in the list
                    }
                }
                if (plugIn == null)
                {
                    currentPlugInUUID = null;
                    return default;
                }
                currentPlugInUUID = plugIn.UUID;
                return (T)Activator.CreateInstance(plugIn.Type);
            }
            else
            {
                currentPlugInUUID = null;
                return default;
            }
        }

        /// <summary>
        /// Helper class to hold plug-in information
        /// </summary>
        private sealed class PlugInInstance
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
            internal PlugInInstance(string uuid, int order, Type type)
            {
                this.UUID = uuid;
                this.Order = order;
                this.Type = type;
            }
        }
    }
}
