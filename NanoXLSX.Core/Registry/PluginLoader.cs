/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
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
using NanoXLSX.Interfaces.Writer;
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
        /// Signature "MZ" at the start of the DOS header of a PE image
        /// </summary>
        private const ushort DOS_SIGNATURE = 0x5A4D;
        /// <summary>
        /// Position of the DOS header field (e_lfanew) that holds the position of the PE header
        /// </summary>
        private const int PE_HEADER_POINTER_POSITION = 0x3C;
        /// <summary>
        /// Signature "PE\0\0" at the start of the PE header
        /// </summary>
        private const uint PE_SIGNATURE = 0x00004550;
        /// <summary>
        /// Offset of the optional header from the start of the PE header (PE signature and COFF file header)
        /// </summary>
        private const int OPTIONAL_HEADER_OFFSET = 24;
        /// <summary>
        /// Magic number of a PE32 optional header
        /// </summary>
        private const ushort PE32_MAGIC = 0x10B;
        /// <summary>
        /// Magic number of a PE32+ optional header
        /// </summary>
        private const ushort PE32_PLUS_MAGIC = 0x20B;
        /// <summary>
        /// Offset of the field NumberOfRvaAndSizes in a PE32 optional header. The data directories follow this field
        /// </summary>
        private const int PE32_RVA_COUNT_OFFSET = 92;
        /// <summary>
        /// Offset of the field NumberOfRvaAndSizes in a PE32+ optional header. The data directories follow this field
        /// </summary>
        private const int PE32_PLUS_RVA_COUNT_OFFSET = 108;
        /// <summary>
        /// Index of the CLI header data directory, which is only present in managed assemblies
        /// </summary>
        private const int CLI_HEADER_DIRECTORY_INDEX = 14;
        /// <summary>
        /// Length of a data directory entry (RVA and size)
        /// </summary>
        private const int DATA_DIRECTORY_LENGTH = 8;

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
            Assembly[] loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();

            // 1. Register plugins from already-loaded assemblies
            foreach (Assembly assembly in loadedAssemblies)
            {
                if (!assembly.IsDynamic && !string.IsNullOrEmpty(assembly.Location))
                {
                    try
                    {
                        RegisterPlugIns(assembly);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to register plugins from {assembly.Location}: {ex.Message}");
                    }
                }
            }

            // 2. Load any remaining managed DLLs that haven't been loaded yet
            IEnumerable<string> allLoadedPaths = loadedAssemblies
                .Where(a => !a.IsDynamic && !string.IsNullOrEmpty(a.Location))
                .Select(a => a.Location);
            HashSet<string> loadedPaths = new HashSet<string>(allLoadedPaths, StringComparer.InvariantCultureIgnoreCase);

            // Guard the assembly-directory enumeration. Directory.GetFiles can throw (e.g. IOException
            // "The parameter is incorrect" when BaseDirectory contains an entry the native enumerator
            // cannot stat). Phase 1 above already registered plugins from loaded assemblies,
            // so on enumeration failure, the already processed result is returned rather than crashing.
            List<string> referencedPaths;
            try
            {
                referencedPaths = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.dll")
                    .Where(path => !loadedPaths.Contains(path))
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to enumerate assemblies in {AppDomain.CurrentDomain.BaseDirectory}: {ex.Message}");
                return;
            }

            foreach (string path in referencedPaths)
            {
                try
                {
                    if (IsManagedAssembly(path))
                    {
                        Assembly assembly = Assembly.LoadFrom(path);
                        RegisterPlugIns(assembly);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to load assembly: {path} - {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Method to determine whether a file is a managed (.NET) assembly, by reading its PE header directly. A file is a managed assembly if it has a DOS and a PE signature, a PE32 or PE32+ optional header and a non-empty CLI header data directory
        /// </summary>
        /// <param name="path">Path of the file to check</param>
        /// <returns>True if the file is a managed assembly, otherwise false (e.g. a native DLL, or a file that is not a PE image or ends within its headers)</returns>
        /// \remark <remarks>The file is read as plain bytes because some runtimes (e.g. Mono) write an error to the host's log before throwing a BadImageFormatException when a native DLL is passed to Assembly.LoadFrom or AssemblyName.GetAssemblyName, and hosts like the Unity editor report each of these log entries as an error. IO exceptions (e.g. a locked file) are passed to the caller</remarks>
        internal static bool IsManagedAssembly(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                long length = stream.Length;
                if (length < PE_HEADER_POINTER_POSITION + sizeof(int) || reader.ReadUInt16() != DOS_SIGNATURE)
                {
                    return false;
                }

                stream.Position = PE_HEADER_POINTER_POSITION;
                long peHeaderPosition = reader.ReadInt32();
                long optionalHeaderPosition = peHeaderPosition + OPTIONAL_HEADER_OFFSET;
                if (peHeaderPosition < 0 || optionalHeaderPosition + sizeof(ushort) > length)
                {
                    return false;
                }
                stream.Position = peHeaderPosition;
                if (reader.ReadUInt32() != PE_SIGNATURE)
                {
                    return false;
                }

                stream.Position = optionalHeaderPosition;
                ushort magic = reader.ReadUInt16();
                long rvaCountPosition;
                if (magic == PE32_MAGIC)
                {
                    rvaCountPosition = optionalHeaderPosition + PE32_RVA_COUNT_OFFSET;
                }
                else if (magic == PE32_PLUS_MAGIC)
                {
                    rvaCountPosition = optionalHeaderPosition + PE32_PLUS_RVA_COUNT_OFFSET;
                }
                else
                {
                    return false;
                }

                long cliHeaderDirectoryPosition = rvaCountPosition + sizeof(uint) + (CLI_HEADER_DIRECTORY_INDEX * DATA_DIRECTORY_LENGTH);
                if (cliHeaderDirectoryPosition + DATA_DIRECTORY_LENGTH > length)
                {
                    return false;
                }
                stream.Position = rvaCountPosition;
                if (reader.ReadUInt32() <= CLI_HEADER_DIRECTORY_INDEX)
                {
                    return false;
                }
                stream.Position = cliHeaderDirectoryPosition;
                uint cliHeaderRva = reader.ReadUInt32();
                uint cliHeaderSize = reader.ReadUInt32();
                return cliHeaderRva != 0 && cliHeaderSize != 0;
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
        /// <returns>List of class types, matching the criteria</returns>
        [ExcludeFromCodeCoverage] // Indirectly tested by InjectPlugins
        private static List<Type> GetAssemblyPlugInsByType(Assembly assembly, Type attributeType)
        {
            List<Type> plugInTypes = new List<Type>();
            Type plugInInterface = typeof(IPlugin);
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
                        if (attribute.PlugInOrder >= plugInClasses[attribute.PlugInUUID].Order)
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
                return (T)Activator.CreateInstance(plugIn.Type, true);
            }
            else
            {
                return fallBackInstance;
            }
        }

        /// <summary>
        /// Indicates whether at least one queue plug-in is registered for the given queue UUID.
        /// </summary>
        /// <param name="queueUUID">UUID of the queue to check</param>
        /// <returns>True if at least one plug-in is registered for the queue, otherwise false</returns>
        internal static bool HasQueuePlugins(string queueUUID)
        {
            return queuePlugInClasses.TryGetValue(queueUUID, out var list) && list.Count > 0;
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
        /// \remark <remarks>This method is not intended to manage preserved plugin instances. When called, always a new instance of the plugin will be created</remarks>
        internal static T GetNextQueuePlugIn<T>(string queueUUID, string lastPlugInUUID, out string currentPlugInUUID)
            where T : class, IPlugin
        {
            return GetNextQueuePlugIn<T>(queueUUID, lastPlugInUUID, out currentPlugInUUID, null);
        }

        /// <summary>
        /// Method to get the next instance of a queue plug-in, providing a writer context to plug-ins that support contextual construction
        /// </summary>
        /// <typeparam name="T">Plug-in type</typeparam>
        /// <param name="queueUUID">UUID of the queue</param>
        /// <param name="lastPlugInUUID">UUID of the last plug-in instance, to determine the next one</param>
        /// <param name="currentPlugInUUID">Out parameter that returns the UUID of the determined next plug-in instance</param>
        /// <param name="baseWriter">Writer context supplied to a compatible constructor</param>
        /// <returns>Plug-in instance or null, if the end of the queue was reached</returns>
        internal static T GetNextQueuePlugIn<T>(string queueUUID, string lastPlugInUUID, out string currentPlugInUUID, IBaseWriter baseWriter)
            where T : class, IPlugin
        {
            currentPlugInUUID = null;
            if (!queuePlugInClasses.TryGetValue(queueUUID, out var plugInList) || plugInList.Count == 0)
            {
                return default;
            }

            int startIndex = 0;
            if (lastPlugInUUID != null)
            {
                int lastIndex = plugInList.FindIndex(p => p.UUID == lastPlugInUUID);
                if (lastIndex < 0)
                {
                    return default;
                }

                startIndex = lastIndex + 1;
            }

            Type requestedType = typeof(T);
            for (int i = startIndex; i < plugInList.Count; i++)
            {
                PlugInInstance plugIn = plugInList[i];
                if (!requestedType.IsAssignableFrom(plugIn.Type))
                {
                    continue;
                }

                object instance = null;
                if (baseWriter != null)
                {
                    ConstructorInfo contextualConstructor = plugIn.Type
                        .GetConstructors(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                        .FirstOrDefault(constructor =>
                        {
                            ParameterInfo[] parameters = constructor.GetParameters();
                            return parameters.Length == 1 && parameters[0].ParameterType == typeof(IBaseWriter);
                        });
                    if (contextualConstructor != null)
                    {
                        instance = contextualConstructor.Invoke(new object[] { baseWriter });
                    }
                }

                if (instance == null)
                {
                    instance = Activator.CreateInstance(plugIn.Type, true);
                }

                currentPlugInUUID = plugIn.UUID;
                return (T)instance;
            }

            return default;
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
            /// \remark <remarks>Plug-ins must provide either an empty constructor or a supported context-aware constructor</remarks>
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
