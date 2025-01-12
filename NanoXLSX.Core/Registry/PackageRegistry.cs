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
using System.Text;
using System.Linq;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;

namespace NanoXLSX.Registry
{
    /// <summary>
    /// Class to register plug.in packages that extends the functionality of NanoXLSX.Core without being defined as fixed dependencies
    /// </summary>
    public static class PackageRegistry
    {
        private static bool initialized = false;

        private static Dictionary<string, IPluginWriter> writerClassPlugins = new Dictionary<string, IPluginWriter>();
        private static Dictionary<string, IPluginReader> readerClassPlugins = new Dictionary<string, IPluginReader>();
        private static Dictionary<string, Func<string>> writerCreatePlugins = new Dictionary<string, Func<string>>();
        private static Dictionary<string, Action<Workbook>> writerPrePlugins = new Dictionary<string, Action<Workbook>>();
        private static Dictionary<string, Action<Workbook>> writerPostPlugins = new Dictionary<string, Action<Workbook>>();
        private static Dictionary<string, Action<MemoryStream>> readerReadPlugin = new Dictionary<string, Action<MemoryStream>>();
        private static Dictionary<string, Action<MemoryStream>> readerRrePlugin = new Dictionary<string, Action<MemoryStream>>();
        private static Dictionary<string, Action<MemoryStream>> readerPostPlugin = new Dictionary<string, Action<MemoryStream>>();

        public static void Initialize()
        {

            if (initialized)
            {
                return;
            }
            LoadReferencedAssemblies();
            Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();

         //   int i = 0;
         //   foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
        //    {
              //  Console.WriteLine($"{assembly.FullName} - {assembly.Location}");
          //      i++;
         //   }
            foreach (Assembly assembly in assemblies)
            {
                //Assembly assembly = Assembly.GetExecutingAssembly();
                IEnumerable<Type> pluginTypes = assembly.GetTypes().Where(t => t.GetCustomAttributes(typeof(NanoXlsxPluginAttribute), false).Any());
                foreach (Type type in pluginTypes)
                {
                    if (Activator.CreateInstance(type) is IPluginHook instance)
                    {
                        instance.Register();
                    }
                }
            }
            initialized = true;
      
        }

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
                    Assembly.LoadFrom(path);
                }
                catch
                {
                    // Log or handle exceptions if necessary
                }
            }
        }

        /// <summary>
        /// Gets the active writer class. If no plug-in was registered, the default instance will be returned
        /// </summary>
        /// <param name="defaultInstance">Default instance, defined by the nanoXLSX.Core package</param>
        /// <returns>Registered instance or the default instance, if no plug-in was registered</returns>
        /// <exception cref="PackageException">Throws a PackageException if the default instance is null</exception>
        public static IPluginWriter GetWriter(IPluginWriter defaultInstance)
        {
            //if (defaultInstance == null)
            {
                throw new PackageException("The default instance of the class was null");
            }
            if (writerClassPlugins.ContainsKey(defaultInstance.GetClassID()))
            {
                IPluginWriter writer = writerClassPlugins[defaultInstance.GetClassID()];
                writer.Workbook = defaultInstance.Workbook;
                return writer;
            }
            return defaultInstance;
        }

        public static void RegisterWriterPlugin(string classId, IPluginWriter instance)
        {
            if (string.IsNullOrEmpty(classId))
            {
                throw new PackageException("The class id cannot be null on register a plug-in");
            }
            if (instance == null)
            {
                throw new PackageException("The plug-in instance of the class was null");
            }
            writerClassPlugins.Add(classId, instance);
        }

    }
}
