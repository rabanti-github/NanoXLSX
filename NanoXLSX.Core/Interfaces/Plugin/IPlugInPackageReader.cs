using NanoXLSX.Interfaces.Reader;

namespace NanoXLSX.Interfaces.Plugin
{
    internal interface IPluginPackageReader : IPluginReader
    {
        /// <summary>
        /// Relative path of the stream entry in the Zip archive
        /// </summary>
        string StreamEntryName { get; }
    }
}
