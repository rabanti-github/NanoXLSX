using System.Runtime.CompilerServices;

// Core packages
[assembly: InternalsVisibleTo("NanoXLSX.Writer")]
[assembly: InternalsVisibleTo("NanoXLSX.Reader")]
// Test packages
[assembly: InternalsVisibleTo("NanoXLSX.Core.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Writer-Reader.Test")]
// Plug-ins (backlog / reserved)
[assembly: InternalsVisibleTo("NanoXLSX.Formula")]
[assembly: InternalsVisibleTo("NanoXLSX.Security")]
[assembly: InternalsVisibleTo("NanoXLSX.Formatting")]
