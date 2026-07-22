using System.Runtime.CompilerServices;

// Meta packages
[assembly: InternalsVisibleTo("NanoXLSX")]
[assembly: InternalsVisibleTo("PicoXLSX")]

// Core packages
[assembly: InternalsVisibleTo("NanoXLSX.Writer")]
[assembly: InternalsVisibleTo("NanoXLSX.Reader")]
// Included test packages
[assembly: InternalsVisibleTo("NanoXLSX.Core.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Writer-Reader.Test")]

// Plug-ins (existing)
[assembly: InternalsVisibleTo("NanoXLSX.Formatting")]
// Plug-in Tests (existing)
[assembly: InternalsVisibleTo("NanoXLSX.Formatting.Test")]

// Plug-ins (backlog / reserved)
[assembly: InternalsVisibleTo("NanoXLSX.Formula")]
[assembly: InternalsVisibleTo("NanoXLSX.Security")]
[assembly: InternalsVisibleTo("NanoXLSX.Data")]
[assembly: InternalsVisibleTo("NanoXLSX.Essentials")]
[assembly: InternalsVisibleTo("NanoXLSX.Automation")]
[assembly: InternalsVisibleTo("NanoXLSX.Chart")]
[assembly: InternalsVisibleTo("NanoXLSX.Compatibility")]
// Plug-in Tests (backlog / reserved)
[assembly: InternalsVisibleTo("NanoXLSX.Formula.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Security.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Data.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Essentials.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Automation.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Chart.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Compatibility.Test")]
