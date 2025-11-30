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
// Plug-ins (backlog / reserved)
[assembly: InternalsVisibleTo("NanoXLSX.Formula")]
[assembly: InternalsVisibleTo("NanoXLSX.Security")]
[assembly: InternalsVisibleTo("NanoXLSX.Formatting")]
[assembly: InternalsVisibleTo("NanoXLSX.Data")]
[assembly: InternalsVisibleTo("NanoXLSX.Essentials")]
[assembly: InternalsVisibleTo("NanoXLSX.Automatization")]
[assembly: InternalsVisibleTo("NanoXLSX.Chart")]
[assembly: InternalsVisibleTo("NanoXLSX.Compatibility")]
// Plug-in Tests (backlog / reserved)
[assembly: InternalsVisibleTo("NanoXLSX.Formula.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Security.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Formatting.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Essentials.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Automatization.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Chart.Test")]
[assembly: InternalsVisibleTo("NanoXLSX.Compatibility.Test")]
