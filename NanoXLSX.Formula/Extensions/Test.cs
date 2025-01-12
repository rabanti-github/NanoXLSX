using System;
using NanoXLSX.Interfaces.Writer;

namespace NanoXLSX.Extensions.Formula
{
    public class Test : IPluginWriter
    {
        public Workbook Workbook { get; set; }
        public IPluginWriter NextWriter { get; set; }

        // static Test()
        // {
        //     Test instance = new Test();
        //     PackageRegistry.RegisterWriterPlugin(instance.GetClassID(), instance);
        // }



        public virtual string CreateDocument(string currentDocument = null)
        {
            return @"
  <Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">  <TotalTime>0</TotalTime>
  <Application>Microsoft Excel LeRab edition</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company>
  </Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
  </Properties>
            ";
        }

        public string GetClassID()
        {
            return "A73923A8-1E7E-4673-AD3F-B22DD3153D7B";
        }

        public void PostWrite(Workbook workbook)
        {
            throw new NotImplementedException();
        }


        public void PreWrite(Workbook workbook)
        {
            throw new NotImplementedException();
        }

    }
}
