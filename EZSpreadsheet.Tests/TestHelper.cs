using DocumentFormat.OpenXml.Vml.Office;
using Microsoft.VisualStudio.TestPlatform.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit.Sdk;

namespace EZSpreadsheet.Tests
{
    public static class TestHelper
    {
        public const string EXPECTED_XML_FOLDER = "Expected-XML";

        public static string GetFormattedXml(string xml)
        {
            return xml.Replace("\n", "")
                .Replace("\r", "")
                .Replace("\t", "")
                .Replace(" ", "");
        }

        public static void AssertXml(string expectedXmlPath, string actualXmlPath, Stream stream)
        {
            var archive = new ZipArchive(stream);
            var actualFile = archive.Entries.Where(x => x.FullName == actualXmlPath).FirstOrDefault();

            if (actualFile == null)
                throw new XunitException("Zip file does not contain the specified file");

            var reader = new StreamReader(actualFile.Open());
            var actualXml = reader.ReadToEnd();

            var expectedXmlFormatted = GetFormattedXml(File.ReadAllText(expectedXmlPath));
            var actualXmlFormatted = GetFormattedXml(actualXml);

            Assert.Equal(expectedXmlFormatted, actualXmlFormatted);
        }
    }
}
