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
        public const string TEST_OUTPUT_FOLDER = "Test-Output";
        public const string EXPECTED_XML_FOLDER = "Expected-XML";

        public static void CreateFolder(string folder)
        {
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
        }

        public static void DeleteFolder(string folder)
        {
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }
        }

        public static string GetFormattedXml(string xml)
        {
            return xml.Replace("\n", "")
                .Replace("\r", "")
                .Replace("\t", "")
                .Replace(" ", "");
        }

        public static List<string> ExtractFiles(string file, string extractPath)
        {
            DeleteFolder(extractPath);
            ZipFile.ExtractToDirectory(file, extractPath);

            var extractedFiles = Directory.GetFiles(extractPath, "*.*", SearchOption.AllDirectories).ToList();
                       
            return extractedFiles.Select(x => x.Replace("\\", "/")).ToList();
        }

        public static void AssertFile(string expectedFile, string actualFile)
        {
            var expectedFileFormatted = GetFormattedXml(File.ReadAllText(expectedFile));
            var actualFileFormatted = File.ReadAllText(actualFile).Replace(" ", "");
            Assert.Equal(expectedFileFormatted, actualFileFormatted);
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
