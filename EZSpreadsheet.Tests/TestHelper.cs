using DocumentFormat.OpenXml.Vml.Office;
using Microsoft.VisualStudio.TestPlatform.PlatformAbstractions.Interfaces;
using Microsoft.VisualStudio.TestPlatform.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace EZSpreadsheet.Tests
{
    public static class TestHelper
    {
        public const string EXPECTED_FILES_FOLDER = "ExpectedFiles";

        public static string GetFormattedXml(string xml)
        {
            return xml.Replace("\n", "")
                .Replace("\r", "")
                .Replace("\t", "")
                .Replace(" ", "");
        }

        public static void AssertSpreadsheet(Stream actualFileStream, string expectedFilePath, bool saveTestFile = false)
        {
            if (saveTestFile)
                SaveTestGeneratedFile(actualFileStream, expectedFilePath.Split("/")[1]);

            using var expectedArchive = ZipFile.OpenRead(expectedFilePath);
            var ignoredFiles = new List<string>()
            {
                "_rels/.rels",
                "xl/_rels/workbook.xml.rels",
                "xl/workbook.xml"
            };
            var expectedXmlFiles = expectedArchive.Entries.Where(x => !ignoredFiles.Contains(x.FullName));

            using var actualArchive = new ZipArchive(actualFileStream);

            foreach (var file in expectedXmlFiles)
            {
                var actualFile = actualArchive.Entries.Where(x => x.FullName == file.FullName).FirstOrDefault();

                if (actualFile == null)
                    throw new XunitException("Zip file does not contain the specified file");

                var expectedXmlFormatted = ReadArchiveFile(file);
                var actualXmlFormatted = ReadArchiveFile(actualFile);

                Assert.Equal(expectedXmlFormatted, actualXmlFormatted);
            }
        }

        private static string ReadArchiveFile(ZipArchiveEntry file)
        {
            var reader = new StreamReader(file.Open());
            return reader.ReadToEnd();
        }

        private static void SaveTestGeneratedFile(Stream testFileStream, string outputFileName)
        {
            if (!Directory.Exists("TestGeneratedFiles"))
                return;

            string destFileName = $"TestGeneratedFiles/{outputFileName}";
            if (File.Exists(destFileName)) 
                File.Delete(destFileName);
            
            using var destArchive = ZipFile.Open(destFileName, ZipArchiveMode.Create);

            using var sourceArchive = new ZipArchive(testFileStream, ZipArchiveMode.Read, true);

            foreach (var entry in sourceArchive.Entries)
            {
                using var existinFileStream = entry.Open();
                var newFile = destArchive.CreateEntry(entry.FullName);
                using var newFileStream = newFile.Open();
                existinFileStream.CopyTo(newFileStream);
            }
        }
    }
}
