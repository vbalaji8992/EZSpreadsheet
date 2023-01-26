using Microsoft.VisualStudio.TestPlatform.Utilities;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public static string GetFormattedFile(string expectedStylesheet)
        {
            return expectedStylesheet.Replace("\n", "")
                .Replace("\r", "")
                .Replace("\t", "")
                .Replace(" ", "");
        }

        public static List<string> ExtractFiles(string file, string extractPath)
        {
            DeleteFolder(extractPath);
            ZipFile.ExtractToDirectory(file, extractPath);

            var extractedFiles = Directory.GetFiles(extractPath, "*.*", SearchOption.AllDirectories).ToList();
                       
            return extractedFiles;
        }

        public static void AssertFiles(string expectedFile, string actualFile)
        {
            var expectedFileFormatted = GetFormattedFile(File.ReadAllText(expectedFile));
            var actualFileFormatted = File.ReadAllText(actualFile).Replace(" ", "");
            Assert.Equal(expectedFileFormatted, actualFileFormatted);
        }
    }
}
