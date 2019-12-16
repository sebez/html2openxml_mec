using System.IO;
using NUnit.Framework;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests Horizontal Lines.
    /// </summary>
    [TestFixture]
    public class MergeTableTests : FromDocumentTestBase
    {
        [Test]
        public void ParseTableMerge()
        {
            // Arrange
            var html = File.ReadAllText(@"D:\ProjetsGit\html2openxml_mec\html2openxml\html2openxml.test\Resources\TableMerge.html");

            // Act
            Check(html);
        }

        [Test]
        public void ParseTableMergeSimple()
        {
            // Arrange
            var html = File.ReadAllText(@"D:\ProjetsGit\html2openxml_mec\html2openxml\html2openxml.test\Resources\TableMergeSimple.html");

            // Act
            Check(html);
        }

        [Test]
        public void ParseTableMergeDouble()
        {
            // Arrange
            var html = File.ReadAllText(@"D:\ProjetsGit\html2openxml_mec\html2openxml\html2openxml.test\Resources\TableMergeDouble.html");

            // Act
            Check(html);
        }

        [Test]
        public void ParseTableMergeRight()
        {
            // Arrange
            var html = File.ReadAllText(@"D:\ProjetsGit\html2openxml_mec\html2openxml\html2openxml.test\Resources\TableMergeRight.html");

            // Act
            Check(html);
        }

        [Test]
        public void ParseTableMergeV1()
        {
            // Arrange
            var html = File.ReadAllText(@"D:\ProjetsGit\html2openxml_mec\html2openxml\html2openxml.test\Resources\TableMergeV1.html");

            // Act
            Check(html);
        }
    }
}