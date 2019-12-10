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
    }
}