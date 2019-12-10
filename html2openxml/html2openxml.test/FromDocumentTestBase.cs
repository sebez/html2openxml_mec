using System;
using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    public abstract class FromDocumentTestBase
    {
        protected void Check(string html)
        {
            var content = Convert(html);
            string path = $@"D:\temp\Dalkia\20191210\D{DateTime.Now:HHmmss}.docx";
            File.WriteAllBytes(path, content);
            Process.Start(path);
        }

        public static byte[] Convert(string mainContenthtml)
        {
            /* Charge le template vierge en mémoire. */
            var templatePath = @"D:\Projets\MEC\Main0\Sources\Dalkia.MEC.Wpfapp\Dalkia.MEC.Core\Editique\Generation\TemplateGrosSite.docx";
            var templateContent = File.ReadAllBytes(templatePath);

            using (var ms = new MemoryStream())
            {
                /* Ecrit le contenu du template vierge dans le flux. */
                ms.Write(templateContent, 0, templateContent.Length);
                ms.Seek(0, SeekOrigin.Begin);

                /* Instancie un document OpenXml avec le flux. */
                using (var wDoc = WordprocessingDocument.Open(ms, isEditable: true))
                {
                    /* Initialise un document Word vide. */
                    MainDocumentPart mainPart = wDoc.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = wDoc.AddMainDocumentPart();
                        new Document(new[] { new Body() }).Save(mainPart); // ?
                    }

                    HtmlConverter htmlConverter = new HtmlConverter(mainPart);

                    /* Parse le HTML du contenu principal dans le document Word. */
                    htmlConverter.ParseHtml(mainContenthtml);

                    /* Enregistre le document Word. */
                    mainPart.Document.Save();
                }

                /* Retourne le contenu binaire du document Word. */
                return ms.ToArray();
            }
        }
    }
}