using GemBox.Document;
using GemBox.Document.MailMerging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace GemBoxMergeFieldIssues
{
    class Program
    {
        static void Main(string[] args)
        {
            var currentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var word = DocumentModel.Load(Path.Combine(currentPath, "doc.docx"), LoadOptions.DocxDefault);

            var model = new MergeModel();
            Merge(word, model);

            word.Save(Path.Combine(currentPath, "output.html"), SaveOptions.HtmlDefault);
        }

        private static void Merge(DocumentModel word, MergeModel model)
        {
            var fields = word.MailMerge.GetMergeFieldNames();

            word.MailMerge.Execute(model);
        }
    }

    public class MergeModel
    {
        public MergeModel()
        {
            Details = new List<Detail>() {
                    new Detail("AAA1", "AAA2"),
                    new Detail("BBB1", "BBB2"),
                    new Detail("CCC1", "CCC2")
                };
        }

        public string Link { get { return $"https://www.google.com"; } }

        public List<Detail> Details { get; set; }
    }

    public class Detail
    {
        public Detail(string cell1, string cell2) { }
        public string Cell1 { get; set; }
        public string Cell2 { get; set; }
    }

}
