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

            var model = new MergeModel()
            {
                BaseObject = BuildDataSource()
            };

            Merge(word, model);

            word.Save(Path.Combine(currentPath, "output.html"), SaveOptions.HtmlDefault);
        }

        private static void Merge(DocumentModel word, MergeModel model)
        {
            word.MailMerge.FieldMerging += (sender, e) =>
            {
            };

            word.MailMerge.Execute(model);
        }

        private static BaseObject BuildDataSource()
        {
            var bo = new BaseObject();

            var co = new ChildObject() { Name = "Child1" };
            co.Keys.Add("Color", "Red");
            co.Keys.Add("Shape", "Circle");
            co.Keys.Add("Size", "Small");
            bo.Children.Add(co);

            co = new ChildObject() { Name = "Child2" };
            co.Keys.Add("Color", "Green");
            co.Keys.Add("Shape", "Square");
            co.Keys.Add("Size", "Small");
            bo.Children.Add(co);

            co = new ChildObject() { Name = "Child3" };
            co.Keys.Add("Color", "Blue");
            co.Keys.Add("Shape", "Circle");
            co.Keys.Add("Size", "Large");
            bo.Children.Add(co);

            return bo;
        }
    }

    public class MergeModel
    {
        public BaseObject BaseObject { get; set; }
    }

    public class BaseObject
    {
        public BaseObject()
        {
            Children = new List<ChildObject>();
        }

        public List<ChildObject> Children { get; set; }
    }

    public class ChildObject
    {
        public ChildObject()
        {
            Keys = new Dictionary<string, string>();
        }

        public string Name { get; set; }
        public Dictionary<string, string> Keys { get; set; }
    }
}
