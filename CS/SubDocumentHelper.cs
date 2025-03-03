using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SubDocumentIterator
{
    public static class SubDocumentHelper
    {
        public delegate void SubDocumentDelegate(SubDocument subDocument);
        public static void ForEachSubDocument(this Document document, SubDocumentDelegate subDocumentProcessor)
        {
            subDocumentProcessor(document);
            ProcessShapes(document.Shapes, subDocumentProcessor);
            ProcessComments(document.Comments, subDocumentProcessor);
            foreach (Section section in document.Sections)
            {
                ProcessSection(section, HeaderFooterType.First, subDocumentProcessor);
                ProcessSection(section, HeaderFooterType.Odd, subDocumentProcessor);
                ProcessSection(section, HeaderFooterType.Even, subDocumentProcessor);
            }
        }
        private static void ProcessSection(Section section, HeaderFooterType headerFooterType, SubDocumentDelegate subDocumentProcessor)
        {
            if (section.HasHeader(headerFooterType))
            {
                SubDocument header = section.BeginUpdateHeader(headerFooterType);
                subDocumentProcessor(header);
                ProcessShapes(header.Shapes, subDocumentProcessor);
                section.EndUpdateHeader(header);
            }
            if (section.HasFooter(headerFooterType))
            {
                SubDocument footer = section.BeginUpdateFooter(headerFooterType);
                subDocumentProcessor(footer);
                ProcessShapes(footer.Shapes, subDocumentProcessor);
                section.EndUpdateFooter(footer);
            }
        }
        private static void ProcessShapes(ShapeCollection shapes, SubDocumentDelegate subDocumentProcessor)
        {
            foreach (Shape shape in shapes)
                if (shape.ShapeFormat.TextBox != null)
                    subDocumentProcessor(shape.ShapeFormat.TextBox.Document);
        }
        private static void ProcessComments(CommentCollection comments, SubDocumentDelegate subDocumentProcessor)
        {
            foreach (Comment comment in comments)
            {
                SubDocument commentSubDocument = comment.BeginUpdate();
                subDocumentProcessor(commentSubDocument);
                comment.EndUpdate(commentSubDocument);
            }
        }
    }
}
