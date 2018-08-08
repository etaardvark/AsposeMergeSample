using Aspose.Words;
using Aspose.Words.Replacing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace AsposeMergeSample
{
    public class FieldPlaceholderMerger : IReplacingCallback
    {
        private int _mergeReplacePosition;
        int ifLevel = 0;
        public int maxIf = 0;
        Stack<int> currentLevel = new Stack<int>();
        private IEnumerable<XElement> _inserts;
        private Document _doc;
        private List<Node> _blankReplacements = new List<Node>();

        DocumentBuilder builder;

        public FieldPlaceholderMerger(Document doc)
        {
            _doc = doc;
            builder = new DocumentBuilder(doc);
        }

        public void ReplacePlaceholderWithInserts(IEnumerable<XElement> inserts)
        {
            _inserts = inserts;

            // Replace the Sections first, then replace the headers and footers afterwards
            foreach (Section section in _doc.Sections)
            {
                Dictionary<HeaderFooterType, List<Node>> headerText = new Dictionary<HeaderFooterType, List<Node>>();
                // Remove the Headers from the section
                HeaderFooterCollection headerFooterCollection = section.HeadersFooters;
                foreach (HeaderFooter header in headerFooterCollection)
                {
                    List<Node> paraCol = new List<Node>();
                    foreach (Paragraph para in header.Paragraphs)
                    {
                        paraCol.Add(para.Clone(true));
                    }
                    headerText.Add(header.HeaderFooterType, paraCol);
                }

                section.ClearHeadersFooters();
                // Replace this section
                section.Range.Replace(new Regex(CompiledMarkupConstants._FIELD_PLACEHOLDER), this, true);
                // Add the Headers back for this section
                foreach (HeaderFooter header in section.HeadersFooters)
                {
                    if (headerText.ContainsKey(header.HeaderFooterType))
                    {
                        IEnumerable<Node> list = (IEnumerable<Node>)headerText[header.HeaderFooterType];
                        list.Reverse();
                        foreach (Node node in list)
                        {
                            header.AppendChild(node);
                        }
                    }
                }
            }
            foreach (Section section in _doc.Sections)
            {
                HeaderFooter header = null;


                // Process Header / Footers in Exact order for Visualfiles
                header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header != null)
                {
                    header.Range.Replace(new Regex(CompiledMarkupConstants._FIELD_PLACEHOLDER), this, true);
                    header = null;
                }
                header = section.HeadersFooters[HeaderFooterType.HeaderFirst];
                if (header != null)
                {
                    header.Range.Replace(new Regex(CompiledMarkupConstants._FIELD_PLACEHOLDER), this, true);
                    header = null;
                }

                header = section.HeadersFooters[HeaderFooterType.HeaderEven];
                if (header != null)
                {
                    header.Range.Replace(new Regex(CompiledMarkupConstants._FIELD_PLACEHOLDER), this, true);
                    header = null;
                }
                header = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (header != null)
                {
                    header.Range.Replace(new Regex(CompiledMarkupConstants._FIELD_PLACEHOLDER), this, true);
                    header = null;
                }
                header = section.HeadersFooters[HeaderFooterType.FooterFirst];
                if (header != null)
                {
                    header.Range.Replace(new Regex(CompiledMarkupConstants._FIELD_PLACEHOLDER), this, true);
                    header = null;
                }
                header = section.HeadersFooters[HeaderFooterType.FooterEven];
                if (header != null)
                {
                    header.Range.Replace(new Regex(CompiledMarkupConstants._FIELD_PLACEHOLDER), this, true);
                    header = null;
                }



            }

            // Replace any Blank replacements to remove formatting around 'blank' tags
            foreach (Node node in _blankReplacements)
            {
                RemoveEmptyNode(node);
            }
        }

        private void RemoveEmptyNode(Node node)
        {
            var parent = node.ParentNode;
            if (parent != null)
            {
                if (node.GetText().Trim() == string.Empty)
                {
                    node.Remove();
                }
                if (parent.ToString(SaveFormat.Text).Trim() == string.Empty)
                {
                    parent.Remove();
                }
            }
        }


        public ReplaceAction Replacing(ReplacingArgs args)
        {
            _mergeReplacePosition++;

            var insert = _inserts.FirstOrDefault(p => p.Attribute(SchemaConstants._INSERTS_ID).Value == _mergeReplacePosition.ToString());
            if (insert != null)
            {

                if (insert.Value == "{{{")
                {

                    ifLevel++;
                    maxIf = Math.Max(ifLevel, maxIf);
                    insert.Value = "{{" + string.Format(@"{0}", ifLevel) + "{";
                    currentLevel.Push(ifLevel);

                }
                if (insert.Value == "}}}")
                {
                    insert.Value = "}}" + string.Format("{0}", currentLevel.Pop() + "}");
                    Run currentRun = args.MatchNode as Run;
                    Paragraph paragraph = currentRun.ParentParagraph;



                    // Reduce the indent level                    
                }
                args.Replacement = insert.Value;
            }
            else
            {
                _blankReplacements.Add(args.MatchNode);
            }
            return ReplaceAction.Replace;
        }
    }
}
