using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AsposeMergeSample
{
    public class ContentFinder
    {
        private Document _doc;

        public ContentFinder(Document doc)
        {
            _doc = doc;
        }

        public IList<Node> FindTextRange(string startText, string endText)
        {
            var textFinder = new GenericTextFinder();

            _doc.Range.Replace(new Regex(Regex.Escape(startText) + "(.*?)" +
                                         Regex.Escape(endText), RegexOptions.IgnoreCase), textFinder, false);

            return textFinder.Nodes;
        }

        /// <summary>
        ///  Find all the matches of {{{ }}} exclusion tags
        /// </summary>

        public IList<Node> FindExclusionTags(int level)
        {
            var textFinder = new EndRangeTextFinder();

            //_doc.Range.Replace(new Regex(@"({{{).*([^{{{]}}}){1}", RegexOptions.Multiline), textFinder, false);

            textFinder.EndRange = "}}" + string.Format("{0}", level) + "}";
            string regex = "({{" + string.Format("{0}", level) + ").*(}}" + string.Format("{0}", level) + "})";
            _doc.Range.Replace(new Regex(regex, RegexOptions.Multiline), textFinder, false);

            return textFinder.Nodes;

        }

        public IList<Node> FindAllMatchingTextRanges(string startText, string endText)
        {
            var textFinder = new GenericTextFinder();
            _doc.Range.Replace(new Regex("(?<=" + Regex.Escape(startText) + ").*(?=" +
                                         Regex.Escape(endText) + ")", RegexOptions.IgnoreCase |
                                                                RegexOptions.Singleline), textFinder, false);

            return textFinder.Nodes;

        }


        private Node GetChildNode(Paragraph parentNode, string searchString)
        {
            Node currentNode = parentNode.FirstChild;
            while (currentNode != null)
            {
                // If this node only contains the requested text
                if (currentNode.GetText().Trim() == searchString)
                {
                    return currentNode;
                }
                // If this node ends with the text - i.e. cover the node containing both the start / end tag in the single node.
                if (currentNode.GetText().Trim().EndsWith(searchString))
                {
                    return currentNode;
                }

                currentNode = currentNode.NextSibling;
            }
            return null;
        }

        /// <summary>
        /// Method to find all the 'ranges' with the following start / end matches
        /// Will return *all* ranges that match the enclosing tags
        /// 
        /// e.g. foreach tags with their matching endfor tag
        /// </summary>
        /// <param name="startTerm">Starting term for the match</param>
        /// <param name="endTerm">Ending term for the match</param>
        /// <returns>List of the matching ranges within the document</returns>
        public IList<NodeRangeMatch> FindMatchingNodes(string startTerm, string endTerm)
        {
            List<NodeRangeMatch> returnList = new List<NodeRangeMatch>();
            NodeCollection allNodes = _doc.GetChildNodes(NodeType.Run, true);

            Node startNode = null;
            Node endNode = null;
            // Search for StartNode
            foreach (Run node in allNodes)
            {
                string s = node.GetText();
                if (node.GetText().Trim().StartsWith(startTerm))
                {
                    startNode = node;
                }
                if (node.GetText().Trim().IndexOf(startTerm) > 0)
                {
                    startNode = node;
                }
                if (node.GetText().Trim().Contains(endTerm))
                {

                    endNode = node;
                    if (startNode != null)
                    {
                        NodeRangeMatch range = new NodeRangeMatch();
                        range.StartNode = startNode;
                        range.EndNode = endNode;
                        returnList.Add(range);
                        startNode = null;
                        endNode = null;
                    }
                }
            }
            return returnList;

        }

        private static Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run)run.Clone(true);
            afterRun.Text = run.Text.Substring(position);

            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);

            return afterRun;
        }
    }
}
