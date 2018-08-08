using Aspose.Words;
using Aspose.Words.Replacing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsposeMergeSample
{
    /// <summary>
    /// Replace does not work across special characters!!! IE \r 
    /// http://www.aspose.com/community/forums/thread/749201/document.range.replace-is-throwing-exception-when-document-is-changed.aspx
    /// </summary>
    public class EndRangeTextFinder : IReplacingCallback
    {
        public List<Node> Nodes { get; private set; } = new List<Node>();

        /// <summary>
        /// Represents the End Tag we are looking for to allow for the limiting with the match
        /// </summary>
        public string EndRange { get; set; }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // This is a Run node that contains either the beginning or the complete match.
            Node currentNode = args.MatchNode;

            // The first (and may be the only) run can contain text before the match, 
            // in this case it is necessary to split the run.
            if (args.MatchOffset > 0)
                currentNode = SplitRun((Run)currentNode, args.MatchOffset);

            int remainingLength = args.Match.Value.IndexOf(EndRange) + EndRange.Length;

            while (
                 (remainingLength > 0) &&
                 (currentNode != null) &&
                 (currentNode.GetText().Length <= remainingLength))
            {
                Nodes.Add(currentNode);
                // If we have the end tag within this run, go up to the end tag and no further
                if (currentNode.GetText().IndexOf(EndRange) >= 0)
                {
                    remainingLength = currentNode.GetText().IndexOf(EndRange);
                }
                else
                {
                    remainingLength = remainingLength - currentNode.GetText().Length;
                }

                // Select the next Run node. 
                // Have to loop because there could be other nodes such as BookmarkStart etc.
                do
                {
                    currentNode = currentNode.NextPreOrder(currentNode.Document);
                    if (currentNode != null)
                    {
                        // Return any 'formatting nodes (Table / Row / Cell so they can be removed as well )
                        if (currentNode.NodeType != NodeType.Run)
                        {
                            if (currentNode.NodeType != NodeType.Paragraph)
                            {
                                Nodes.Add(currentNode);
                            }
                            switch (currentNode.NodeType)
                            {
                                case NodeType.Table:
                                    {
                                        remainingLength--;
                                        break;
                                    }
                                case NodeType.Row:
                                    {
                                        remainingLength--;
                                        break;
                                    }
                            }
                            //remainingLength--;
                        }
                        //if (currentNode != null && string.IsNullOrEmpty(currentNode.ToString(SaveFormat.Text).Trim()))
                        //    Nodes.Add(currentNode);

                    }
                }
                while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
            }

            // Split the last run that contains the match if there is any text left.
            if (currentNode != null && currentNode is Run && remainingLength > 0)
            {
                SplitRun((Run)currentNode, remainingLength);
                Nodes.Add(currentNode);
            }

            // Signal to the replace engine to do nothing because we have already done all what we wanted.
            return ReplaceAction.Skip;
        }

        /// <summary>
        /// Splits text of the specified run into two runs.
        /// Inserts the new run just after the specified run.
        /// </summary>
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
