using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsposeMergeSample
{
    /// <summary>
    /// Class to contain a node range match between 2 'tags' in the document
    /// 
    /// for example the start and end nodes for a ForEach in the document
    /// there can be multiple matches within a document so will 
    /// 
    /// 
    /// </summary>
    public class NodeRangeMatch
    {
        /// <summary>
        /// Start node of the Match
        /// </summary>
        public Node StartNode;
        /// <summary>
        /// End node of the Match
        /// </summary>
        public Node EndNode;
        /// <summary>
        /// Represents the Nesting Level of the match (i.e. is this match within another match)
        /// </summary>
        public int Level = 0;
    }
}
