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
    public class FieldPlaceholderAdder : IReplacingCallback
    {
        private int _mergeReplacePosition;
        private IEnumerable<XElement> _inserts;
        private Document _doc;

        public FieldPlaceholderAdder(Document doc)
        {
            _doc = doc;
        }

        public void ReplacePlaceholderWithInserts(Node startingNode, string fieldName)
        {
            startingNode.Range.Replace(new Regex(fieldName), this, true);
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            args.Replacement = "<<field>>";
            return ReplaceAction.Replace;
        }
    }
}
