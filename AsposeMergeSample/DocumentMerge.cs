using Aspose.Words;
using Aspose.Words.Tables;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace AsposeMergeSample
{
    public class DocumentMerge
    {

        private string _targetDocumentPath;
        private XDocument _parameters;
        public int maxIfLevel;

        internal Document CompiledDoc
        {
            get;
            private set;
        }

        /// <summary>
        /// Accept the document & parameters.  Copies the orginal to the target path
        /// </summary>
        /// <param name="compiledDocPath">Path of compiled document</param>
        /// <param name="parametersPath">Path of parameters xml</param>
        /// <param name="targetDocPath">Path of destination document to work with</param>
        public void Initialise(string compiledDocPath,
                               string parametersPath,
                               string targetDocPath)
        {
            CompiledDoc = new Document(compiledDocPath);

            using (Stream str = new FileStream(parametersPath, FileMode.Open, FileAccess.Read))
            {
                _parameters = XDocument.Load(str);
            }

            CompiledDoc.Save(targetDocPath);
            CompiledDoc = new Document(targetDocPath);

            _targetDocumentPath = targetDocPath;
        }

        public void MergeInsertionsInRange()
        {
            ProcessIterativeMarkup();
            MergeFields();
            RemoveExcludedContent();

            CompiledDoc.Save(_targetDocumentPath);
        }


        internal void VerifyParameterNodes(Node startNode, Node endNode)
        {
            // The order in which these checks are done is important.
            if (startNode == null)
                throw new ArgumentException("Start node cannot be null");
            if (endNode == null)
                throw new ArgumentException("End node cannot be null");

            if (!startNode.Document.Equals(endNode.Document))
                throw new ArgumentException("Start node and end node must belong to the same document");

            if (startNode.GetAncestor(NodeType.Body) == null || endNode.GetAncestor(NodeType.Body) == null)
                throw new ArgumentException("Start node and end node must be a child or descendant of a body");

            // Check the end node is after the start node in the DOM tree
            // First check if they are in different sections, then if they're not check their position in the body of the same section they are in.
            Section startSection = (Section)startNode.GetAncestor(NodeType.Section);
            Section endSection = (Section)endNode.GetAncestor(NodeType.Section);

            int startIndex = startSection.ParentNode.IndexOf(startSection);
            int endIndex = endSection.ParentNode.IndexOf(endSection);

            if (startIndex == endIndex)
            {
                if (startSection.Body.IndexOf(startNode) > endSection.Body.IndexOf(endNode))
                    throw new ArgumentException("The end node must be after the start node in the body");
            }
            else if (startIndex > endIndex)
                throw new ArgumentException("The section of end node must be after the section start node");
        }

        internal bool IsInline(Node node)
        {
            // Test if the node is desendant of a Paragraph or Table node and also is not a paragraph or a table a paragraph inside a comment class which is decesant of a pararaph is possible.
            return ((node.GetAncestor(NodeType.Paragraph) != null || node.GetAncestor(NodeType.Table) != null) && !(node.NodeType == NodeType.Paragraph || node.NodeType == NodeType.Table));
        }

        internal void ProcessMarker(CompositeNode cloneNode, ArrayList nodes, Node node, bool isInclusive, bool isStartMarker, bool isEndMarker)
        {
            // If we are dealing with a block level node just see if it should be included and add it to the list.
            if (!IsInline(node))
            {
                // Don't add the node twice if the markers are the same node
                if (!(isStartMarker && isEndMarker))
                {
                    if (isInclusive)
                        nodes.Add(cloneNode);
                }
                return;
            }

            // If a marker is a FieldStart node check if it's to be included or not.
            // We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
            if (node.NodeType == NodeType.FieldStart)
            {
                // If the marker is a start node and is not be included then skip to the end of the field.
                // If the marker is an end node and it is to be included then move to the end field so the field will not be removed.
                if ((isStartMarker && !isInclusive) || (!isStartMarker && isInclusive))
                {
                    while (node.NextSibling != null && node.NodeType != NodeType.FieldEnd)
                        node = node.NextSibling;

                }
            }

            // If either marker is part of a comment then to include the comment itself we need to move the pointer forward to the Comment
            // Node found after the CommentRangeEnd node.
            if (node.NodeType == NodeType.CommentRangeEnd)
            {
                while (node.NextSibling != null && node.NodeType != NodeType.Comment)
                    node = node.NextSibling;

            }

            // Find the corresponding node in our cloned node by index and return it.
            // If the start and end node are the same some child nodes might already have been removed. Subtract the
            // Difference to get the right index.
            int indexDiff = node.ParentNode.ChildNodes.Count - cloneNode.ChildNodes.Count;

            // Child node count identical.
            if (indexDiff == 0)
                node = cloneNode.ChildNodes[node.ParentNode.IndexOf(node)];
            else
                node = cloneNode.ChildNodes[node.ParentNode.IndexOf(node) - indexDiff];

            // Remove the nodes up to/from the marker.
            bool isSkip = false;
            bool isProcessing = true;
            bool isRemoving = isStartMarker;
            Node nextNode = cloneNode.FirstChild;

            while (isProcessing && nextNode != null)
            {
                Node currentNode = nextNode;
                isSkip = false;

                if (currentNode.Equals(node))
                {
                    if (isStartMarker)
                    {
                        isProcessing = false;
                        if (isInclusive)
                            isRemoving = false;
                    }
                    else
                    {
                        isRemoving = true;
                        if (isInclusive)
                            isSkip = true;
                    }
                }

                nextNode = nextNode.NextSibling;
                if (isRemoving && !isSkip)
                    currentNode.Remove();
            }

            // After processing the composite node may become empty. If it has don't include it.
            if (!(isStartMarker && isEndMarker))
            {
                if (cloneNode.HasChildNodes)
                    nodes.Add(cloneNode);
            }


        }

        internal ArrayList ExtractContent(Node startNode, Node endNode, bool isInclusive)
        {
            // First check that the nodes passed to this method are valid for use.
            VerifyParameterNodes(startNode, endNode);

            // Create a list to store the extracted nodes.
            ArrayList nodes = new ArrayList();

            // Keep a record of the original nodes passed to this method so we can split marker nodes if needed.
            Node originalStartNode = startNode;
            Node originalEndNode = endNode;

            // Extract content based on block level nodes (paragraphs and tables). Traverse through parent nodes to find them.
            // We will split the content of first and last nodes depending if the marker nodes are inline
            while (startNode.ParentNode.NodeType != NodeType.Body)
                startNode = startNode.ParentNode;

            while (endNode.ParentNode.NodeType != NodeType.Body)
                endNode = endNode.ParentNode;

            bool isExtracting = true;
            bool isStartingNode = true;
            bool isEndingNode = false;
            // The current node we are extracting from the document.
            Node currNode = startNode;

            // Begin extracting content. Process all block level nodes and specifically split the first and last nodes when needed so paragraph formatting is retained.
            // Method is little more complex than a regular extractor as we need to factor in extracting using inline nodes, fields, bookmarks etc as to make it really useful.
            while (isExtracting)
            {
                // Clone the current node and its children to obtain a copy.
                CompositeNode cloneNode = (CompositeNode)currNode.Clone(true);
                isEndingNode = currNode.Equals(endNode);

                if (isStartingNode || isEndingNode)
                {
                    // We need to process each marker separately so pass it off to a separate method instead.
                    if (isStartingNode)
                    {
                        ProcessMarker(cloneNode, nodes, originalStartNode, isInclusive, isStartingNode, isEndingNode);
                        isStartingNode = false;
                    }

                    // Conditional needs to be separate as the block level start and end markers maybe the same node.
                    if (isEndingNode)
                    {
                        ProcessMarker(cloneNode, nodes, originalEndNode, isInclusive, isStartingNode, isEndingNode);
                        isExtracting = false;
                    }
                }
                else
                    // Node is not a start or end marker, simply add the copy to the list.
                    nodes.Add(cloneNode);

                // Move to the next node and extract it. If next node is null that means the rest of the content is found in a different section.
                if (currNode.NextSibling == null && isExtracting)
                {
                    // Move to the next section.
                    Section nextSection = (Section)currNode.GetAncestor(NodeType.Section).NextSibling;
                    currNode = nextSection.Body.FirstChild;
                }
                else
                {
                    // Move to the next node in the body.
                    currNode = currNode.NextSibling;
                }
            }

            // Return the nodes between the node markers.
            return nodes;
        }


        internal void ProcessIterativeMarkup()
        {
            StringBuilder repeatText = new StringBuilder();
            ContentFinder finder = new ContentFinder(CompiledDoc);
            if (DocumentElement.Element("ITERATIONS") != null)
            {
                var totalIterations = Convert.ToInt32(DocumentElement.Element("ITERATIONS").Descendants().Count());
                List<int> iterationCounts = new List<int>();
                for (int j = 0; j < totalIterations; j++)
                {
                    iterationCounts.Add(Convert.ToInt32(DocumentElement.Element("ITERATIONS").Elements("ITERATION").ToArray<XElement>()[j].Value));
                }

                for (int j = 0; j < totalIterations; j++)
                {
                    string startTerm = string.Empty;
                    string endTerm = string.Empty;
                    if (j == 0)
                    {
                        startTerm = CompiledMarkupConstants._FOREACH_BEGIN;
                        endTerm = CompiledMarkupConstants._FOREACH_END;
                    }
                    else
                    {
                        startTerm = string.Format("{0}-{1}", CompiledMarkupConstants._FOREACH_BEGIN, j);
                        endTerm = string.Format("{0}{1}>>", CompiledMarkupConstants._FOREACH_END_MULTIPLE, j);
                    }
                    var foundForEachRangeContent = finder.FindMatchingNodes(startTerm,
                                                                        endTerm);

                    if (foundForEachRangeContent.Count > 0)
                    {
                        foreach (NodeRangeMatch match in foundForEachRangeContent)
                        {
                            // Process the Iterations in *match* order
                            var iterationCount = iterationCounts.First();
                            iterationCounts.RemoveAt(0);
                            bool beginTagInTable = false;

                            Node begin = match.StartNode;
                            Node end = match.EndNode;
                            ReplaceMarkupWithField(begin, startTerm);
                            ReplaceMarkupWithField(end, endTerm);

                            // Copy the nodes *including the start term that is now a dummy field tag* to match
                            // the existing expansion
                            List<Node> extractedNodes = ContentExtractor.ExtractContent(begin, end, true, false, true);
                            Node insertAfter = begin.GetAncestor(NodeType.Paragraph);

                            beginTagInTable = IsTagInTable(begin);

                            // Expand the Foreach section for the number of iterations
                            extractedNodes.Reverse();
                            for (var i = 0; i < iterationCount - 1; i++)
                            {
                                insertAfter = AddRowForStartIteration(begin, insertAfter);
                                foreach (Node insertNode in extractedNodes)
                                {
                                    if (insertNode.NodeType == NodeType.Table && beginTagInTable)
                                    {
                                        // Insert the *contents* of the table rather than the table as a whole
                                        InsertTableContents(insertNode, insertAfter);
                                    }
                                    else
                                    {
                                        insertAfter.ParentNode.InsertAfter(insertNode.Clone(true), insertAfter);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private bool IsTagInTable(Node node)
        {
            bool returnVal = false;

            Paragraph parent = node.ParentNode as Paragraph;
            if (parent != null && parent.IsInCell)
            {
                returnVal = true;
            }
            return returnVal;
        }

        private void InsertTableContents(Node insertNode, Node insertAfter)
        {
            Table insertTable = insertNode as Table;
            Paragraph para = insertAfter as Paragraph;
            bool notfirstCellInTable = false;
            foreach (Node nodeToAdd in insertTable.GetChildNodes(NodeType.Any, true))
            {
                switch (nodeToAdd.NodeType)
                {
                    case NodeType.Cell:
                        {
                            // Insert a cell and change the insertAfter to a new Paragraph in this cell
                            // But only if not the first cell
                            if (notfirstCellInTable)
                            {
                                Cell currentCell = para.ParentNode as Cell;
                                if (currentCell != null)
                                {
                                    Cell newCell = new Cell(currentCell.Document);
                                    Paragraph newPara = new Paragraph(currentCell.Document);
                                    newCell.AppendChild(newPara);
                                    currentCell.ParentRow.AppendChild(newCell);
                                    insertAfter = newPara;
                                    para = newPara;
                                }
                            }
                            else
                            {
                                notfirstCellInTable = true;
                            }
                            break;
                        }
                    case NodeType.Row:
                        {
                            break;
                        }
                    case NodeType.Run:
                        {
                            break;
                        }
                    default:
                        {
                            insertAfter.ParentNode.InsertAfter(nodeToAdd.Clone(true), insertAfter);
                            break;
                        }
                }
            }
        }

        private Node AddRowForStartIteration(Node beginNode, Node insertAfter)
        {
            Node returnNode = insertAfter;
            Paragraph parentParagraph = beginNode.ParentNode as Paragraph;
            if (parentParagraph != null)
            {
                if (parentParagraph.IsInCell)
                {
                    Cell parentCell = parentParagraph.ParentNode as Cell;
                    if (parentCell != null)
                    {
                        if (parentCell.IsFirstCell)
                        {
                            // Add a new row per iteration
                            //Row newRow = (Row)parentCell.ParentRow.Clone(true);
                            //foreach(Cell childCell in newRow.Cells)
                            //{
                            //    childCell.RemoveAllChildren();
                            //}

                            //parentCell.ParentRow.ParentTable.AppendChild(newRow);

                            Row newRow = new Row(CompiledDoc);
                            Cell newCell = new Cell(CompiledDoc);
                            newRow.AppendChild(newCell);
                            parentCell.ParentRow.ParentTable.AppendChild(newRow);
                            Paragraph para = new Paragraph(CompiledDoc);
                            newRow.FirstCell.AppendChild(para);
                            returnNode = para;
                        }
                    }
                }
            }
            return returnNode;
        }

        internal XElement DocumentElement
        {
            get
            {
                return _parameters.Element("VFILE_DATA").Element("DOCUMENT");
            }
        }

        internal void RemoveMarkup(Node node)
        {
            var parent = node.ParentNode;
            node.Remove();
            if (parent.ToString(SaveFormat.Text).Trim() == string.Empty)
            {
                parent.Remove();
            }
        }

        internal void ReplaceMarkupWithField(Node node, string term)
        {
            FieldPlaceholderAdder fieldAdder = new FieldPlaceholderAdder(CompiledDoc);

            fieldAdder.ReplacePlaceholderWithInserts(node, term);

        }

        internal void ReplaceMarkupWithFieldGreedy(Node node, string term)
        {
            FieldPlaceholderAdder fieldAdder = new FieldPlaceholderAdder(CompiledDoc);

            term = string.Format(@"{0}(\S*)", term);
            fieldAdder.ReplacePlaceholderWithInserts(node, term);
        }

        internal void MergeFields()
        {
            var inserts = GetInserts();

            FieldPlaceholderMerger fieldMerge = new FieldPlaceholderMerger(CompiledDoc);
            fieldMerge.ReplacePlaceholderWithInserts(inserts);

            // Store the Maximum level of 'if' tag to replace
            maxIfLevel = fieldMerge.maxIf;

        }


        internal void RemoveExcludedContent()
        {

            ContentFinder finder = new ContentFinder(CompiledDoc);

            // Replace the If levels in reverse order (i.e. most indented to least indented )
            for (int level = maxIfLevel; level > 0; level--)
            {
                var foundText = finder.FindExclusionTags(level);
                // Remove Table formatting in Order (Cells / Then Rows then Tables )

                List<Node> itemsToRemove = new List<Node>();
                var cells = from Node node in foundText
                            where node.NodeType == NodeType.Cell
                            select node;
                foreach (Cell cell in cells)
                {
                    RemoveChildNode(cell);
                    itemsToRemove.Add(cell);
                }

                var rows = from Node node in foundText
                           where node.NodeType == NodeType.Row
                           select node;
                foreach (Row row in rows)
                {
                    RemoveChildNode(row);
                    itemsToRemove.Add(row);
                }

                var tables = from Node node in foundText
                             where node.NodeType == NodeType.Table
                             select node;
                foreach (Table table in tables)
                {
                    RemoveChildNode(table);
                    itemsToRemove.Add(table);
                }

                foreach (Node removeNode in itemsToRemove)
                {
                    foundText.Remove(removeNode);
                }

                // Runs in the sequence.
                foreach (Node run in foundText)
                {
                    RemoveChildNode(run);
                }

            }

        }

        private void RemoveChildNode(Node nodeToRemove)
        {
            Node para = nodeToRemove.ParentNode;
            if (para != null)
            {
                nodeToRemove.Remove();

                if (para.ToString(SaveFormat.Text).Trim() == string.Empty)
                {
                    if (para.ParentNode != null)
                    {
                        para.Remove();
                    }
                }
            }

        }


        internal IEnumerable<XElement> GetInserts()
        {
            var insertQuery = from i in _parameters.Descendants(SchemaConstants._INSERTS_ELEMENT)
                              select i;

            return insertQuery.FirstOrDefault().Elements();
        }

    }
}
