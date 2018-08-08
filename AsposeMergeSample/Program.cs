using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsposeMergeSample
{
    class Program
    {
        static void Main(string[] args)
        {
            Aspose.Words.License license = new Aspose.Words.License();
            license.SetLicense(@"c:\temp\Aspose.Total.lic");

            // Problem One - inconsistent finding of <<&foreach tags based on position in document range
            // This seems to be entirely dependent on how the document was edited and how the runs / paragraphs fall 
            // as to whether it finds the tags or not


            // Working document 
            DocumentMerge merge = new DocumentMerge();
            merge.Initialise(@"Documents\WorkingExpansion\WorkingExpansion.docx", @"Documents\WorkingExpansion\InputData.xml",@"Documents\WorkingExpansion\OutputDocument.docx");            
           
            // Expand any ForEach blocks in the Document
            merge.ProcessIterativeMarkup();
            merge.CompiledDoc.Save(@"Documents\WorkingExpansion\ExpansionComplete.docx");

            // None working document - only difference is the position of the opening foreach tag
            // but *both* tags are now not expanded

            merge = new DocumentMerge();
            merge.Initialise(@"Documents\NoneWorkingExpansion\NoneWorkingExpansion.docx", @"Documents\NoneWorkingExpansion\InputData.xml", @"Documents\NoneWorkingExpansion\OutputDocument.docx");

            // Expand any ForEach blocks in the Document
            merge.ProcessIterativeMarkup();
            merge.CompiledDoc.Save(@"Documents\NoneWorkingExpansion\ExpansionComplete.docx");


            // Problem Two - Removing Content between two fields 

            merge = new DocumentMerge();
            merge.Initialise(@"Documents\RemoveExcludedContent\InputDocument.docx", @"Documents\RemoveExcludedContent\InputData.xml", @"Documents\RemoveExcludedContent\OutputDocument.docx");
            // Replace the Field markers with the XML data, creating unique tag id's
            merge.MergeFields();
            merge.CompiledDoc.Save(@"Documents\RemoveExcludedContent\MergeComplete.docx");

            // Remove the sections that should be removed
            merge.RemoveExcludedContent();
            merge.CompiledDoc.Save(@"Documents\RemoveExcludedContent\RemovalComplete.docx");

            // Note that the Second section has gone wrong - the Bold "s" is missing 
        }
    }
}
