using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
//@Model IDictionaru<string, string>
namespace TestTaskForLM
{
    public class OfficeDocumentDocx : OfficeDocument
    {
        public override string processing(string fileNAme)
        {
            string summaryinformation="This is OpenXML format  ";
            using (WordprocessingDocument document = WordprocessingDocument.Open(fileNAme, false))
            {
                var props = document.PackageProperties;
                summaryinformation+=" Creator = " + props.Creator + ";  ";
                summaryinformation += " Created = " + props.Created + ";  "; 
                summaryinformation += " Title = " + props.Title + ";  ";
                summaryinformation += " Category = " + props.Category + ";  ";
                summaryinformation += " Description = " + props.Description + ";  ";

                summaryinformation += " Keywords = " + props.Keywords + ";  ";
                summaryinformation += " Language = " + props.Language + ";  ";
                summaryinformation += " LastPrinted = " + props.LastPrinted + ";  ";
                summaryinformation += " Modified = " + props.Modified + ";  ";
                summaryinformation += " Revision = " + props.Revision + ";  ";

                summaryinformation += " Version = " + props.Version + ";  ";
                summaryinformation += " LastModifiedBy = " + props.LastModifiedBy + ";  ";
                summaryinformation += " Identifier = " + props.Identifier + ";  ";
                summaryinformation += " ContentStatus = " + props.ContentStatus + ";  ";
                summaryinformation += " ContentStatus = " + props.Subject + ";  ";
            }
            return summaryinformation;
        }
    }
}