using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TestTaskForLM
{
    public class Document
    {
        public string exe;
        public string namefile;
        public OfficeDocument Doc { get; set; }
        
        public Document(OfficeDocument determinationFile, string name)
        {
            this.namefile = name;
            this.Doc = determinationFile;
            if (determinationFile is OfficeDocumentDoc)
                exe = ".doc";
            if (determinationFile is OfficeDocumentDocx)
                exe = ".docx";
        }
        public string processing()
        {
            return Doc.processing(namefile);
        }
    }
}