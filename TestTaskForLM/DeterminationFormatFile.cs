using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;

namespace TestTaskForLM
{
    public static class DeterminationFormatFile
    {
        private static readonly byte[] DOC = { 208, 207, 17, 224, 161, 177, 26, 225 };
        private static readonly byte[] DOCX = { 80, 75, 3, 4 };

        public static  OfficeDocument GetFormatFile(byte[] file, string fileName)
        {
            //Ensure that the filename isn't empty or null
            //Get the MIME Type
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return null;
            }
            if (file.Take(8).SequenceEqual(DOC) )
            {
                return new OfficeDocumentDoc();
            }
            else if (file.Take(4).SequenceEqual(DOCX))
            {
                return  new OfficeDocumentDocx();
            }else
            {
                if (Path.GetExtension(fileName) == ".doc")
                    return new OfficeDocumentDoc();
                if (Path.GetExtension(fileName) == ".docx")
                    return new OfficeDocumentDocx();
            }
            return null;
        }
    }
}