using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace TestTaskForLM
{
    public class OfficeDocumentDoc: OfficeDocument
    {
        public static string[] properties = {"Last Print Date", "Number of Words","Number of Characters","Security",  "Number of Pages",
            "Total Editing Time","Application Name","Comments","Author",  "Last Save Time", "Keywords","Subject","Template",
            "Title", "Creation Date", "Revision Number", "Last Author", "Company" };

        // я перерыл много разных источников, но смог найти ток такой способ
        // оно работает медленно, нужно подождать
        // и кстати, это тоже работает с новым форматом
        public override string processing(string fileName)
        {
             //create word app class object                                                                                                                         //           object file = FILENAME;                                               //this is the path to file to open
            string summaryInformation = "This is old format  \n";
            var nullobject = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordObject = null;
            Microsoft.Office.Interop.Word.Document docs=null;
            try
            {
                wordObject = new Microsoft.Office.Interop.Word.Application();
                docs = wordObject.Documents.Open(fileName, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject,
                 nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject);           
            }
            catch (Exception ex)
            {
                wordObject.Quit();
                return summaryInformation + " but, couldn't open and procces this file";
            }
            var wordProperties = docs.BuiltInDocumentProperties;
            Type typeDocBuiltInProps = wordProperties.GetType();
                for (int i = 0; i < properties.Length; i++)
                {
                    try
                    {
                        Object Title = typeDocBuiltInProps.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty
                       , null, wordProperties, new object[] { properties[i] });
                        Type typeTitleprop = Title.GetType();
                        string strTitleprop = typeTitleprop.InvokeMember("Value", BindingFlags.Default | BindingFlags.GetProperty,
                            null, Title, new object[] { }).ToString();
                        summaryInformation += properties[i] + ":  " + strTitleprop+";  ";
                    }
                    catch (Exception j)
                    {
                        summaryInformation += j.Message;
                    }
                }                               
                docs.Close(WdSaveOptions.wdDoNotSaveChanges, nullobject, nullobject);
                ((_Application)wordObject).Quit(WdSaveOptions.wdDoNotSaveChanges);                        
            return summaryInformation;
        }
    }
}
