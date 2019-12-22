using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;
using Microsoft.Office.Interop.Word;
namespace TestTaskForLM.Controllers
{
    public class UploadFileController : Controller
    {
        // GET: UploadFile
        public ActionResult Index()
        {
            return View();
        }
        public static string[] properties = {"Last Print Date", "Number of Words","Number of Characters","Security",  "Number of Pages",
            "Total Editing Time","Application Name","Comments","Author",  "Last Save Time", "Keywords","Subject","Template",
            "Title", "Creation Date", "Revision Number", "Last Author", "Company" };  
        [HttpPost]
        public ActionResult Process(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0) 
            {
                string filename   = Path.GetFileName(file.FileName);
                byte[] bytesfile = new byte[10]; // 10 bytes is enought for determination
                file.InputStream.Read(bytesfile, 0, 10);
                string SaveLocation = Path.Combine(Server.MapPath("~/Data"), filename);
                Document uploadedFile = new Document(DeterminationFormatFile.GetFormatFile(bytesfile, SaveLocation), SaveLocation);
                if (uploadedFile.Doc != null && (uploadedFile.exe == ".docx" || uploadedFile.exe== ".doc"))
                {                    
                    try
                    {
                        file.SaveAs(SaveLocation);
                        Response.Write("The file " + filename + " has been uploaded. ");
                        ViewBag.Info = uploadedFile.processing();
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Info = ex.ToString();
                        //Note: Exception.Message returns a detailed message that describes the current exception. 
                        //For security reasons, we do not recommend that you return Exception.Message to end users in 
                        //production environments. It would be better to put a generic error message. 
                    }
                }
                else
                {
                    ViewBag.Info = "only word";
                }             
            }
            return View("Index");
        }
    }  
}