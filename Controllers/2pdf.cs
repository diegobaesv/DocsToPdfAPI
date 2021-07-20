
using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Text.Json;

namespace DocsToPdfAPI.Controllers
{
    
       [Route("api/2pdf")]
    [ApiController]
    public class _2pdf : ControllerBase {

        static String ROUTE_BASE = @"C:\API2PDF\";

        [HttpPost("")]
        public String Post(Base64dto base64)

        {

            try
            {
                createFolder(ROUTE_BASE);
                if (base64.extension == "doc" || base64.extension == "docx")
                {
                    return convertDocumentToPDF(base64); ;
                }
                else
                {
                    throw new NotSupportedException("EXTENSION NO SOPORTADA");
                }
            }
            catch(Exception e)
            {
                return e.ToString();
            }
           

        }

        private String convertDocumentToPDF(Base64dto base64) {

            String fullPathDoc = base64ToDisk(base64);
            /*Type wordType = Type.GetTypeFromProgID("Word.Application");
            dynamic msword = Activator.CreateInstance(wordType);*/

            Microsoft.Office.Interop.Word._Application word = new Microsoft.Office.Interop.Word.Application();
           
            //Application word = new Application();
            object oMissing = System.Reflection.Missing.Value;

            FileInfo wordFile = new FileInfo(fullPathDoc);

            word.Visible = false;
            word.ScreenUpdating = false;

            Object filename = (Object)wordFile.FullName;

            Document doc = word.Documents.Open(ref filename,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Activate();

            object outputFilename = wordFile.FullName.Replace(base64.extension,"pdf");
            object fileFormat = WdSaveFormat.wdFormatPDF;

            doc.SaveAs2(ref outputFilename,
                ref fileFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            object savechanges = WdSaveOptions.wdSaveChanges;
            ((_Document)doc).Close(ref savechanges, ref oMissing, ref oMissing);
            doc = null;

            ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;

            return JsonSerializer.Serialize(diskToBase64dto(outputFilename.ToString(), "pdf"));
        }

        private String base64ToDisk(Base64dto base64) {
           
            String subPath = Environment.ExpandEnvironmentVariables(ROUTE_BASE + getFolderRoute());
            createFolder(subPath);

            String fileName = getFileName() + "." + base64.extension;
            System.IO.File.WriteAllBytes(subPath + fileName, Convert.FromBase64String(base64.data));
            return subPath+fileName;
        }

        private Base64dto diskToBase64dto(String routeFile, String extension) {
            Byte[] bytes = System.IO.File.ReadAllBytes(routeFile);
            String data = Convert.ToBase64String(bytes);

            return new Base64dto(data, extension);
        }


        private void createFolder(String path)
        {
            bool exists = Directory.Exists(path);
            if (!exists)
               Directory.CreateDirectory(path);
        }

        private String getFolderRoute(){
            DateTime localDate = DateTime.Now;
            String month = localDate.Month < 10 ? "0"+localDate.Month.ToString(): localDate.Month.ToString();
            String day = localDate.Day < 10 ? "0" + localDate.Day.ToString() : localDate.Day.ToString();
            return localDate.Year.ToString()+month+day+@"\";
        }

        private String getFileName(){
            Random random = new Random();
            DateTime localDate = DateTime.Now;
            String hour = localDate.Hour < 10 ? "0" + localDate.Hour.ToString() : localDate.Hour.ToString();
            String minute = localDate.Minute < 10 ? "0" + localDate.Minute.ToString() : localDate.Minute.ToString();
            String second = localDate.Second < 10 ? "0" + localDate.Second.ToString() : localDate.Second.ToString();
            char letter = (char)('a' + random.Next(0, 26));
            return hour+minute+second+letter.ToString().ToUpper();
        }
        
    }
}

public class Base64dto
{
    public Base64dto(String dat, String ext)
    {
        extension = ext;
        data = dat;
    }

    private Base64dto() { }
    public String extension { get; set; }   
    public String data { get; set; }
}
