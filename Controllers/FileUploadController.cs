using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml; // EPPlus namespace
using uploadfile;
using Xceed.Document.NET;
using Xceed.Words.NET;


[Route("api/[controller]")]
[ApiController]
public class FileUploadController : ControllerBase
{
    [HttpPost("upload")]
    public async Task<IActionResult> UploadExcelFile(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest("No file uploaded.");
        }

        // Optional: Check if it's an Excel file
        if (!file.FileName.EndsWith(".xls") && !file.FileName.EndsWith(".xlsx"))
        {
            return BadRequest("Only Excel files are allowed.");
        }

        var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "UploadedFiles");
        if (!Directory.Exists(uploadsFolder))
        {
            Directory.CreateDirectory(uploadsFolder);
        }

        //var filePath = Path.Combine(uploadsFolder, file.FileName);
        //string OnlyFileName = Path.GetFileNameWithoutExtension(file.FileName);
        //var DocFilePath= Path.Combine(uploadsFolder, OnlyFileName+".docx");
        

        //using (var stream = new FileStream(filePath, FileMode.Create))
        //{
        //    await file.CopyToAsync(stream);
        //}

    //    try
    //    {
           
    //        CreateDocx createDocx = new CreateDocx();
    //        createDocx.CreateWordDocx(filePath, DocFilePath);
                

               
            
    //    }
    //    catch (Exception ex)
    //    {
    //        return StatusCode(500, "Excel processing failed: " + ex.Message);
    //    }

       return Ok("File uploaded and processed successfully.");
    }

       


}
