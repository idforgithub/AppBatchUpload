using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using OfficeOpenXml;
using System.Text;

namespace AppBatchUpload.Controllers
{
    public class ImportExportController : Controller
    {

        private readonly IHostingEnvironment _iHostingEnviroment;

        public ImportExportController(IHostingEnvironment iHostingEnviroment)
        {
            this._iHostingEnviroment = iHostingEnviroment;
        }

        public RedirectResult Import()
        {
            string sWebRootFolder = _iHostingEnviroment.WebRootPath;
            string sFileName = @"Template.xlsx";
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            try
            {
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    StringBuilder sb = new StringBuilder();
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;
                    bool bHeaderRow = true;
                    for (int r = 0; r <= rowCount; r++)
                    {
                        for (int c = 0; c <= rowCount; c++)
                        {
                            if (bHeaderRow)
                            {
                                sb.Append(worksheet.Cells[r, c].Value.ToString() + "\t");
                            }
                            else
                            {
                                sb.Append(worksheet.Cells[r, c].Value.ToString() + "\t");
                            }
                            sb.Append(Environment.NewLine);
                        }
                    }
                    return Redirect(sb.ToString());
                }
            }
            catch (Exception e)
            {
                return Redirect(e.Message);
            }
        }

        public RedirectResult Export()
        {
            string sWebRootFolder = _iHostingEnviroment.WebRootPath;
            string sFileName = @"template.xlsx";
            string URL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, sFileName);
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            }
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("TEST");

                //Add in Headers
                worksheet.Cells[1, 1].Value = "Name";
                worksheet.Cells[1, 2].Value = "Role";


                //Add in values
                worksheet.Cells["A2"].Value = "EXAMPLE NAME";
                worksheet.Cells["A3"].Value = "EXAMPLE NAME";
                worksheet.Cells["A4"].Value = "EXAMPLE NAME";

                worksheet.Cells["B2"].Value = "EXAMPLE ROLE NAME";
                worksheet.Cells["B3"].Value = "EXAMPLE ROLE NAME";
                worksheet.Cells["B4"].Value = "EXAMPLE ROLE NAME";
                package.Save();
            }
            return Redirect(URL);
        }
    }
}