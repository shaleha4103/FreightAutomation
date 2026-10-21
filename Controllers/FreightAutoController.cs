using FreightAutomation.BAL;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace FreightAutomation.Controllers
{
    public class FreightAutoController : Controller
    {
        private readonly FreightBAL freightBAL;
        public FreightAutoController()
        {
            freightBAL = new FreightBAL();
        }


        [HttpGet]
        public ActionResult Upload()
        {
            return View();
        }

        // ==========================================
        // UPLOAD EXCEL
        // ==========================================

        [HttpPost]
        public JsonResult UploadExcel(
            HttpPostedFileBase excelfile)
        {
            // Check file
            if (excelfile == null ||
                excelfile.ContentLength == 0)
            {
                return Json(new
                {
                    success = false,
                    message = "Please select an Excel file."
                });
            }

            try
            {
                // ==========================================
                // UPLOAD EXCEL THROUGH BAL
                // ==========================================

                int uploadId =
                    freightBAL.UploadExcel(excelfile);

                // ==========================================
                // RETURN UPLOAD ID TO JAVASCRIPT
                // ==========================================

                return Json(new
                {
                    success = true,
                    message =
                        "Excel data successfully saved into database.",
                    uploadId = uploadId
                });
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    success = false,
                    message =
                        "Error: " + ex.Message
                });
            }
        }

        // ==========================================
        // DOWNLOAD EXCEL
        // ==========================================

        [HttpGet]
        public ActionResult DownloadExcel(int uploadId)
        {
            DataTable dt = new DataTable();

            string connectionString =
                ConfigurationManager
                .ConnectionStrings["CreateCon"]
                .ConnectionString;

            // ==========================================
            // GET RECORDS FOR THIS UPLOAD ID ONLY
            // ==========================================

            using (SqlConnection con =
                new SqlConnection(connectionString))
            {
                con.Open();

                string query = @"
                    SELECT *
                    FROM FreightCal
                    WHERE UploadId = @UploadId
                    ORDER BY Id";

                using (SqlCommand cmd =  new SqlCommand(query, con))
                {
                    cmd.Parameters.AddWithValue(
                        "@UploadId",
                        uploadId);

                    using (SqlDataAdapter da =
                        new SqlDataAdapter(cmd))
                    {
                        da.Fill(dt);
                    }
                }
            }

            // ==========================================
            // CHECK RECORDS
            // ==========================================

            if (dt.Rows.Count == 0)
            {
                return new HttpStatusCodeResult(
                    404,
                    "No records found for UploadId "
                    + uploadId);
            }

            // ==========================================
            // CREATE PLAIN EXCEL
            // ==========================================

            using (
                ClosedXML.Excel.XLWorkbook workbook =
                new ClosedXML.Excel.XLWorkbook())
            {
                ClosedXML.Excel.IXLWorksheet worksheet =
                    workbook.Worksheets.Add(
                        "FreightCalq");

                // ==========================================
                // WRITE COLUMN HEADERS
                // ==========================================

                for (int column = 0;
                     column < dt.Columns.Count;
                     column++)
                {
                    worksheet.Cell(
                        1,
                        column + 1
                    ).Value =
                        dt.Columns[column].ColumnName;
                }

                // ==========================================
                // WRITE DATA
                // ==========================================

                for (int row = 0;
                     row < dt.Rows.Count;
                     row++)
                {
                    for (int column = 0;
                         column < dt.Columns.Count;
                         column++)
                    {
                        object value =
                            dt.Rows[row][column];

                        if (value == DBNull.Value)
                        {
                            worksheet.Cell(
                                row + 2,
                                column + 1
                            ).Value = "";
                        }
                        else
                        {
                            worksheet.Cell(
                                row + 2,
                                column + 1
                            ).Value =
                                value.ToString();
                        }
                    }
                }

               

                using (MemoryStream stream =
                    new MemoryStream())
                {
                    workbook.SaveAs(stream);

                    return File(
                        stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Updated_Freight_"
                        + uploadId
                        + ".xlsx"
                    );
                }
            }
        }
    }
}