using convertExcel.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using Focus.Common.DataStructs;


namespace convertExcel.Controllers
{
    public class ExcelController : Controller
    {
        // GET: Excel
        public ActionResult Index()
        {
            return View("Index");
        }
        [HttpPost]
        public ActionResult SubmitOrder(HttpPostedFileBase importFile)
        {
            FMYDateTime ft1 = new FMYDateTime();
            try
            {
                var dt = ToDataTable(importFile);
                HashData objHashRequest = new HashData();
                JArray jsonArray = new JArray();

                Dictionary<string, List<DataRow>> groupedRows = new Dictionary<string, List<DataRow>>();

                foreach (DataRow row in dt.Rows)
                {
                    string voucherNumber = row[0].ToString(); // Assuming voucher number is in the first column

                    if (!groupedRows.ContainsKey(voucherNumber))
                    {
                        groupedRows[voucherNumber] = new List<DataRow>();
                    }

                    groupedRows[voucherNumber].Add(row);
                }


                foreach (var group in groupedRows)
                {
                    string voucherNumber = group.Key;
                    List<DataRow> rows = group.Value;


                    JArray bodyArray = new JArray();


                    JObject objHeader = new JObject();


                    DataRow firstRow = rows[0];
                    DateTime date = DateTime.Parse(firstRow.ItemArray[1].ToString());
                    int dtint = new FMYDateTime().StringToIntDate(date.ToString("dd/MM/yyyy"));
                    objHeader.Add("Date", dtint);
                    objHeader.Add("CustomerAC__Code", null);
                    objHeader.Add("CustomerAC__Name", firstRow.ItemArray[3].ToString());
                    objHeader.Add("sNarration", firstRow.ItemArray[4].ToString());

                    foreach (DataRow row in rows)
                    {

                        JObject objBody = new JObject();
                        objBody.Add("Product__Code", row.ItemArray[5].ToString());
                        objBody.Add("Quantity", double.Parse(row.ItemArray[6].ToString()));
                        objBody.Add("Rate", double.Parse(row.ItemArray[7].ToString()));
                        int gross = Convert.ToInt32(row.ItemArray[6].ToString()) * Convert.ToInt32(row.ItemArray[7].ToString());
                        objBody.Add("Gross", gross);


                        bodyArray.Add(objBody);
                    }


                    if (rows.Count > 1)
                    {
                        JToken firstBody = bodyArray.First;
                        bodyArray.RemoveAt(0);
                        bodyArray.Add(firstBody);
                    }


                    JObject objHash = new JObject();
                    objHash.Add("Body", bodyArray);
                    objHash.Add("Header", objHeader);
                    objHash.Add("Footer", new JArray());


                    jsonArray.Add(objHash);
                }

                string json = JsonConvert.SerializeObject(jsonArray);
                List<Hashtable> dataList = JsonConvert.DeserializeObject<List<Hashtable>>(json);
                objHashRequest.data = dataList;
                string serializedObjHashRequest = JsonConvert.SerializeObject(objHashRequest);

                try
                {

                    using (var client = new WebClient())
                    {
                        client.Headers.Add("fSessionId", "14062024165516292361");
                        client.Headers.Add(HttpRequestHeader.ContentType, "application/json"); // Set content type to JSON
                        string sUrl = "http://localhost/focus8api/Transactions/Vouchers/Sales Orders";
                        string strResponse = client.UploadString(sUrl, serializedObjHashRequest);
                        HashData objHashResponse = JsonConvert.DeserializeObject<HashData>(strResponse);
                          
                      


                    }
                }
                catch (WebException ex)
                {
                    ViewBag.ErrorMessage = "Error occurred: " + ex.Message;
                    return View("Error");
                }


                return View();

            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = "Error occurred: " + ex.Message;
                return View("Error");
            }
        }


        private DataTable ToDataTable(HttpPostedFileBase importFile)
        {
            string fileName = Path.GetFileName(importFile.FileName);
            string filePath = Path.Combine(Server.MapPath("~/Files"), fileName);
            importFile.SaveAs(filePath);

            DataTable dt = new DataTable();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.End.Row;
                int colCount = worksheet.Dimension.Columns;

                int startRow = 1;
                for (int row = 1; row <= rowCount; row++)
                {
                    bool isEmptyRow = true;
                    for (int col = 1; col <= colCount; col++)
                    {
                        if (worksheet.Cells[row, col].Value != null)
                        {
                            isEmptyRow = false;
                            break;
                        }
                    }
                    if (!isEmptyRow)
                    {
                        startRow = row;
                        break;
                    }
                }

                int startCol = 1;
                for (int col = 1; col <= colCount; col++)
                {
                    bool isEmptyColumn = true;
                    for (int row = startRow; row <= rowCount; row++)
                    {
                        if (worksheet.Cells[row, col].Value != null)
                        {
                            isEmptyColumn = false;
                            break;
                        }
                    }
                    if (!isEmptyColumn)
                    {
                        startCol = col;
                        break;
                    }
                }

                for (int col = startCol; col <= colCount; col++)
                {
                    string columnName = worksheet.Cells[startRow, col].Value?.ToString();
                    if (!string.IsNullOrEmpty(columnName))
                        dt.Columns.Add(columnName);
                }

                for (int row = startRow + 1; row <= rowCount; row++)
                {
                    DataRow dataRow = dt.NewRow();
                    for (int col = startCol; col <= colCount; col++)
                    {
                        dataRow[col - startCol] = worksheet.Cells[row, col].Value?.ToString();
                    }
                    dt.Rows.Add(dataRow);
                }
            }

            return dt;
        }
        public ActionResult Success()
        {
            return View();
        }
    }
}