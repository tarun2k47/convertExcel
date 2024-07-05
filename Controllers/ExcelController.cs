using convertExcel.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Net.Http;
using System.Text;

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
            #region Getting SessionID
           // string Error_Message = "";
            Data _Data = new Data();
            loginData _loginData = new loginData();
            _loginData.userName = "su";
            _loginData.password = "su";
            _loginData.CompanyId = "36";
            _Data.data.Add(_loginData);
            string Json1 = JsonConvert.SerializeObject(_Data);
            var client = new HttpClient();

            client.DefaultRequestHeaders.Add("fSessionId", "Session Id");
            string url = "http://localhost/focus8api/login";
            var content = new StringContent(Json1, Encoding.UTF8, "application/json");
            HttpResponseMessage response = client.PostAsync(url, content).GetAwaiter().GetResult();
            string responseBody = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            var responseObject = JsonConvert.DeserializeObject<Data>(responseBody);
            var sessionId = responseObject.data[0].fSessionId;
            #endregion
         //  List<string> Errors = new List<string>();
            FMYDateTime ft1 = new FMYDateTime();
            DataTable dt=new DataTable();
            try {  dt= ToDataTable(importFile);}
            catch(Exception e)
            {
                return Json(new { result = -1, message = e.Message+" Some Values or Space may be there please check Excel" });
            }
            #region Adding Vouchers
            HashData objHashRequest = new HashData();
            JArray jsonArray = new JArray();
            Dictionary<string, List<DataRow>> groupedRows = new Dictionary<string, List<DataRow>>();
            try
            {
                foreach (DataRow row in dt.Rows)
            {
                #region  Customer
                if (string.IsNullOrEmpty(row.ItemArray[3].ToString()))
                {
                    return Json(new { result = -1, message = "Customer Name is not provided or is empty" });
                }
                Account1 acc1 = new Account1();
                acc1.iAccountType = "5";
                acc1.sname = row.ItemArray[3].ToString();
                acc1.scode = row.ItemArray[3].ToString();
                hashdata1 h11 = new hashdata1();
                h11.data.Add(acc1);
                string Json111 = JsonConvert.SerializeObject(h11);
                var client11 = new HttpClient();
                client11.DefaultRequestHeaders.Add("fSessionId", sessionId);
                string url11 = "http://localhost/focus8api/Masters/Core__Account";
                var content11 = new StringContent(Json111, Encoding.UTF8, "application/json");
                client11.PostAsync(url11, content11).GetAwaiter().GetResult();
                #endregion
                #region Product
                if (string.IsNullOrEmpty(row.ItemArray[4].ToString()))
                {
                    return Json(new { result = -1, message = "Item is not provided or is empty" });
                }
                product p11 = new product();
                p11.sname = row.ItemArray[4].ToString();
                p11.scode = row.ItemArray[4].ToString();
                hash1 h12 = new hash1();
                h12.data.Add(p11);
                string Jsonp = JsonConvert.SerializeObject(h12);
                var client1p = new HttpClient();
                client1p.DefaultRequestHeaders.Add("fSessionId", sessionId);
                string urlp = "http://localhost/focus8api/Masters/Core__Product";
                var content1p = new StringContent(Jsonp, Encoding.UTF8, "application/json");
                client1p.PostAsync(urlp, content1p).GetAwaiter().GetResult();
                #endregion
                #region Delete
                var client1d = new HttpClient();
                client1d.DefaultRequestHeaders.Add("fSessionId", sessionId);
                string urld = "http://localhost/focus8api/Transactions/Sales%20Orders/" + row.ItemArray[0].ToString();
                HttpResponseMessage mess = client1p.DeleteAsync(urld).GetAwaiter().GetResult();
                string responseBody123 = mess.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                #endregion
                if (string.IsNullOrEmpty(row.ItemArray[0].ToString()))
                {
                    return Json(new { result = -1, message = "Voucher Number is not provided or is empty" });
                }
                string voucherNumber = row[0].ToString();
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
                    DateTime date;
                    string dateString = firstRow.ItemArray[1]?.ToString();
                    if (string.IsNullOrEmpty(dateString))
                    {
                        return Json(new { result = -1, message = "Date is not provided or is empty" });
                    }
                    if (!DateTime.TryParse(dateString, out date))
                    {
                        return Json(new { result = -1, message = "Date Format is not valid" });
                    }
                    DateTime startDate = new DateTime(2023, 1, 1);
                    DateTime endDate = new DateTime(2023, 10, 31);

                    if (date < startDate || date > endDate)
                    {
                        return Json(new { result = -1, message = "Date must be between 1/01/2023 and 10/31/2023" });
                    }
                    int dtint = new FMYDateTime().StringToIntDate(date.ToString("dd/MM/yyyy"));
                    objHeader.Add("DocNo", firstRow.ItemArray[0].ToString());
                    objHeader.Add("Date", dtint);
                    objHeader.Add("CustomerAC__Name", firstRow.ItemArray[3].ToString());
                    objHeader.Add("sNarration", firstRow.ItemArray[8].ToString());
                    foreach (DataRow row in rows)
                    {
                        JObject objBody = new JObject();
                        if (string.IsNullOrEmpty(row.ItemArray[5].ToString()))
                        {
                            return Json(new { result = -1, message = "Quantity Should not empty" });
                        }
                        if (string.IsNullOrEmpty(row.ItemArray[6].ToString()))
                        {
                            return Json(new { result = -1, message = "Rate Should not empty" });
                        }
                        if (double.Parse(row.ItemArray[5].ToString()) <= 0)
                        {
                            return Json(new { result = -1, message = "Quantity can't be Zero Or Negative" });
                        }
                        if (double.Parse(row.ItemArray[6].ToString()) < 0)
                        {
                            return Json(new { result = -1, message = "Rate Should not negative" });
                        }
                        if (double.Parse(row.ItemArray[7].ToString()) < 0)
                        {
                            return Json(new { result = -1, message = "Discount Should not negative" });
                        }
                        objBody.Add("Product__Code", row.ItemArray[4].ToString());
                        objBody.Add("Quantity", double.Parse(row.ItemArray[5].ToString()));
                        objBody.Add("Rate", double.Parse(row.ItemArray[6].ToString()));
                        double gross = Convert.ToDouble(row.ItemArray[5].ToString()) * Convert.ToDouble(row.ItemArray[6].ToString());
                        objBody.Add("Gross", gross);
                        objBody.Add("Discount", double.Parse(row.ItemArray[7].ToString()));
                        objBody.Add("sRemarks", row.ItemArray[9].ToString());
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
            }
            catch(Exception e)
            {
                return Json(new { result = -1, message = e.Message });
            }
            string json = JsonConvert.SerializeObject(jsonArray);
            List<Hashtable> dataList = JsonConvert.DeserializeObject<List<Hashtable>>(json);
            objHashRequest.data = dataList;
            string serializedObjHashRequest = JsonConvert.SerializeObject(objHashRequest);
            try
            {
                using (var client1 = new WebClient())
                {
                    client1.Headers.Add("fSessionId", sessionId);
                    client1.Headers.Add(HttpRequestHeader.ContentType, "application/json");
                    string sUrl = "http://localhost/focus8api/Transactions/Vouchers/Sales Orders";
                    string strResponse = client1.UploadString(sUrl, serializedObjHashRequest);
                    HashData objHashResponse = JsonConvert.DeserializeObject<HashData>(strResponse);
                    string messageFromResponse = objHashResponse.message;
                    int res = objHashResponse.result;
                    return Json(new { result = res, message = messageFromResponse });
                }
            }
            catch (WebException ex)
            {
                ViewBag.ErrorMessage = "Error occurred: " + ex.Message;
                return Json(new { status = -1, message = "An Error Occured" });
            }
            #endregion 
        }

        #region Excel Conversion into DataTable
        private DataTable ToDataTable(HttpPostedFileBase importFile)
        {
            string fileName = Path.GetFileName(importFile.FileName);
            string filePath = Path.Combine(Server.MapPath("~/Files"), fileName);
            importFile.SaveAs(filePath);
            DataTable dt = new DataTable();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.End.Row;
                int colCount = worksheet.Dimension.Columns;

                for (int col = 1; col <= colCount; col++)
                {
                    string columnName = worksheet.Cells[1, col].Value?.ToString();
                    if (!string.IsNullOrEmpty(columnName))
                        dt.Columns.Add(columnName);
                }
                for (int row = 2; row <= rowCount; row++)
                {
                    DataRow dataRow = dt.NewRow();
                    for (int col = 1; col <= colCount; col++)
                    {
                        dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString();
                    }
                    dt.Rows.Add(dataRow);
                }
            }
            return dt;
        }
        #endregion
    }
}