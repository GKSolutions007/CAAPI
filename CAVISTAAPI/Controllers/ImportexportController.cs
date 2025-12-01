using CAVISTAAPI.BuisnessLayer;
using CAVISTAAPI.Models;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Http;

namespace CAVISTAAPI.Controllers
{
    public class ImportexportController : ApiController
    {
        clsBusinessLayer objBL = new clsBusinessLayer();
        public string strSheetName { get; set; }
        public int ConstitutionID { get; set; }
        public int StateID { get; set; }
        public string strExtension = ".xlsx";
        public string strFileName = "";
        public string strFilePath
        {
            get; set;
        }
        public DataTable dtData { get; set; }
        public DataTable dtHeaderData { get; set; }
        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/uploadimportfile")]
        public IHttpActionResult LoadSelectFiledata()
        {
            string Msg = "";
            string dt = "";
            List<ImportResults> MTM = new List<ImportResults>();
            try
            {
                var file = HttpContext.Current.Request.Files.Count > 1 ? HttpContext.Current.Request.Files[0] : null;
                //var data = Request.Files[0].InputStream.Read;                                                       
                if (HttpContext.Current.Request.Files.Count > 0)
                {
                    string TransID = HttpContext.Current.Request.Files.AllKeys[0].ToString();
                    string TransName = HttpContext.Current.Request.Files.AllKeys[1].ToString();
                    string fileName = HttpContext.Current.Request.Files[2].FileName;
                    string fileContentType = HttpContext.Current.Request.Files[2].ContentType;
                    string UserID = HttpContext.Current.Request.Files.AllKeys[2].ToString();
                    //strFilePath = AppDomain.CurrentDomain.BaseDirectory + "Upload Files\\";
                    string FPt = System.Configuration.ConfigurationManager.AppSettings["SupportFilePath"];
                    strFilePath = FPt + "Upload Files\\";
                    strFileName = TransName + "_Upload_" + fileName;
                    if (!Directory.Exists(strFilePath))
                    {
                        Directory.CreateDirectory(strFilePath);
                    }
                    HttpContext.Current.Request.Files[3].SaveAs(strFilePath + strFileName);
                    bool blResult = true;
                    List<string> lst = null;
                    #region Validate Columns
                    if (TransID == "1")//Customer
                    {
                        lst = CustomerMasterTemp();
                    }                    
                    bool ErrorColAlreadyExisist = false;
                    ColumnValidation(lst, ref blResult);
                    if (!blResult)
                    {
                        if (TransID == "1")//Customer
                        {
                            lst = CustomerMasterTempWithErrCol();
                        }                        
                        ColumnValidation(lst, ref blResult);
                        ErrorColAlreadyExisist = true;
                    }
                    #endregion
                    if (blResult)
                    {
                        DataTable dtCorrectValues = new DataTable();
                        DataTable dtWrongValues = new DataTable();
                        foreach (string str in lst)
                        {
                            dtCorrectValues.Columns.Add(str);
                            dtWrongValues.Columns.Add(str);
                        }
                        dtCorrectValues.Columns.Add("TaxPern");
                        dtCorrectValues.Columns.Add("UOM");
                        if (!ErrorColAlreadyExisist)
                        {
                            dtCorrectValues.Columns.Add("Error");
                            dtWrongValues.Columns.Add("Error");
                        }
                        #region Customer
                        else if (TransID == "1")
                        {
                            if (dtData.Rows.Count > 0)
                            {
                                bool NoErrors = true;
                                foreach (DataRow item in dtData.Rows)
                                {
                                    DataTable dtValidate = dtData.Clone();
                                    dtValidate.TableName = "Validation";
                                    dtValidate.Rows.Add(item.ItemArray);
                                    string RowError = CustomerImpValiation(dtValidate);
                                    if (string.IsNullOrEmpty(RowError))
                                    {
                                        //fill all the data
                                        dtWrongValues.Rows.Add(item.ItemArray);
                                        int rid = dtWrongValues.Rows.Count;
                                        dtWrongValues.Rows[rid - 1]["Error"] = RowError;
                                        //fill valid data only
                                        dtCorrectValues.Rows.Add(item.ItemArray);
                                    }
                                    else
                                    {
                                        NoErrors = false;
                                        //fill all the data
                                        dtWrongValues.Rows.Add(item.ItemArray);
                                        int rid = dtWrongValues.Rows.Count;
                                        dtWrongValues.Rows[rid - 1]["Error"] = RowError;
                                    }
                                }
                                if (NoErrors)//dtWrongValues.Rows.Count == 0// no error, create new records
                                {
                                    bool NoErrorwhenInsert = true;
                                    for (int i = 0; i < dtCorrectValues.Rows.Count; i++)
                                    {
                                        DataTable DDT = objBL.BL_ExecuteParamSP("uspManageCustomerMasterImport",
                                            dtCorrectValues.Rows[i]["Code *"].ToString(),
                                            dtCorrectValues.Rows[i]["Name *"].ToString(),
                                            dtCorrectValues.Rows[i]["Billing Address 1"].ToString(),
                                            dtCorrectValues.Rows[i]["Billing Address 2"].ToString(),
                                            dtCorrectValues.Rows[i]["Billing Address 3"].ToString(),
                                            dtCorrectValues.Rows[i]["Shipping Address 1"].ToString(),
                                            dtCorrectValues.Rows[i]["Shipping Address 2"].ToString(),
                                            dtCorrectValues.Rows[i]["Shipping Address 3"].ToString(),
                                            dtCorrectValues.Rows[i]["Pincode *"].ToString(),
                                            dtCorrectValues.Rows[i]["Contact Person"].ToString(),
                                            dtCorrectValues.Rows[i]["Phone No 1"].ToString(),
                                            dtCorrectValues.Rows[i]["Phone No 2"].ToString(),
                                            dtCorrectValues.Rows[i]["Mobile No 1"].ToString(),
                                            dtCorrectValues.Rows[i]["Mobile No 2"].ToString(),
                                            dtCorrectValues.Rows[i]["Email ID"].ToString(),
                                            dtCorrectValues.Rows[i]["PAN Number"].ToString(),
                                            dtCorrectValues.Rows[i]["Aadhar No"].ToString(),
                                            dtCorrectValues.Rows[i]["DL No 20"].ToString(),
                                            dtCorrectValues.Rows[i]["DL No 21"].ToString(),
                                            dtCorrectValues.Rows[i]["FSSAI No"].ToString(),
                                            dtCorrectValues.Rows[i]["State Name"].ToString(),
                                            dtCorrectValues.Rows[i]["GSTIN"].ToString(),
                                            dtCorrectValues.Rows[i]["Credit Term"].ToString(),
                                            dtCorrectValues.Rows[i]["Payment Mode"].ToString(),
                                            dtCorrectValues.Rows[i]["Tax Type *"].ToString(),
                                            objBL.BL_dValidation(dtCorrectValues.Rows[i]["Over Due Value"].ToString()),
                                            objBL.BL_nValidation(dtCorrectValues.Rows[i]["Over Due Inv Count"].ToString()),
                                            objBL.BL_dValidation(dtCorrectValues.Rows[i]["Credit Limit Value"].ToString()),
                                            objBL.BL_nValidation(dtCorrectValues.Rows[i]["Credit Limit Count"].ToString()),
                                            objBL.BL_dValidation(dtCorrectValues.Rows[i]["Over Due Value"].ToString()),
                                            dtCorrectValues.Rows[i]["Price Type *"].ToString(),
                                            dtCorrectValues.Rows[i]["Owner Name"].ToString(),
                                            objBL.BL_dValidation(dtCorrectValues.Rows[i]["Discount %"].ToString()),
                                            dtCorrectValues.Rows[i]["Track Point"].ToString() == "Y" ? "1" : "0", 0,
                                            dtCorrectValues.Rows[i]["TCS Tax"].ToString() == "Y" ? "1" : "0", null, null,
                                            dtCorrectValues.Rows[i]["Distance"].ToString(),
                                            dtCorrectValues.Rows[i]["Remark"].ToString(),
                                            dtCorrectValues.Rows[i]["Active"].ToString() == "Y" ? "1" : "0",
                                            objBL.BL_nValidation(UserID),
                                            dtCorrectValues.Rows[i]["Customer Type"].ToString(),
                                            dtCorrectValues.Rows[i]["Rating"].ToString());
                                        if (DDT.Columns.Count == 3)
                                        {
                                            NoErrorwhenInsert = false;
                                            dtWrongValues.Rows.Add(dtCorrectValues.Rows[i].ItemArray);
                                            int rid = dtWrongValues.Rows.Count;
                                            dtWrongValues.Rows[rid - 1]["Error"] = DDT.Rows[0][0].ToString();
                                        }
                                    }
                                    if (!NoErrorwhenInsert)
                                    {
                                        strFilePath = FPt + "Error Files\\";
                                        strFileName = TransName + "_Error_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                                        strSheetName = "Data";
                                        ExportToExcel(dtWrongValues);
                                        MTM.Add(new ImportResults()
                                        {
                                            ID = "1",
                                            Msg = "Data Saved with errors. Error occured in some documents. See error list in downloads.",
                                            FileName = strFileName + strExtension,
                                            FilePath = strFilePath + strFileName + strExtension,
                                            Total = Convert.ToString(dtData.Rows.Count),
                                            Saved = Convert.ToString(dtCorrectValues.Rows.Count),
                                            UnSaved = Convert.ToString(dtWrongValues.Rows.Count),
                                        });
                                    }
                                    else
                                    {
                                        MTM.Add(new ImportResults()
                                        {
                                            ID = "0",
                                            Msg = "Data Saved Successfully.",
                                            Total = Convert.ToString(dtData.Rows.Count),
                                            Saved = Convert.ToString(dtCorrectValues.Rows.Count),
                                            UnSaved = Convert.ToString(dtWrongValues.Rows.Count),
                                        });
                                    }
                                }
                                else
                                {
                                    //strFilePath = AppDomain.CurrentDomain.BaseDirectory + "Error Files\\";
                                    strFilePath = FPt + "Error Files\\";
                                    strFileName = TransName + "_Error_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                                    strSheetName = "Data";
                                    ExportToExcel(dtWrongValues);
                                    MTM.Add(new ImportResults()
                                    {
                                        ID = "1",
                                        Msg = "Data Not Saved. Error occured in some documents. See error list in downloads.",
                                        FileName = strFileName + strExtension,
                                        FilePath = strFilePath + strFileName + strExtension,
                                        Total = Convert.ToString(dtData.Rows.Count),
                                        Saved = Convert.ToString(dtCorrectValues.Rows.Count),
                                        UnSaved = Convert.ToString(dtWrongValues.Rows.Count),
                                    });
                                }
                                return Ok(MTM);
                            }
                            else
                            {
                                Msg = "0";// no records found;
                            }
                        }
                        #endregion
                    }
                }
                else
                {
                    Msg = "1";// file not choosing
                }
            }
            catch (Exception ex)
            {
                MTM.Add(new ImportResults()
                {
                    ID = "2",
                    Msg = ex.Message + " Date : " + dt,
                });
                return Ok(MTM);
            }
            return Ok(Msg);
        }
        public string CustomerImpValiation(DataTable dtCheck)
        {
            string RowError = "";

            if (string.IsNullOrEmpty(dtCheck.Rows[0]["Code *"].ToString()))
            {
                RowError += "Code : Code should not be empty\n";
            }
            else
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Code *"].ToString()))
                {
                    RowError += "Code : Invalid character\n";
                }
            }
            if (string.IsNullOrEmpty(dtCheck.Rows[0]["Name *"].ToString()))
            {
                RowError += "Name : Name should not be empty\n";
            }
            else
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Name *"].ToString()))
                {
                    RowError += "Name : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Customer Type"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Customer Type"].ToString()))
                {
                    RowError += "Customer Type : Invalid character\n";
                }
            }
            if (string.IsNullOrEmpty(dtCheck.Rows[0]["Price Type *"].ToString()))
            {
                RowError += "Price Type should not be empty\n";
            }
            else
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Price Type *"].ToString()))
                {
                    RowError += "Price Type : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Owner Name"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Owner Name"].ToString()))
                {
                    RowError += "Owner Name : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Contact Person"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Contact Person"].ToString()))
                {
                    RowError += "Contact Person : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Email ID"].ToString()))
            {
                if (!objBL.BL_Email(dtCheck.Rows[0]["Email ID"].ToString()))
                {
                    RowError += "Email ID : Invalid Email Format\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Phone No 1"].ToString()))
            {
                if (!objBL.BL_MobileNumberValidate(dtCheck.Rows[0]["Phone No 1"].ToString()))
                {
                    RowError += "Phone No 1 : Invalid Phone No Format\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Phone No 2"].ToString()))
            {
                if (!objBL.BL_MobileNumberValidate(dtCheck.Rows[0]["Phone No 2"].ToString()))
                {
                    RowError += "Phone No 2 : Invalid Phone No Format\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Mobile No 1"].ToString()))
            {
                if (!objBL.BL_MobileNumberValidate(dtCheck.Rows[0]["Mobile No 1"].ToString()))
                {
                    RowError += "Mobile No 1 : Invalid Mobile No Format\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Mobile No 2"].ToString()))
            {
                if (!objBL.BL_MobileNumberValidate(dtCheck.Rows[0]["Mobile No 2"].ToString()))
                {
                    RowError += "Mobile No 2 : Invalid Mobile No Format\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Billing Address 1"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Billing Address 1"].ToString()))
                {
                    RowError += "Billing Address 1 : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Billing Address 2"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Billing Address 2"].ToString()))
                {
                    RowError += "Billing Address 2 : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Billing Address 3"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Billing Address 3"].ToString()))
                {
                    RowError += "Billing Address 3 : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Shipping Address 1"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Shipping Address 1"].ToString()))
                {
                    RowError += "Shipping Address 1 : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Shipping Address 2"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Shipping Address 2"].ToString()))
                {
                    RowError += "Shipping Address 2 : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Shipping Address 3"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Shipping Address 3"].ToString()))
                {
                    RowError += "Shipping Address 3 : Invalid character\n";
                }
            }
            if (string.IsNullOrEmpty(dtCheck.Rows[0]["Pincode *"].ToString()))
            {
                RowError += "Pincode should not be empty\n";
            }
            else
            {
                if (!objBL.BL_PinNumberValidate(dtCheck.Rows[0]["Pincode *"].ToString()))
                {
                    RowError += "Pincode : Invalid character(Numbers only)\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Distance"].ToString()))
            {
                if (!objBL.BL_Numeric(dtCheck.Rows[0]["Distance"].ToString()))
                {
                    RowError += "Distance : Invalid character(Numbers only)\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Credit Limit Value"].ToString()))
            {
                if (!objBL.BL_NumericWithDecimal(dtCheck.Rows[0]["Credit Limit Value"].ToString()))
                {
                    RowError += "Credit Limit Value : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Credit Limit Count"].ToString()))
            {
                if (!objBL.BL_Numeric(dtCheck.Rows[0]["Credit Limit Count"].ToString()))
                {
                    RowError += "Credit Limit Count : Invalid character(Numbers only)\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Over Due Value"].ToString()))
            {
                if (!objBL.BL_NumericWithDecimal(dtCheck.Rows[0]["Over Due Value"].ToString()))
                {
                    RowError += "Over Due Value : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Over Due Inv Count"].ToString()))
            {
                if (!objBL.BL_Numeric(dtCheck.Rows[0]["Over Due Inv Count"].ToString()))
                {
                    RowError += "Over Due Inv Count : Invalid character(Numbers only)\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["PAN Number"].ToString()))
            {
                if (!objBL.BL_PANValidation(dtCheck.Rows[0]["PAN Number"].ToString()))
                {
                    RowError += "PAN Number : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Aadhar No"].ToString()))
            {
                if (!objBL.BL_AadhaarValidate(dtCheck.Rows[0]["Aadhar No"].ToString()))
                {
                    RowError += "Aadhar No : Invalid character(Numbers only)\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["FSSAI No"].ToString()))
            {
                if (!objBL.BL_FSSAIValidate(dtCheck.Rows[0]["FSSAI No"].ToString()))
                {
                    RowError += "FSSAI No : Invalid character(Numbers only)\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["State Name"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["State Name"].ToString()))
                {
                    RowError += "State Name : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["GSTIN"].ToString()))
            {
                if (!objBL.BL_isValidGSTIN(dtCheck.Rows[0]["GSTIN"].ToString()))
                {
                    RowError += "GSTIN : Invalid character\n";
                }
            }
            if (string.IsNullOrEmpty(dtCheck.Rows[0]["Tax Type *"].ToString()))
            {
                RowError += "Tax Type should not be empty\n";
            }
            else
            {
                if (!objBL.BL_AlphaNumeric(dtCheck.Rows[0]["Tax Type *"].ToString()))
                {
                    RowError += "Tax Type : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Payment Mode"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Payment Mode"].ToString()))
                {
                    RowError += "Payment Mode : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Credit Term"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Credit Term"].ToString()))
                {
                    RowError += "Credit Term : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Discount %"].ToString()))
            {
                if (!objBL.BL_NumericWithDecimal(dtCheck.Rows[0]["Discount %"].ToString()))
                {
                    RowError += "Discount % : Invalid character\n";
                }
                else
                {
                    if (objBL.BL_dValidation(dtCheck.Rows[0]["Discount %"].ToString()) > 100)
                    {
                        RowError += "Discount % : % should be less than 100\n";
                    }
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Remark"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Remark"].ToString()))
                {
                    RowError += "Remark : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Rating"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Rating"].ToString()))
                {
                    RowError += "Rating : Invalid character\n";
                }
            }
            if (string.IsNullOrEmpty(dtCheck.Rows[0]["TCS Tax"].ToString()))
            {
                RowError += "TCS Tax should not be empty\n\n";
            }
            else
            {
                if (dtCheck.Rows[0]["TCS Tax"].ToString().ToUpper() != "Y" && dtCheck.Rows[0]["TCS Tax"].ToString().ToUpper() != "N")
                {
                    RowError += "TCS Tax : Value should be Y or N\n";
                }
            }
            if (string.IsNullOrEmpty(dtCheck.Rows[0]["Track Point"].ToString()))
            {
                RowError += "Track Point should not be empty\n\n";
            }
            else
            {
                if (dtCheck.Rows[0]["Track Point"].ToString().ToUpper() != "Y" && dtCheck.Rows[0]["Track Point"].ToString().ToUpper() != "N")
                {
                    RowError += "Track Point : Value should be Y or N\n";
                }
            }
            if (string.IsNullOrEmpty(dtCheck.Rows[0]["Active"].ToString()))
            {
                RowError += "Active should not be empty\n\n";
            }
            else
            {
                if (dtCheck.Rows[0]["Active"].ToString().ToUpper() != "Y" && dtCheck.Rows[0]["Active"].ToString().ToUpper() != "N")
                {
                    RowError += "Active : Value should be Y or N\n";
                }
            }
            //Code *	Name *	Customer Type	Price Type *	Owner Name	Contact Person	Email ID	Phone No 1	
            //Phone No 2	Mobile No 1	Mobile No 2	Billing Address 1	Billing Address 2	Billing Address 3	
            //Shipping Address 1	Shipping Address 2	Shipping Address 3	Pincode *	Distance	Credit Limit Value	
            //Credit Limit Count	Over Due Value	Over Due Inv Count	PAN Number	Aadhar No	FSSAI No	DL No 20	
            //DL No 21	State Name	GSTIN	Tax Type *	Payment Mode	Credit Term	Discount %	Remark	Rating	
            //TCS Tax	Track Point	Active

            return RowError;
        }
        public void ColumnValidation(List<string> lst, ref bool blResult)
        {
            try
            {
                blResult = true;
                List<string> lstdtColumn = new List<string>();
                string ffp = strFilePath + strFileName;
                SpreadsheetDocument docSelected = SpreadsheetDocument.Open(strFilePath + strFileName, false);
                IEnumerable<Sheet> AllSheet = docSelected.WorkbookPart.Workbook.Descendants<Sheet>();
                strSheetName = "Data";
                Sheet sCurrent = GetSelectedSheet(AllSheet);
                if (sCurrent != null)
                {
                    Worksheet worksheet = (docSelected.WorkbookPart.GetPartById(sCurrent.Id.Value) as WorksheetPart).Worksheet;
                    IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                    // Add Header Columns
                    foreach (Row row in rows)
                    {
                        if (row.RowIndex.Value == 1)
                        {
                            foreach (Cell cell in row.Descendants<Cell>())
                            {
                                lstdtColumn.Add(GetValue(docSelected, cell));
                            }
                        }
                        break;
                    }
                    // Verify Columns Count
                    if (lst.Count != lstdtColumn.Count)
                    {
                        blResult = false;
                    }
                    string ErrMsg = "";
                    // Verify Columns Names Are Same Or Not
                    foreach (string str in lst)
                    {
                        if (!lstdtColumn.Contains(str))
                        {
                            ErrMsg = str;
                            blResult = false;
                            break;
                        }
                    }
                    if (blResult)
                    {
                        GetTable(docSelected, rows);
                        // Get the elapsed time as a TimeSpan value.
                    }
                    //docSelected.Close();
                }
            }
            catch (IOException)
            {

            }
            catch (Exception)
            {
                throw;
            }
        }
        public void GetTable(SpreadsheetDocument docSelected, IEnumerable<Row> rows)
        {
            DataTable dCheck = new DataTable();
            List<string> lstv = new List<string>();
            // Iterate Every Rows In Excel Sheet
            int TotalRowCount = rows.Count();

            decimal dRowFact = (decimal)TotalRowCount / 100;

            int TempRowCount = 0;

            foreach (Row row in rows)
            {
                if (row.RowIndex.Value == 1)
                {
                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        dCheck.Columns.Add(GetValue(docSelected, cell));
                        lstv.Add(Regex.Replace(cell.CellReference, @"[\d-]", string.Empty));
                    }
                }
                else
                {
                    dCheck.Rows.Add();
                    int nCount = 0, index = 0, TempCount;
                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        var vCellHeader = Regex.Replace(cell.CellReference, @"[\d-]", string.Empty);
                        var Temp = lstv[nCount];
                        if (lstv[nCount] != vCellHeader)
                        {
                            index = lstv.FindIndex(x => x.StartsWith(vCellHeader));
                            TempCount = nCount;
                            while (index > 0 && index > TempCount)
                            {
                                dCheck.Rows[dCheck.Rows.Count - 1][nCount] = null;
                                nCount++;
                                index--;
                            }
                        }
                        // Added By Sriram G
                        // Excel Cell Value Decimal Should be RoundOff 6 Digits
                        decimal dOutValue = 0.00M;
                        string strCellValue = GetValue(docSelected, cell);
                        //if (!string.IsNullOrEmpty(strCellValue))
                        //{
                        //    if (strCellValue.Contains('.'))
                        //    {
                        //        if (decimal.TryParse(strCellValue, out dOutValue))
                        //        {
                        //            strCellValue = Convert.ToString(Math.Round(Convert.ToDecimal(strCellValue), 6));
                        //        }
                        //    }
                        //}
                        dCheck.Rows[dCheck.Rows.Count - 1][nCount] = strCellValue;
                        nCount++;
                    }
                }

                TempRowCount++;
            }
            dtData = dCheck;
        }
        private string GetValue(SpreadsheetDocument doc, Cell cell)
        {
            try
            {
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements[(int.Parse(cell.CellValue.InnerText))].InnerText;
                    return null;
                }
                else
                if (cell.StyleIndex != null)
                {
                    CellFormat cf = doc.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[int.Parse(cell.StyleIndex.InnerText)] as CellFormat;
                    if (cf.NumberFormatId == 14)
                    {
                        return DateTime.FromOADate(double.Parse(cell.CellValue.InnerText)).ToString("dd/MM/yyyy");
                    }
                    return cell.InnerText;
                }
                else
                {
                    return cell.InnerText;
                }
            }
            catch (NullReferenceException)
            {
                return null;
            }
            catch
            {
                throw;
            }
        }
        private Sheet GetSelectedSheet(IEnumerable<Sheet> Sheets)
        {
            foreach (Sheet sName in Sheets)
            {
                if (sName.Name == strSheetName)
                {
                    return sName;
                }
            }
            return null;
        }
        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("api/generatetemplate")]
        public HttpResponseMessage GenerateTemplate(int TransID, string TransName)
        {
            DataTable dt = new DataTable();
            List<string> strTemp = null;
            if (TransID == 1)
            {
                strTemp = CustomerMasterTemp();
            }            
            OpenTemplate(strTemp, 1, TransID, TransName);
            var sDocument = strFilePath + strFileName + strExtension;
            byte[] fileBytes = System.IO.File.ReadAllBytes(sDocument);
            string fileName = strFileName + strExtension;
            //return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            if (!File.Exists(strFilePath + strFileName + strExtension))
                return new HttpResponseMessage(HttpStatusCode.NotFound);

            var result = new HttpResponseMessage(HttpStatusCode.OK);
            var stream = new FileStream(strFilePath + strFileName + strExtension, FileMode.Open, FileAccess.Read);
            result.Content = new StreamContent(stream);
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = fileName
            };
            return result;
        }
        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("api/exporttemplate")]
        public HttpResponseMessage ExportData(int TransID, string TransName)
        {
            List<string> strTemp = null;
            if (TransID == 1)
            {
                strTemp = CustomerMasterTemp();
            }           
            DataTable dt = new DataTable();
            OpenTemplate(strTemp, 2, TransID, TransName);
            var sDocument = strFilePath + strFileName + strExtension;
            byte[] fileBytes = System.IO.File.ReadAllBytes(sDocument);
            string fileName = strFileName + strExtension;
            //return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            if (!File.Exists(strFilePath + strFileName + strExtension))
                return new HttpResponseMessage(HttpStatusCode.NotFound);

            var result = new HttpResponseMessage(HttpStatusCode.OK);
            var stream = new FileStream(strFilePath + strFileName + strExtension, FileMode.Open, FileAccess.Read);
            result.Content = new StreamContent(stream);
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = fileName
            };
            return result;
            //return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("api/downloaderrordata")]
        public HttpResponseMessage DownloadErrorData(string FPath, string FName)
        {
            DataTable dt = new DataTable();
            //OpenTemplate(ImpDataTemp());
            //var sDocument = FPath;
            //byte[] fileBytes = System.IO.File.ReadAllBytes(sDocument);
            //string fileName = FName;
            var sDocument = FPath;
            byte[] fileBytes = System.IO.File.ReadAllBytes(sDocument);
            string fileName = FName;
            //return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            if (!File.Exists(FPath))
                return new HttpResponseMessage(HttpStatusCode.NotFound);

            var result = new HttpResponseMessage(HttpStatusCode.OK);
            var stream = new FileStream(FPath, FileMode.Open, FileAccess.Read);
            result.Content = new StreamContent(stream);
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = fileName
            };
            return result;
            //return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        public void OpenTemplate(List<string> str, int Type, int TransID, string TransName, string FromDate = null, string ToDate = null)
        {
            DataTable dt = new DataTable();
            DataTable dtDefaultData = new DataTable();
            DataTable dtExampleData = new DataTable();
            string FPt = System.Configuration.ConfigurationManager.AppSettings["SupportFilePath"];
            //strFilePath = AppDomain.CurrentDomain.BaseDirectory + "\\Export Data\\";
            strFilePath = FPt + "\\Export Data\\";
            strFileName = TransName + (Type == 1 ? "_Import_" : "_Export_") + DateTime.Now.ToString("yyyyMMddHHmmss");
            strSheetName = TransName;            
            if (Type == 1)
            {
                if (TransID == 9)
                {
                    dt = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 1, TransID, null, DateTime.Now.AddYears(-21), DateTime.Now.AddYears(-20));
                    dtDefaultData = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 2, TransID, DateTime.Now.AddYears(-21), DateTime.Now.AddYears(-20));
                }
                else
                {
                    foreach (string strHeaderName in str)
                    {
                        dt.Columns.Add(strHeaderName, typeof(string));
                    }
                    dtDefaultData = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 2, TransID, null, FromDate, ToDate);
                    dtExampleData = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 3, TransID, null, FromDate, ToDate);
                }
            }
            else if (Type == 2)
            {
                dt = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 1, TransID, null, FromDate, ToDate);
                dtDefaultData = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 2, TransID, null, FromDate, ToDate);
                dtExampleData = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 3, TransID, null, FromDate, ToDate);
            }
            strSheetName = "Data";
            if (TransID == 9)
            {
                ExportToExcelTwoSheet(dt, "Header", dtDefaultData, "Items");
            }
            else
            {
                ExportToExcelThreeSheet(dt, "Data", dtDefaultData, "Default Data", dtExampleData, "Example");
            }
        }
        public void ExportToExcel(DataTable DtData)
        {
            try
            {
                //Exporting to Excel
                if (!Directory.Exists(strFilePath))
                {
                    Directory.CreateDirectory(strFilePath);
                }
                using (XLWorkbook wb = new XLWorkbook())
                {
                    Int32 len = strSheetName.Length;
                    wb.Worksheets.Add(DtData, strSheetName.Substring(0, len).Trim());
                    wb.SaveAs(strFilePath + strFileName + strExtension);
                }
            }
            catch (IOException ex)
            {
                objBL.BL_WriteErrorMsginLog("Import/Emport", "ExportToExcel", ex.Message);
            }
            catch (Exception ex)
            {
                objBL.BL_WriteErrorMsginLog("Import/Emport", "ExportToExcel", ex.Message);
                throw;
            }
        }
        public void ExportToExcelTwoSheet(DataTable DtDataSheet1, string Sheet1Name, DataTable DtDataSheet2, string Sheet2Name)
        {
            try
            {
                //Exporting to Excel
                if (!Directory.Exists(strFilePath))
                {
                    Directory.CreateDirectory(strFilePath);
                }
                using (XLWorkbook wb = new XLWorkbook())
                {
                    Int32 len = Sheet1Name.Length;
                    wb.Worksheets.Add(DtDataSheet1, Sheet1Name.Substring(0, len).Trim());
                    len = Sheet2Name.Length;
                    wb.Worksheets.Add(DtDataSheet2, Sheet2Name.Substring(0, len).Trim());
                    wb.SaveAs(strFilePath + strFileName + strExtension);
                }
            }
            catch (IOException ex)
            {
                objBL.BL_WriteErrorMsginLog("Import Export", "ExportToExcelTwoSheet", ex.Message);
            }
            catch (Exception ex)
            {
                objBL.BL_WriteErrorMsginLog("Import Export", "ExportToExcelTwoSheet 1", ex.Message);
                throw;
            }
        }
        public void ExportToExcelThreeSheet(DataTable DtDataSheet1, string Sheet1Name, DataTable DtDataSheet2, string Sheet2Name, DataTable DtDataSheet3, string Sheet3Name)
        {
            try
            {
                //Exporting to Excel
                if (!Directory.Exists(strFilePath))
                {
                    Directory.CreateDirectory(strFilePath);
                }
                using (XLWorkbook wb = new XLWorkbook())
                {
                    Int32 len = Sheet1Name.Length;
                    wb.Worksheets.Add(DtDataSheet1, Sheet1Name.Trim());//.Substring(0, len)
                    len = Sheet2Name.Length;
                    wb.Worksheets.Add(DtDataSheet2, Sheet2Name.Trim());//.Substring(0, len)
                    len = Sheet3Name.Length;
                    wb.Worksheets.Add(DtDataSheet3, Sheet3Name.Trim());//.Substring(0, len)
                    wb.SaveAs(strFilePath + strFileName + strExtension);
                }
            }
            catch (IOException ex)
            {
                objBL.BL_WriteErrorMsginLog("Import Export", "ExportToExcelThreeSheet", ex.Message);
            }
            catch (Exception ex)
            {
                objBL.BL_WriteErrorMsginLog("Import Export", "ExportToExcelThreeSheet 1", ex.Message);
                throw;
            }
        }
        public static List<string> CustomerMasterTemp()
        {
            return new List<string>()
            {
                "Code *","Name *","Customer Type","Price Type *","Owner Name","Contact Person","Email ID",
                "Phone No 1","Phone No 2","Mobile No 1","Mobile No 2","Billing Address 1","Billing Address 2",
                "Billing Address 3","Shipping Address 1","Shipping Address 2","Shipping Address 3","Pincode *","Distance",
                "Credit Limit Value","Credit Limit Count","Over Due Value","Over Due Inv Count","PAN Number","Aadhar No",
                "FSSAI No","DL No 20","DL No 21","State Name","GSTIN","Tax Type *","Payment Mode","Credit Term","Discount %",
                "Remark","Rating","TCS Tax","Track Point","Active"
                };
        }
        public static List<string> CustomerMasterTempWithErrCol()
        {
            return new List<string>()
            {
                "Code *","Name *","Customer Type","Price Type *","Owner Name","Contact Person","Email ID",
                "Phone No 1","Phone No 2","Mobile No 1","Mobile No 2","Billing Address 1","Billing Address 2",
                "Billing Address 3","Shipping Address 1","Shipping Address 2","Shipping Address 3","Pincode *","Distance",
                "Credit Limit Value","Credit Limit Count","Over Due Value","Over Due Inv Count","PAN Number","Aadhar No",
                "FSSAI No","DL No 20","DL No 21","State Name","GSTIN","Tax Type *","Payment Mode","Credit Term","Discount %",
                "Remark","Rating","TCS Tax","Track Point","Active","Error"
                };
        }
    }
}
