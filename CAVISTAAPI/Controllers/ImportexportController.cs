using CAVISTAAPI.BuisnessLayer;
using CAVISTAAPI.DALHelper;
using CAVISTAAPI.Models;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
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
        public int CustomerID { get; set; }
        public int CustomerGroupCodeID { get; set; }
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
            string Msg = "",TName = "";
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
                    TName= HttpContext.Current.Request.Files.AllKeys[1].ToString();
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
                        dtCorrectValues.Columns.Add("ID");                        
                        if (!ErrorColAlreadyExisist)
                        {
                            dtCorrectValues.Columns.Add("Error");
                            dtWrongValues.Columns.Add("Error");
                        }
                        #region Customer
                        if (TransID == "1")
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
                                        int lastrowid = dtCorrectValues.Rows.Count;
                                        dtCorrectValues.Rows[lastrowid - 1]["ID"] = CustomerID;
                                        dtCorrectValues.Rows[lastrowid - 1]["Customer Group Code *"] = CustomerGroupCodeID;
                                        dtCorrectValues.Rows[lastrowid - 1]["Constitution *"] = ConstitutionID;
                                        dtCorrectValues.Rows[lastrowid - 1]["State Name"] = StateID;                                        
                                        string itpassword = dtCorrectValues.Rows[lastrowid - 1]["IT Password"].ToString();
                                        if (!string.IsNullOrEmpty(itpassword))
                                        {
                                            dtCorrectValues.Rows[lastrowid - 1]["IT Password"] = clsEncryptDecrypt.Encrypt(itpassword);
                                        }
                                        else
                                        {
                                            dtCorrectValues.Rows[lastrowid - 1]["IT Password"] = null;
                                        }
                                        string gstpassword = dtCorrectValues.Rows[lastrowid - 1]["GST Password"].ToString();
                                        if (!string.IsNullOrEmpty(gstpassword))
                                        {
                                            dtCorrectValues.Rows[lastrowid - 1]["GST Password"] = clsEncryptDecrypt.Encrypt(gstpassword);
                                        }
                                        else
                                        {
                                            dtCorrectValues.Rows[lastrowid - 1]["GST Password"] = null;
                                        }
                                        if(!string.IsNullOrEmpty(dtCorrectValues.Rows[lastrowid - 1]["Date of Registration/Incorporation"].ToString()))
                                        {
                                            DateTime date = DateTime.ParseExact(dtCorrectValues.Rows[lastrowid - 1]["Date of Registration/Incorporation"].ToString(), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                            string formattedDate = date.ToString("yyyy-MM-dd");
                                            dt = Convert.ToDateTime(formattedDate).ToString();//"yyyy-MM-dd"
                                            dtCorrectValues.Rows[lastrowid - 1]["Date of Registration/Incorporation"] = formattedDate;
                                        }
                                        else
                                        {
                                            dtCorrectValues.Rows[lastrowid - 1]["Date of Registration/Incorporation"] = null;
                                        }
                                        if (!string.IsNullOrEmpty(dtCorrectValues.Rows[lastrowid - 1]["Date of Birth"].ToString()))
                                        {
                                            DateTime date = DateTime.ParseExact(dtCorrectValues.Rows[lastrowid - 1]["Date of Birth"].ToString(), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                            string formattedDate = date.ToString("yyyy-MM-dd");
                                            dt = Convert.ToDateTime(formattedDate).ToString();//"yyyy-MM-dd"
                                            dtCorrectValues.Rows[lastrowid - 1]["Date of Birth"] = formattedDate;
                                        }
                                        else
                                        {
                                            dtCorrectValues.Rows[lastrowid - 1]["Date of Birth"] = null;
                                        }

                                        ////
                                        if(!string.IsNullOrEmpty(dtCorrectValues.Rows[lastrowid - 1]["GST Registration Date"].ToString()))
                                        {
                                            DateTime date = DateTime.ParseExact(dtCorrectValues.Rows[lastrowid - 1]["GST Registration Date"].ToString(), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                            string formattedDate = date.ToString("yyyy-MM-dd");
                                            dt = Convert.ToDateTime(formattedDate).ToString();//"yyyy-MM-dd"
                                            dtCorrectValues.Rows[lastrowid - 1]["GST Registration Date"] = formattedDate;
                                        }
                                        else
                                        {
                                            dtCorrectValues.Rows[lastrowid - 1]["GST Registration Date"] = null;
                                        }
                                        if (!string.IsNullOrEmpty(dtCorrectValues.Rows[lastrowid - 1]["Date of OPT"].ToString()))
                                        {
                                            DateTime date = DateTime.ParseExact(dtCorrectValues.Rows[lastrowid - 1]["Date of OPT"].ToString(), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                            string formattedDate = date.ToString("yyyy-MM-dd");
                                            dt = Convert.ToDateTime(formattedDate).ToString();//"yyyy-MM-dd"
                                            dtCorrectValues.Rows[lastrowid - 1]["Date of OPT"] = formattedDate;
                                        }
                                        else
                                        {
                                            dtCorrectValues.Rows[lastrowid - 1]["Date of OPT"] = null;
                                        }
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
                                        int customerid = objBL.BL_nValidation(dtCorrectValues.Rows[i]["ID"].ToString());
                                        DataTable DDT = objBL.BL_ExecuteParamSP("uspManageCustomerMaster", customerid > 0 ? 2 : 1,
                                            customerid,
                                            dtCorrectValues.Rows[i]["Customer Code *"].ToString(),
                                            dtCorrectValues.Rows[i]["Customer Group Code *"].ToString(),
                                            dtCorrectValues.Rows[i]["First Name *"].ToString(),
                                            dtCorrectValues.Rows[i]["Father Name *"].ToString(),
                                            dtCorrectValues.Rows[i]["Constitution *"].ToString(),
                                            dtCorrectValues.Rows[i]["Name of Business"].ToString(),
                                            dtCorrectValues.Rows[i]["Business PAN No"].ToString(),
                                            dtCorrectValues.Rows[i]["GST No"].ToString(),
                                            dtCorrectValues.Rows[i]["PAN No"].ToString(),
                                            !string.IsNullOrEmpty(dtCorrectValues.Rows[i]["Date of Registration/Incorporation"].ToString()) ? dtCorrectValues.Rows[i]["Date of Registration/Incorporation"].ToString() : null,
                                            !string.IsNullOrEmpty(dtCorrectValues.Rows[i]["Date of Birth"].ToString()) ? dtCorrectValues.Rows[i]["Date of Birth"].ToString() : null,                                            
                                            dtCorrectValues.Rows[i]["Account Manager"].ToString(),
                                            dtCorrectValues.Rows[i]["IT User Name"].ToString(),
                                            dtCorrectValues.Rows[i]["IT Password"].ToString(),
                                            dtCorrectValues.Rows[i]["IT File No"].ToString(),
                                            dtCorrectValues.Rows[i]["IT Registered Email ID"].ToString(),
                                            dtCorrectValues.Rows[i]["IT Registered Contact No"].ToString(),
                                            dtCorrectValues.Rows[i]["GST File No"].ToString(),
                                            dtCorrectValues.Rows[i]["GST User Name"].ToString(),
                                            dtCorrectValues.Rows[i]["GST Password"].ToString(),
                                            !string.IsNullOrEmpty(dtCorrectValues.Rows[i]["GST Registration Date"].ToString()) ? dtCorrectValues.Rows[i]["GST Registration Date"].ToString() : null,
                                            dtCorrectValues.Rows[i]["GST Registration Type"].ToString(),
                                            !string.IsNullOrEmpty(dtCorrectValues.Rows[i]["Date of OPT"].ToString()) ? dtCorrectValues.Rows[i]["Date of OPT"].ToString() : null,
                                            dtCorrectValues.Rows[i]["GST Registered Email ID"].ToString(),
                                            dtCorrectValues.Rows[i]["GST Registered Contact No"].ToString(),
                                            dtCorrectValues.Rows[i]["Mobile No"].ToString(),
                                            dtCorrectValues.Rows[i]["Additional Mobile No"].ToString(),
                                            dtCorrectValues.Rows[i]["Email ID"].ToString(),
                                            dtCorrectValues.Rows[i]["Additional Email ID"].ToString(),
                                            dtCorrectValues.Rows[i]["LLI PIN No"].ToString(),
                                            dtCorrectValues.Rows[i]["TAN No"].ToString(),
                                            dtCorrectValues.Rows[i]["CIN No"].ToString(),
                                            dtCorrectValues.Rows[i]["DIN No"].ToString(),
                                            dtCorrectValues.Rows[i]["Address"].ToString(),
                                            dtCorrectValues.Rows[i]["City"].ToString(),
                                            dtCorrectValues.Rows[i]["State Name"].ToString(),
                                            dtCorrectValues.Rows[i]["Country"].ToString(),
                                            dtCorrectValues.Rows[i]["Pincode"].ToString(),
                                            dtCorrectValues.Rows[i]["Active"].ToString() == "Y" ? "1" : "0",
                                            objBL.BL_nValidation(UserID)); 
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
                    else
                    {
                        Msg = "2";// column names mismatching
                    }
                }
                else
                {
                    Msg = "1";// file not choosing
                }
            }
            catch (Exception ex)
            {
                objBL.BL_WriteErrorMsginLog("Import", TName, ex.Message);
                MTM.Add(new ImportResults()
                {
                    ID = "2",
                    Msg = ex.Message,
                });
                return Ok(MTM);
            }
            return Ok(Msg);
        }
        public string CustomerImpValiation(DataTable dtCheck)
        {
            string RowError = "";

            if (string.IsNullOrEmpty(dtCheck.Rows[0]["Customer Code *"].ToString()))
            {
                RowError += "Customer Code : Code should not be empty\n";
            }
            else
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Customer Code *"].ToString()))
                {
                    RowError += "Customer Code : Invalid character\n";
                }
                else
                {
                    DataTable dt = objBL.BL_ExecuteParamSP("uspgetidfromnameforimport", "customercode", dtCheck.Rows[0]["Customer Code *"].ToString());
                    if (dt.Rows.Count == 0)
                    {
                        CustomerID = 0;
                    }
                    else
                    {
                        CustomerID = Convert.ToInt32(dt.Rows[0][0].ToString());
                    }
                }
            }
            if (string.IsNullOrEmpty(dtCheck.Rows[0]["Customer Group Code *"].ToString()))
            {
                RowError += "Customer Group Code : Customer Group Code should not be empty\n";
            }
            else
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Customer Group Code *"].ToString()))
                {
                    RowError += "Customer Group Code : Invalid character\n";
                }
                else
                {
                    string CGC = dtCheck.Rows[0]["Customer Group Code *"].ToString().ToUpper();
                    if(CGC == "TAX AUDIT")
                    {
                        CustomerGroupCodeID = 1;
                    }
                    else if(CGC == "NON TAX AUDIT")
                    {
                        CustomerGroupCodeID = 2;
                    }
                    else
                    {
                        RowError += "Customer Group Code :Invalid Group Code. Code must be TAX AUDIT or NON TAX AUDIT only \n";
                    }
                }
            }
            if (string.IsNullOrEmpty(dtCheck.Rows[0]["First Name *"].ToString()))
            {
                RowError += "First Name should not be empty\n";
            }
            else
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["First Name *"].ToString()))
                {
                    RowError += "First Name : Invalid character\n";
                }
            }

            if (string.IsNullOrEmpty(dtCheck.Rows[0]["Father Name *"].ToString()))
            {
                RowError += "Father Name should not be empty\n";
            }
            else
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Father Name *"].ToString()))
                {
                    RowError += "Father Name : Invalid character\n";
                }
            }

            if (string.IsNullOrEmpty(dtCheck.Rows[0]["Constitution *"].ToString()))
            {
                RowError += "Constitution : Code should not be empty\n";
            }
            else
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Constitution *"].ToString()))
                {
                    RowError += "Constitution * : Invalid character\n";
                }
                else
                {
                    DataTable dt = objBL.BL_ExecuteParamSP("uspgetidfromnameforimport", "constitutions", dtCheck.Rows[0]["Constitution *"].ToString());
                    if (dt.Rows.Count == 0)
                    {
                        ConstitutionID = 0;
                        RowError += "Constitution :Invalid Constitution. Check with Default sheet Constitutions. \n";
                    }
                    else
                    {
                        ConstitutionID = Convert.ToInt32(dt.Rows[0][0].ToString());
                    }
                }
            }

            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Name of Business"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Name of Business"].ToString()))
                {
                    RowError += "Name of Business : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Date of Registration/Incorporation"].ToString()))
            {
                if (!objBL.BL_DateformatDMY(dtCheck.Rows[0]["Date of Registration/Incorporation"].ToString()))
                {
                    RowError += "Date of Registration/Incorporation : Invalid Date Format(Format : dd/MM/yyyy)\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Account Manager"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Account Manager"].ToString()))
                {
                    RowError += "Account Manager : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Mobile No"].ToString()))
            {
                if (!objBL.BL_MobileNumberValidate(dtCheck.Rows[0]["Mobile No"].ToString()))
                {
                    RowError += "Mobile No : Invalid Phone No Format\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Email ID"].ToString()))
            {
                if (!objBL.BL_Email(dtCheck.Rows[0]["Email ID"].ToString()))
                {
                    RowError += "Email ID : Invalid Email Format\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Date of Birth"].ToString()))
            {
                if (!objBL.BL_DateformatDMY(dtCheck.Rows[0]["Date of Birth"].ToString()))
                {
                    RowError += "Date of Birth : Invalid Date Format(Format : dd/MM/yyyy)\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["PAN No"].ToString()))
            {
                if (!objBL.BL_PANValidation(dtCheck.Rows[0]["PAN No"].ToString()))
                {
                    RowError += "PAN No : Invalid PAN No/ Foramt\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Business PAN No"].ToString()))
            {
                if (!objBL.BL_PANValidation(dtCheck.Rows[0]["Business PAN No"].ToString()))
                {
                    RowError += "Business PAN No : Invalid PAN No/ Foramt\n";
                }
            }

            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["IT File No"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["IT File No"].ToString()))
                {
                    RowError += "IT File No : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["IT User Name"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["IT User Name"].ToString()))
                {
                    RowError += "IT User Name : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["IT Registered Contact No"].ToString()))
            {
                if (!objBL.BL_MobileNumberValidate(dtCheck.Rows[0]["IT Registered Contact No"].ToString()))
                {
                    RowError += "IT Registered Contact No : Invalid Format\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["IT Registered Email ID"].ToString()))
            {
                if (!objBL.BL_Email(dtCheck.Rows[0]["IT Registered Email ID"].ToString()))
                {
                    RowError += "IT Registered Email ID : Invalid Email Format\n";
                }
            }

            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["GST No"].ToString()))
            {
                if (!objBL.BL_isValidGSTIN(dtCheck.Rows[0]["GST No"].ToString()))
                {
                    RowError += "GST No : Invalid GST No/ Foramt\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["GST File No"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["GST File No"].ToString()))
                {
                    RowError += "GST File No : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["GST User Name"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["GST User Name"].ToString()))
                {
                    RowError += "GST User Name : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["GST Registration Date"].ToString()))
            {
                if (!objBL.BL_DateformatDMY(dtCheck.Rows[0]["GST Registration Date"].ToString()))
                {
                    RowError += "GST Registration Date : Invalid Date Format(Format : dd/MM/yyyy)\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["GST Registration Type"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["GST Registration Type"].ToString()))
                {
                    RowError += "GST Registration Type : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Date of OPT"].ToString()))
            {
                if (!objBL.BL_DateformatDMY(dtCheck.Rows[0]["Date of OPT"].ToString()))
                {
                    RowError += "Date of OPT : Invalid Date Format(Format : dd/MM/yyyy)\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["GST Registered Contact No"].ToString()))
            {
                if (!objBL.BL_MobileNumberValidate(dtCheck.Rows[0]["GST Registered Contact No"].ToString()))
                {
                    RowError += "GST Registered Contact No : Invalid Format\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["GST Registered Email ID"].ToString()))
            {
                if (!objBL.BL_Email(dtCheck.Rows[0]["GST Registered Email ID"].ToString()))
                {
                    RowError += "GST Registered Email ID : Invalid Email Format\n";
                }
            }

            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Additional Mobile No"].ToString()))
            {
                if (!objBL.BL_MobileNumberValidate(dtCheck.Rows[0]["Additional Mobile No"].ToString()))
                {
                    RowError += "Additional Mobile No : Invalid Phone No Format\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Additional Email ID"].ToString()))
            {
                if (!objBL.BL_Email(dtCheck.Rows[0]["Additional Email ID"].ToString()))
                {
                    RowError += "Additional Email ID : Invalid Email Format\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Address"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["Address"].ToString()))
                {
                    RowError += "Address : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["City"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["City"].ToString()))
                {
                    RowError += "City : Invalid character\n";
                }
            }
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["Pincode"].ToString()))
            {
                if (!objBL.BL_PinNumberValidate(dtCheck.Rows[0]["Pincode"].ToString()))
                {
                    RowError += "Pincode : Invalid character(Numbers only)\n";
                }
            }
            
           
           
            if (!string.IsNullOrEmpty(dtCheck.Rows[0]["State Name"].ToString()))
            {
                if (!objBL.BL_AlphaNumericSpl(dtCheck.Rows[0]["State Name"].ToString()))
                {
                    RowError += "State Name : Invalid character\n";
                }
                else
                {
                    DataTable dt = objBL.BL_ExecuteParamSP("uspgetidfromnameforimport", "statename", dtCheck.Rows[0]["State Name"].ToString());
                    if (dt.Rows.Count == 0)
                    {
                        StateID = 0;
                        RowError += "State Name :Invalid State Name. Check with Default sheet State Name. \n";
                    }
                    else
                    {
                        StateID = Convert.ToInt32(dt.Rows[0][0].ToString());
                    }
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
                if (Type == 1)
                {
                    foreach (string strHeaderName in str)
                    {
                        dt.Columns.Add(strHeaderName, typeof(string));
                    }
                    dtDefaultData = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 2, TransID, null);
                    dtExampleData = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 3, TransID, null);
                }
            }
            else if (Type == 2)
            {
                dt = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 1, TransID, null);
                dtDefaultData = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 2, TransID, null);
                dtExampleData = objBL.BL_ExecuteParamSP("uspgetsetImportExport", 3, TransID, null);
            }
            strSheetName = "Data";
            ExportToExcelThreeSheet(dt, "Data", dtDefaultData, "Default Data", dtExampleData, "Example");
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
                "Customer Code *",  "Customer Group Code *",    "First Name *", "Father Name *",  "Constitution *", "Name of Business", "Date of Registration/Incorporation",   "Account Manager",  "Mobile No",    "Email ID", "Date of Birth",    "PAN No",   "Business PAN No",  "IT User Name", "IT Password",  "IT File No",   "IT Registered Email ID",   "IT Registered Contact No", "LLI PIN No",   "TAN No",   "CIN No",   "DIN No",   "GST No",   "GST File No",  "GST User Name",    "GST Password", "GST Registration Date",    "GST Registration Type",    "Date of OPT",  "GST Registered Email ID",  "GST Registered Contact No",    "Additional Mobile No", "Additional Email ID",  "Address",  "City", "State Name",   "Country",  "Pincode",  "Active"
                };
        }
        public static List<string> CustomerMasterTempWithErrCol()
        {
            return new List<string>()
            {
                "Customer Code *",  "Customer Group Code *",    "First Name *", "Father Name *",  "Constitution *", "Name of Business", "Date of Registration/Incorporation",   "Account Manager",  "Mobile No",    "Email ID", "Date of Birth",    "PAN No",   "Business PAN No",  "IT User Name", "IT Password",  "IT File No",   "IT Registered Email ID",   "IT Registered Contact No", "LLI PIN No",   "TAN No",   "CIN No",   "DIN No",   "GST No",   "GST File No",  "GST User Name",    "GST Password", "GST Registration Date",    "GST Registration Type",    "Date of OPT",  "GST Registered Email ID",  "GST Registered Contact No",    "Additional Mobile No", "Additional Email ID",  "Address",  "City", "State Name",   "Country",  "Pincode",  "Active","Error"
};
        }
    }
}
