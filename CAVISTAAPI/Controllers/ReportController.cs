using CAVISTAAPI.BuisnessLayer;
using CAVISTAAPI.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace CAVISTAAPI.Controllers
{
    public class ReportController : ApiController
    {
        public string strExtension = ".xlsx";
        public string strFileName = "";
        public string strSheetName { get; set; }
        public string strFilePath
        {
            get; set;
        }
        clsBusinessLayer bl = new clsBusinessLayer();
        [HttpGet]
        [Route("api/Reportpermissions")]
        public IHttpActionResult GetPermissionsReports(string UID,string RID)
        {
            DataSet ds = new DataSet();
            DataTable dtReportParent = bl.BL_ExecuteParamSP("uspReportPermission", 1, RID);
            dtReportParent.TableName = "ParentRepMenu";
            ds.Tables.Add(dtReportParent);
            DataTable dtReportPermission = bl.BL_ExecuteParamSP("uspReportPermission", 2, RID, UID);
            dtReportPermission.TableName = "UserRepMenus";
            ds.Tables.Add(dtReportPermission);
            string dtjson = JsonConvert.SerializeObject(ds);
            return Ok(dtjson);
        }
        [HttpGet]
        [Route("api/reportscript/get")]
        public HttpResponseMessage GetreportscriptData(string ReportID, string ReportName)
        {
            string strAppStartPath = System.Configuration.ConfigurationManager.AppSettings["SupportFilePath"] + "\\Report_Script_Data\\";
            if (!Directory.Exists(strAppStartPath))
            {
                Directory.CreateDirectory(strAppStartPath);
            }
            string strFileName = ReportName + "_" + DateTime.Now.ToString("yyyymmddhhmmss") + ".txt";
            using (StreamWriter sw = System.IO.File.CreateText(System.IO.Path.Combine(strAppStartPath, strFileName)))
            {
                DataTable dt = new DataTable();
                for (int nCount = 1; nCount <= 10; nCount++)
                {
                    dt = bl.BL_ExecuteParamSP("uspReportScript", nCount, ReportID);
                    if (dt.Rows.Count > 0)
                    {
                        for (int iRow = 0; iRow < dt.Rows.Count; iRow++)
                        {
                            if (nCount == 8)
                            {
                                string strQuery = dt.Rows[iRow][0].ToString(), strColumnQuery = "";
                                for (int iCol = 1; iCol < dt.Columns.Count; iCol++)
                                {
                                    strColumnQuery = strColumnQuery + (string.IsNullOrEmpty(Convert.ToString(dt.Rows[iRow][iCol])) ? (iCol == dt.Columns.Count - 1 ? "NULL" : "NULL,")
                                        : "'" + Convert.ToString(dt.Rows[iRow][iCol]) + (iCol == dt.Columns.Count - 1 ? "'" : "',"));
                                }
                                strQuery = strQuery + strColumnQuery + ")";
                                sw.WriteLine(strQuery);
                                if (iRow == (dt.Rows.Count - 1))
                                {
                                    sw.WriteLine("GO");
                                }
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(dt.Rows[iRow][0]).Trim()))
                                    sw.WriteLine(Convert.ToString(dt.Rows[iRow][0]).Trim());
                            }
                        }
                        sw.WriteLine("");
                    }

                }
            }
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            var sDocument = System.IO.Path.Combine(strAppStartPath, strFileName);
            byte[] fileBytes = System.IO.File.ReadAllBytes(sDocument);
            //return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            if (!System.IO.File.Exists(System.IO.Path.Combine(strAppStartPath, strFileName)))
                return new HttpResponseMessage(HttpStatusCode.NotFound);

            var result = new HttpResponseMessage(HttpStatusCode.OK);
            var stream = new FileStream(System.IO.Path.Combine(strAppStartPath, strFileName), FileMode.Open, FileAccess.Read);
            result.Content = new StreamContent(stream);
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("text/plain");
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = strFileName
            };
            return result;
        }
        [HttpGet]
        [Route("api/reportcolumnsettings/getRepColumnSettings")]
        public IHttpActionResult GetColumnData(string Mode, string ReportID, string TableID)
        {
            if (Mode == "1")
            {
                DataTable dtResult = bl.BL_ExecuteParamSP("uspGetSetReportColumnSettings", Mode, ReportID, TableID);
                string JSONCONV = JsonConvert.SerializeObject(dtResult);
                return Ok(JSONCONV);
            }
            if (Mode == "2")
            {
                List<ReportColumnDataModel> list = new List<ReportColumnDataModel>();
                DataTable dtResult = bl.BL_ExecuteParamSP("uspGetSetReportColumnSettings", Mode, ReportID, TableID);
                for (int i = 0; i < dtResult.Rows.Count; i++)
                {                    
                    list.Add(new ReportColumnDataModel()
                    {
                        field = dtResult.Rows[i]["ColumnName"].ToString(),
                        header = dtResult.Rows[i]["DisplayColumnName"].ToString(),
                        visible = dtResult.Rows[i]["Visible"].ToString() == "1" ? true : false,
                        width = Convert.ToInt32(dtResult.Rows[i]["Width"].ToString()),
                        ShowinColumnOption = dtResult.Rows[i]["IsHiddenColumn"].ToString() == "1" ? false : true,
                        align = dtResult.Rows[i]["Alignment"].ToString() == "1" ? "left" : dtResult.Rows[i]["Alignment"].ToString() == "2" ? "right" : "center",
                        total = dtResult.Rows[i]["Total"].ToString(),
                        TotalYN = dtResult.Rows[i]["TotalYN"].ToString(),
                        precision = "2",
                        type = "label",
                        EnableColumnMenu = dtResult.Rows[i]["EnableColumnMenu"].ToString() == "True" ? true : false,
                        EnableSum = dtResult.Rows[i]["EnableSum"].ToString() == "True" ? true : false,
                        EnableAvg = dtResult.Rows[i]["EnableAvg"].ToString() == "True" ? true : false,
                        ClickPopup = dtResult.Rows[i]["ClickPopup"].ToString() == "1" ? true : false,
                    });
                }
                return Ok(list);
            }
            return Ok();
        }

        [HttpGet]
        [Route("api/reportparameters/get")]
        public IHttpActionResult GetData(string Mode, string ReportID, string ALName = null)
        {
            DataTable DDT = new DataTable();
            if (Mode == "0")
            {
                DDT = bl.BL_ExecuteParamSP("uspManageReports", Mode, ReportID);
                string JSONCONV = JsonConvert.SerializeObject(DDT);
                return Ok(JSONCONV);
            }
            if (Mode == "1")
            {
                DDT = bl.BL_ExecuteParamSP("uspManageReports", Mode, ReportID);
                List<ReportParameters> list = new List<ReportParameters>();
                for (int i = 0; i < DDT.Rows.Count; i++)
                {
                    list.Add(new ReportParameters
                    {
                        ParameterID = DDT.Rows[i]["ParameterID"].ToString(),
                        ReportID = DDT.Rows[i]["ReportID"].ToString(),
                        ParameterName = DDT.Rows[i]["ParameterName"].ToString(),
                        ParameterType = DDT.Rows[i]["ParameterType"].ToString(),
                        IsMandatory = DDT.Rows[i]["IsMandatory"].ToString(),
                        ParamOrder = DDT.Rows[i]["ParamOrder"].ToString(),
                        AutolistName = DDT.Rows[i]["AutolistName"].ToString()
                    });
                }
                return Ok(list);
            }
            else if (Mode == "2")
            {
                var list = new List<object>();
                DDT = bl.BL_ExecuteParamSP("uspManageReports", Mode, ReportID, ALName);
                for (int i = 0; i < DDT.Rows.Count; i++)
                {
                    list.Add(new
                    {
                        ID = DDT.Rows[i]["ID"].ToString(),
                        Name = DDT.Rows[i]["Name"].ToString(),
                    });
                }
                return Ok(list);
            }
            return Ok();
        }
        [HttpPost]
        [Route("api/reportgenerate/get")]
        public IHttpActionResult GeerateData(ReportParameters listParams)
        {
            DataTable DDT = new DataTable();
            if (listParams != null)
            {
                object[] objParamValue = new object[listParams.lstvFilters.Count];
                for (int i = 0; i < objParamValue.Length; i++)
                {
                    objParamValue[i] = !string.IsNullOrEmpty(listParams.lstvFilters[i].Param1) ? listParams.lstvFilters[i].Param1 : null;
                }
                DDT = bl.BL_ExecuteParamSP(listParams.ProcedureName, objParamValue);//, listParams.Param2, listParams.Param3, listParams.Param4
                if (DDT.Rows.Count > 0)
                {
                    string JSONCONV = JsonConvert.SerializeObject(DDT);
                    return Ok(JSONCONV);
                }
                else
                {
                    return Ok();
                }
            }
            else
            {
                return Ok();
            }
        }
        [HttpPost]
        [Route("api/reportcolumnsettings/SaveRepColumnSettings")]
        public IHttpActionResult GetColumnData(List<ReportModel> ColumnSettingData)
        {
            if (ColumnSettingData != null)
            {
                List<SaveMessage> list = new List<SaveMessage>();
                foreach (ReportModel item in ColumnSettingData)
                {
                    bl.BL_ExecuteParamSP("uspSaveReportColumnSettings", item.ReportID, item.TableID, item.ColumnID, item.ColumnName,
                      item.DisplayColumnName, item.Width, item.Visible, item.Alignment, item.DisplayIndex, item.TotalYN);
                }
                list.Add(new SaveMessage()
                {
                    MsgID = "0",
                    Message = "Saved Successfully"
                });
                return Ok(list);
            }
            return Ok();
        }
    }
}
