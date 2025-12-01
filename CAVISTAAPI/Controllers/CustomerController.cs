using CAVISTAAPI.BuisnessLayer;
using CAVISTAAPI.DALHelper;
using CAVISTAAPI.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Web;
using System.Web.Http;

namespace CAVISTAAPI.Controllers
{
    public class CustomerController : ApiController
    {
        clsBusinessLayer bl = new clsBusinessLayer();

        [HttpGet]
        [Route("api/customer/get")]
        public IHttpActionResult getcustomer(string Mode, string ID)
        {
            if (Mode == "1")//initial data
            {
                DataSet dataSet = bl.BL_ExecuteParamSPDataset("uspgetsetCustomerMaster", Mode);
                DataTable dtConst = dataSet.Tables[0];
                DataTable dtStates = dataSet.Tables[1];
                string Consts = JsonConvert.SerializeObject(dtConst);
                string states = JsonConvert.SerializeObject(dtStates);
                var Response = new List<object>();
                Response.Add(Consts);
                Response.Add(states);
                return Ok(Response);
            }
            if (Mode == "2") //customer list data
            {
                DataTable dtCustomers = bl.BL_ExecuteParamSP("uspgetsetCustomerMaster", Mode);
                string Customerfilter = JsonConvert.SerializeObject(dtCustomers);
                return Ok(Customerfilter);
            }
            if (Mode == "3") //customer data BY ID
            {
                DataTable dtCustomers = bl.BL_ExecuteParamSP("uspgetsetCustomerMaster", Mode, ID);
                if (dtCustomers.Rows.Count > 0)
                {
                    List<CustomerMasterModel> list = new List<CustomerMasterModel>();
                    for (int i = 0; i < dtCustomers.Rows.Count; i++)
                    {
                        list.Add(new CustomerMasterModel
                        {
                            ID = dtCustomers.Rows[i]["ID"].ToString(),
                            CustomerCode = dtCustomers.Rows[i]["CustomerCode"].ToString(),
                            CustomerGroupCode = dtCustomers.Rows[i]["CustomerGroupCode"].ToString(),
                            FirstName = dtCustomers.Rows[i]["FirstName"].ToString(),
                            FatherName = dtCustomers.Rows[i]["FatherName"].ToString(),
                            Constitution = dtCustomers.Rows[i]["Constitution"].ToString(),
                            NameofBusiness = dtCustomers.Rows[i]["NameofBusiness"].ToString(),
                            BusinessPanno = dtCustomers.Rows[i]["BusinessPanno"].ToString(),
                            GSTNo = dtCustomers.Rows[i]["GSTNo"].ToString(),
                            PANno = dtCustomers.Rows[i]["PANno"].ToString(),
                            DateRegorIncorp = dtCustomers.Rows[i]["DateRegorIncorp"].ToString(),
                            DOB = dtCustomers.Rows[i]["DOB"].ToString(),
                            AccountManage = dtCustomers.Rows[i]["AccountManage"].ToString(),
                            ITUserName = dtCustomers.Rows[i]["ITUserName"].ToString(),
                            ITPassword = !string.IsNullOrEmpty(dtCustomers.Rows[i]["ITPassword"].ToString()) ? clsEncryptDecrypt.Decrypt(dtCustomers.Rows[i]["ITPassword"].ToString()) : null,
                            ITFileNo = dtCustomers.Rows[i]["ITFileNo"].ToString(),
                            ITRegEmailID = dtCustomers.Rows[i]["ITRegEmailID"].ToString(),
                            ITRegContactNo = dtCustomers.Rows[i]["ITRegContactNo"].ToString(),
                            GSTFileNo = dtCustomers.Rows[i]["GSTFileNo"].ToString(),
                            GSTUserName = dtCustomers.Rows[i]["GSTUserName"].ToString(),
                            GSTPassword = !string.IsNullOrEmpty(dtCustomers.Rows[i]["GSTPassword"].ToString()) ? clsEncryptDecrypt.Decrypt(dtCustomers.Rows[i]["GSTPassword"].ToString()) : null,
                            GSTRegDate = dtCustomers.Rows[i]["GSTRegDate"].ToString(),
                            GSTRegType = dtCustomers.Rows[i]["GSTRegType"].ToString(),
                            DateofOPT = dtCustomers.Rows[i]["DateofOPT"].ToString(),
                            GSTRegEmailID = dtCustomers.Rows[i]["GSTRegEmailID"].ToString(),
                            GSTRegContactNo = dtCustomers.Rows[i]["GSTRegContactNo"].ToString(),
                            MobileNo = dtCustomers.Rows[i]["MobileNo"].ToString(),
                            AddMobileNo = dtCustomers.Rows[i]["AddMobileNo"].ToString(),
                            EmailID = dtCustomers.Rows[i]["EmailID"].ToString(),
                            AddEmailID = dtCustomers.Rows[i]["AddEmailID"].ToString(),
                            LLIPINno = dtCustomers.Rows[i]["LLIPINno"].ToString(),
                            TANNo = dtCustomers.Rows[i]["TANNo"].ToString(),
                            CINNo = dtCustomers.Rows[i]["CINNo"].ToString(),
                            DINNo = dtCustomers.Rows[i]["DINNo"].ToString(),
                            Address = dtCustomers.Rows[i]["Address"].ToString(),
                            City = dtCustomers.Rows[i]["City"].ToString(),
                            State = dtCustomers.Rows[i]["State"].ToString(),
                            Country = dtCustomers.Rows[i]["Country"].ToString(),
                            Pincode = dtCustomers.Rows[i]["Pincode"].ToString(),
                            Active = dtCustomers.Rows[i]["Active"].ToString()
                        });
                    }
                    return Ok(list);
                }
                return Ok();
            }
            return Ok();
        }
            [HttpPost]
        [Route("api/customer/save")]
        public IHttpActionResult savecustomer(CustomerMasterModel lstMaster)
        {
            if (lstMaster != null)
            {
                List<SaveMessage> list = new List<SaveMessage>();
                DataTable DDT = bl.BL_ExecuteParamSP("uspManageCustomerMaster", lstMaster.Mode, bl.BL_nValidation(lstMaster.ID), lstMaster.CustomerCode,
                        lstMaster.CustomerGroupCode, lstMaster.FirstName, lstMaster.FatherName, lstMaster.Constitution, lstMaster.NameofBusiness, lstMaster.BusinessPanno,
                        lstMaster.GSTNo, lstMaster.PANno,
                        !string.IsNullOrEmpty(lstMaster.DateRegorIncorp) ? lstMaster.DateRegorIncorp : null,
                        !string.IsNullOrEmpty(lstMaster.DOB) ? lstMaster.DOB :null, 
                        lstMaster.AccountManage, lstMaster.ITUserName, 
                        !string.IsNullOrEmpty(lstMaster.ITPassword) ? clsEncryptDecrypt.Encrypt(lstMaster.ITPassword) : null,
                        lstMaster.ITFileNo, lstMaster.ITRegEmailID, lstMaster.ITRegContactNo, lstMaster.GSTFileNo, lstMaster.GSTUserName,
                        !string.IsNullOrEmpty(lstMaster.GSTPassword) ? clsEncryptDecrypt.Encrypt(lstMaster.GSTPassword) : null,
                        !string.IsNullOrEmpty(lstMaster.GSTRegDate) ? lstMaster.GSTRegDate : null, 
                        lstMaster.GSTRegType, 
                        !string.IsNullOrEmpty(lstMaster.DateofOPT) ? lstMaster.DateofOPT : null, lstMaster.GSTRegEmailID, lstMaster.GSTRegContactNo, lstMaster.MobileNo,
                        lstMaster.AddMobileNo, lstMaster.EmailID, lstMaster.AddEmailID, lstMaster.LLIPINno, lstMaster.TANNo, lstMaster.CINNo, lstMaster.DINNo,
                        lstMaster.Address, lstMaster.City, lstMaster.State, lstMaster.Country, lstMaster.Pincode, lstMaster.Active, lstMaster.Cby);
                if (DDT.Columns.Count == 1)
                {                   
                    //Success message
                    list.Add(new SaveMessage()
                    {
                        ID = DDT.Rows[0][0].ToString(),
                        MsgID = "0",
                        Message = "Saved Successfully"
                    });
                }
                else
                {
                    //Error message
                    list.Add(new SaveMessage()
                    {
                        ID = "0",
                        MsgID = "1",
                        Message = DDT.Rows[0][0].ToString(),
                    });
                }
                return Ok(list);
            }
            return Ok();
        }
    }
}
