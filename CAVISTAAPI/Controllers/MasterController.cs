using CAVISTAAPI.BuisnessLayer;
using CAVISTAAPI.DALHelper;
using CAVISTAAPI.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace CAVISTAAPI.Controllers
{
    public class MasterController : ApiController
    {
        clsBusinessLayer bl = new clsBusinessLayer();
        [HttpGet]
        [Route("api/role/get")]
        public IHttpActionResult getcustomer(string Mode, string ID)
        {
            if (Mode == "3") //role list data
            {
                DataTable dtRole = bl.BL_ExecuteParamSP("uspManageRoleMaster", Mode);
                string Rolefilter = JsonConvert.SerializeObject(dtRole);
                return Ok(Rolefilter);
            }
            if (Mode == "4") //customer data BY ID
            {
                DataTable dtRole = bl.BL_ExecuteParamSP("uspManageRoleMaster", Mode, ID);
                if (dtRole.Rows.Count > 0)
                {
                    string Rolefilter = JsonConvert.SerializeObject(dtRole);
                    return Ok(Rolefilter);                   
                }
                return Ok();
            }
            return Ok();
        }
        [Route("api/role/save")]
        public IHttpActionResult savecustomer(MasterModel lstMaster)
        {
            if (lstMaster != null)
            {
                List<SaveMessage> list = new List<SaveMessage>();
                DataTable DDT = bl.BL_ExecuteParamSP("uspManageRoleMaster", lstMaster.Mode, bl.BL_nValidation(lstMaster.ID), lstMaster.Name,
                        lstMaster.Active, lstMaster.Cby);
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
