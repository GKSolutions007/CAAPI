using CAVISTAAPI.BuisnessLayer;
using CAVISTAAPI.DALHelper;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CAVISTAAPI.Models;
using System.Web.Http.Cors;
using System.Text;
using System.Web;

namespace CAVISTAAPI.Controllers
{
    //[EnableCors(origins: "*", headers: "*", methods: "*")]
    public class LoginController : ApiController
    {
        clsBusinessLayer bl = new clsBusinessLayer();
        [HttpGet]
        [Route("api/signup/get")]
        public IHttpActionResult GetsignupData(string Mode, string Name)
        {
            if (Mode == "3" || Mode == "4")
            {
                DataTable DDT = new DataTable();
                List<UserModel> list = new List<UserModel>();
                DDT = bl.BL_ExecuteParamSP("uspManageUsers", Mode, Name);
                for (int i = 0; i < DDT.Rows.Count; i++)
                {
                    list.Add(new UserModel
                    {
                        ID = DDT.Rows[i]["ID"].ToString(),
                        UserName = DDT.Rows[i]["UserName"].ToString(),
                        Active = DDT.Rows[i]["Active"].ToString(),
                        Password = clsEncryptDecrypt.Decrypt(DDT.Rows[i]["Password"].ToString()),
                        Mobilenumber = DDT.Rows[i]["Mobilenumber"].ToString(),
                        EMailID = DDT.Rows[i]["EMailID"].ToString(),
                        RoleID = DDT.Rows[i]["RoleID"].ToString(),
                        RoleName = DDT.Rows[i]["RoleName"].ToString(),
                        PwdResetCount = DDT.Rows[i]["PwdResetCount"].ToString(),
                        PwdResetTime = DDT.Rows[i]["PwdResetTime"].ToString(),
                        LPin = DDT.Rows[i]["LPin"].ToString(),
                        UserID = DDT.Rows[i]["CBy"].ToString(),
                        CBy = DDT.Rows[i]["AUserName"].ToString(),
                        CDate = DDT.Rows[i]["LastActionTime"].ToString(),
                    });
                }
                return Ok(list);
            }
            return Ok();
        }
        [HttpPost]
        [Route("api/signup/save")]
        public IHttpActionResult Savesignup(UserModel lstMaster)
        {
            if (lstMaster != null)
            {
                List<SaveMessage> list = new List<SaveMessage>();
                DataTable DDT = bl.BL_ExecuteParamSP("uspManageUsers", lstMaster.Mode, lstMaster.ID, lstMaster.UserID, lstMaster.UserName
                    , clsEncryptDecrypt.Encrypt(lstMaster.Password), lstMaster.Mobilenumber, lstMaster.EMailID, lstMaster.RoleID, lstMaster.PwdResetCount, lstMaster.PwdResetTime,
                    lstMaster.LPin, lstMaster.Active, lstMaster.LoginUserID);
                if (DDT.Columns.Count == 1)
                {
                    if (!string.IsNullOrEmpty(lstMaster.UIURL))
                    {
                        string lk = Convert.ToString(lstMaster.UIURL);
                        string EUID = clsEncryptDecrypt.Encrypt(DDT.Rows[0][0].ToString());
                        string AALk = lk + "Login/AACM?" + "AAlk=" + HttpUtility.UrlEncode(EUID);
                        bool IsSend = bl.SendEmail("Account Activation", "Click the following link to Activate your Account : <h2 style='color:brown;'>" + AALk + "</h2>", lstMaster.EMailID);
                    }
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
        [HttpGet]
        [Route("api/signup/activateaccount")]
        public IHttpActionResult AccountActivation(string UID)
        {
            DataTable dtRes = bl.BL_ExecuteParamSP("uspManageUsers", 3, UID);
            return Ok();
        }
        [HttpGet]
        [Route("api/signup/checkapi")]
        public IHttpActionResult checkapirun(string msg)
        {
            return Ok(msg + " API RUNNING");
        }
    }
}
