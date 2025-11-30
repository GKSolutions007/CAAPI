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
using Newtonsoft.Json;
using Ionic.Zip;
using System.IO;

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
            if (Mode == "6" || Mode == "7")
            {
                DataTable DDT = new DataTable();
                DDT = bl.BL_ExecuteParamSP("uspManageUsers", Mode, Name);
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
            
            if (Mode == "5")
            {
                DataTable DDT = new DataTable();                
                DDT = bl.BL_ExecuteParamSP("uspManageUsers", Mode, Name);
                if (DDT.Rows.Count > 0)
                {
                    var list = new List<object>();
                    for (int i = 0; i < DDT.Rows.Count; i++)
                    {

                        list.Add(new 
                        {
                            ID = DDT.Rows[i]["ID"].ToString(),
                            UserID = DDT.Rows[i]["UserID"].ToString(),
                            UserName = DDT.Rows[i]["UserName"].ToString(),
                            Active = DDT.Rows[i]["Active"].ToString(),
                            Password = clsEncryptDecrypt.Decrypt(DDT.Rows[i]["Password"].ToString()),
                            Mobilenumber = DDT.Rows[i]["Mobilenumber"].ToString(),
                            EMailID = DDT.Rows[i]["EMailID"].ToString(),
                            RoleID = DDT.Rows[i]["RoleID"].ToString(),
                            //RoleName = DDT.Rows[i]["RoleName"].ToString(),
                            PwdResetCount = DDT.Rows[i]["PwdResetCount"].ToString(),
                            PwdResetTime = DDT.Rows[i]["PwdResetTime"].ToString(),
                            LPin = DDT.Rows[i]["LPin"].ToString(),                            
                        });
                    }
                    return Ok(list);
                }
                else
                {
                    return Ok();
                }                
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
                DataTable DDT = bl.BL_ExecuteParamSP("uspManageUsers", lstMaster.Mode, bl.BL_nValidation(lstMaster.ID), lstMaster.UserID, lstMaster.UserName
                    , !string.IsNullOrEmpty(lstMaster.Password) ? clsEncryptDecrypt.Encrypt(lstMaster.Password) : lstMaster.Password, lstMaster.Mobilenumber, lstMaster.EMailID, lstMaster.RoleID, lstMaster.PwdResetCount, lstMaster.PwdResetTime,
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
        [Route("api/login/get")]
        public IHttpActionResult GetloginData(string UserName, string Password)
        {
            if (!string.IsNullOrEmpty(UserName) && !string.IsNullOrEmpty(Password))
            {
                DataTable DDT = bl.BL_ExecuteParamSP("uspManageUsers", 4, 0, UserName, null, clsEncryptDecrypt.Encrypt(Password));
                List<UserModel> list = new List<UserModel>();
                if (DDT.Rows.Count > 0)
                {
                    string MsgID = "0";
                    string Message = "Login Successful.";
                    var Response = new List<object>();
                    var ValidUser = true;
                    if(DDT.Rows[0]["Active"].ToString() != "True")
                    {
                        ValidUser = false;
                        MsgID = "1";
                        Message = "User is not activated. Check mail to activate or Contact Admin";
                    }
                    else if (bl.BL_nValidation(DDT.Rows[0]["RoleID"].ToString()) == 0)
                    {
                        ValidUser = false;
                        MsgID = "1";
                        Message = "Role is not assigned for this User. Contact Admin";
                    }
                    if (ValidUser)
                    {
                        list.Add(new UserModel
                        {
                            ID = DDT.Rows[0]["ID"].ToString(),
                            UserID = DDT.Rows[0]["UserID"].ToString(),
                            UserName = DDT.Rows[0]["UserName"].ToString(),
                            Active = DDT.Rows[0]["Active"].ToString(),
                            //Password = DDT.Rows[0]["Password"].ToString(),
                            Mobilenumber = DDT.Rows[0]["Mobilenumber"].ToString(),
                            EMailID = DDT.Rows[0]["EMailID"].ToString(),
                            RoleID = DDT.Rows[0]["RoleID"].ToString(),                            
                        });
                        bl.BL_ExecuteParamSP("uspManageLoginDetails", 1, DDT.Rows[0]["ID"].ToString(), null, null, 1);
                        //return Ok(list);
                    }
                    Response.Add(new
                    {
                        MsgID = MsgID,
                        Message = Message,
                        userinfo = list,
                    });
                    return Ok(Response);
                }
            }
            return Ok();
        }
        [HttpGet]
        [Route("api/signup/updatelogout")]
        public IHttpActionResult updatelogout(string UID)
        {
            DataTable dtRes = bl.BL_ExecuteParamSP("uspManageLoginDetails", 3, UID);
            return Ok();
        }
        [HttpGet]
        [Route("api/validatepermissions")]
        public IHttpActionResult validatepermissionsData(string UID)
        {
            DataSet ds = new DataSet();           
            DataTable dtAppconfig = bl.BL_ExecuteParamSP("uspApplicationConfigValue");
            dtAppconfig.TableName = "AppConfig";
            ds.Tables.Add(dtAppconfig);
            DataTable dtRes = bl.BL_ExecuteParamSP("uspManageUsers", 5, UID);
            dtRes.TableName = "UserData";
            ds.Tables.Add(dtRes);
            string RID = dtRes.Rows[0]["RoleID"].ToString();
            DataTable dtParent = bl.BL_ExecuteParamSP("uspMenuPermission", 1, null);
            dtParent.TableName = "ParentMenu";
            ds.Tables.Add(dtParent);
            DataTable dtPermission = bl.BL_ExecuteParamSP("uspMenuPermission", 2, RID, UID);
            dtPermission.TableName = "UserMenus";
            ds.Tables.Add(dtPermission);
            //DataTable dtReportParent = bl.BL_ExecuteParamSP("uspReportPermission", 1, RID);
            //dtReportParent.TableName = "ParentRepMenu";
            //ds.Tables.Add(dtReportParent);
            //DataTable dtReportPermission = bl.BL_ExecuteParamSP("uspReportPermission", 2, RID, UID);
            //dtReportPermission.TableName = "UserRepMenus";
            //ds.Tables.Add(dtReportPermission);
            string dtjson = JsonConvert.SerializeObject(ds);
            return Ok(dtjson);
        }
        [HttpGet]
        [Route("api/backupfile")]
        public IHttpActionResult backupfile()
        {
            List<SaveMessage> list = new List<SaveMessage>();
            try
            {
                string MgID = "1", Msg = "", bakfullpath = "", filename = "";
                DataTable dt = bl.BL_ExecuteParamSP("uspAutoBackUp", 1);
                if (dt.Rows.Count > 0)
                {
                    if (dt.Columns.Count != 3)
                    {
                        bakfullpath = dt.Rows[0][1].ToString();
                        filename = dt.Rows[0][0].ToString();
                        string tempsource = dt.Rows[0][1].ToString();
                        MgID = "1";
                        Msg = "Backup Successfully";
                        //tempsource = @"D:\Host\NewWebShineAPI\BACKUPFILES\BK_DEV001_APR2024_20250925112574553.bak";
                        if (File.Exists(tempsource))
                        {
                            MgID = "0";
                            string sourceFile = tempsource;
                            string Dest = tempsource.Replace(".bak", ".zip");
                            string zipPath = Dest;// @"D:\2025\shineweb_ui\shaineweb.ui\BAK\BK_GKBS01_APR2024_20250503153834267.zip";                            
                            using (Ionic.Zip.ZipFile zfp = new Ionic.Zip.ZipFile())
                            {
                                zfp.UseZip64WhenSaving = Zip64Option.AsNecessary; // ✅ Allow large files
                                FileInfo fi = new FileInfo(sourceFile);
                                zfp.AddFile(sourceFile);
                                DirectoryInfo dff = new DirectoryInfo(sourceFile);
                                zfp.Save(zipPath);

                                if (File.Exists(sourceFile))
                                {
                                    //bl.BL_WriteErrorMsginLog("Source File", "File Path", sourceFile);
                                    File.Delete(sourceFile);
                                }
                            }


                        }
                        else
                        {
                            Msg = "Backup File does not exists for zip";
                        }
                    }
                    else
                    {
                        Msg = "Backup failed. " + dt.Rows[0][0].ToString();
                    }
                }
                list.Add(new SaveMessage()
                {
                    MsgID = MgID,
                    Message = Msg,
                    FileName = filename,
                    FilePath = bakfullpath
                });
                return Ok(list);
            }
            catch (Exception ex)
            {
                bl.BL_WriteErrorMsginLog("Backup", "Backup", ex.Message);
            }
            list.Add(new SaveMessage()
            {
                MsgID = "1",
                Message = "Back up completed. But zip file failed"
            });
            return Ok(list);
        }
        [HttpGet]
        [Route("api/signup/manageforgotpassword")]
        public IHttpActionResult ManageForgotPassword(int Flag, string Val1, string Val2, string Val3, string Val4)
        {
            try
            {
                var objNames = new List<object>();
                DataTable dt = new DataTable();
                if (Flag == 1)
                {
                    dt = bl.BL_ExecuteParamSP("uspManageForgotPassword", Flag, Val1);
                    if (dt.Rows.Count > 0)
                    {
                        string RecEmail = dt.Rows[0]["EMailID"].ToString();
                        string UID = dt.Rows[0]["ID"].ToString();
                        string UserName = dt.Rows[0]["UserName"].ToString();
                        objNames.Add(new
                        {
                            ID = "2",
                            Msg = "User verified",
                            Email = RecEmail,
                            UID = UID,
                            UserName = UserName
                        });
                    }
                    else
                    {
                        objNames.Add(new
                        {
                            ID = "1",
                            Msg = "User is not exists, Invalid User ID or E-Mail ID",
                        });
                    }
                }
                if (Flag == 2)
                {
                    Random rd = new Random();
                    int otp = rd.Next(100000, 999999);
                    bool IsSend = bl.SendEmail("Pasword Recovery", "You recovery OTP : <h2 style='color:brown;'>" + otp.ToString() + "</h2>", Val2);
                    if (IsSend)
                    {
                        int OTPID = 0;
                        DataTable dtOTP = bl.BL_ExecuteParamSP("uspManageOTP", 1, 0, "forgotpassword", otp);
                        if (dtOTP.Rows.Count > 0)
                        {
                            OTPID = Convert.ToInt32(dtOTP.Rows[0][0].ToString());
                        }
                        objNames.Add(new
                        {
                            ID = "2",
                            Msg = "OTP is send to Recovery E-Mail ID.",
                            Val1 = OTPID.ToString()
                        });
                    }
                    else
                    {
                        objNames.Add(new
                        {
                            ID = "3",
                            Msg = "Recovery E-Mail is not sending. Please try again",
                        });
                    }
                }
                else if (Flag == 3)
                {
                    dt = bl.BL_ExecuteParamSP("uspManageForgotPassword", 2, Val1, clsEncryptDecrypt.Encrypt(Val3));
                    if (dt.Rows.Count > 0)
                    {
                        objNames.Add(new
                        {
                            ID = "1",
                            Msg = "Password Reset Successfully.",
                        });
                    }
                    else
                    {
                        objNames.Add(new
                        {
                            ID = "2",
                            Msg = "Password Reset failed. Give the correct values",
                        });
                    }
                }
                return Ok(objNames);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        [HttpGet]
        [Route("api/signup/manageotp")]
        public IHttpActionResult ManageOTP(int Mode, string FormName, int ID, string ToMail, string OTP)
        {
            var objNames = new List<object>();
            if (Mode == 1)// Sending OTP to ToMail ID and OTP store in table
            {
                Random rd = new Random();
                int otp = rd.Next(100000, 999999);
                bool IsSend = bl.SendEmail("VISTA Authentication OTP", "Your Authentication OTP : <h2 style='color:brown;'>" + otp.ToString() + "</h2>", ToMail);
                if (IsSend)
                {
                    int OTPID = 0;
                    DataTable dtOTP = bl.BL_ExecuteParamSP("uspManageOTP", 1, 0, FormName, otp);
                    if (dtOTP.Rows.Count > 0)
                    {
                        OTPID = Convert.ToInt32(dtOTP.Rows[0][0].ToString());
                    }
                    objNames.Add(new
                    {
                        ID = "2",
                        Msg = "OTP is send to given E-Mail ID.",
                        Val1 = OTPID.ToString()
                    });
                }
                else
                {
                    objNames.Add(new
                    {
                        ID = "3",
                        Msg = "OTP E-Mail is not sending. Please check E-mail ID and try again",
                    });
                }
            }
            else if (Mode == 2)
            {
                DataTable dtOTP = bl.BL_ExecuteParamSP("uspManageOTP", 2, ID, FormName, OTP);
                if (dtOTP.Rows.Count > 0)
                {
                    objNames.Add(new
                    {
                        ID = "0",
                        Msg = "OTP Authentication Success.",
                        Val1 = ""
                    });
                }
                else
                {
                    objNames.Add(new
                    {
                        ID = "1",
                        Msg = "Invalid OTP",
                        Val1 = ""
                    });
                }
            }
            return Ok(objNames);
        }
        [HttpGet]
        [Route("api/signup/checkapi")]
        public IHttpActionResult checkapirun(string msg)
        {
            return Ok(msg + " API RUNNING");
        }
    }
}
