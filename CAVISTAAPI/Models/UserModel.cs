using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CAVISTAAPI.Models
{
    public class UserModel
    {
        public string Mode { get; set; }
        public string ID { get; set; }
        public string UserName { get; set; }        
        public string Password { get; set; }
        public string Mobilenumber { get; set; }
        public string EMailID { get; set; }
        public string RoleID { get; set; }
        public string RoleName { get; set; }
        public string PwdResetCount { get; set; }
        public string PwdResetTime { get; set; }
        public string LPin { get; set; }
        public string Active { get; set; }
        public string UserID { get; set; }
        public string LoginUserID { get; set; }
        public string CBy { get; set; }
        public string CDate { get; set; }
        public string UIURL { get; set; }
    }
    public class SaveMessage
    {
        public string ID { get; set; }
        public string MsgID { get; set; }
        public string Message { get; set; }
        public string RowID { get; set; }
        public string FilePath { get; set; }
        public string FileName { get; set; }
    }
}