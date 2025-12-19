using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CAVISTAAPI.Models
{
    public class CustomerMasterModel
    {
        public string Mode { get; set; }
        public string ID { get; set; }
        public string CustomerCode { get; set; }
        public string CustomerGroupCode { get; set; }
        public string FirstName { get; set; }
        public string FatherName { get; set; }
        public string Constitution { get; set; }
        public string NameofBusiness { get; set; }
        public string BusinessPanno { get; set; }
        public string GSTNo { get; set; }
        public string PANno { get; set; }
        public string DateRegorIncorp { get; set; }
        public string DOB { get; set; }
        public string AccountManage { get; set; }
        public string ITUserName { get; set; }
        public string ITPassword { get; set; }
        public string ITFileNo { get; set; }
        public string ITRegEmailID { get; set; }
        public string ITRegContactNo { get; set; }
        public string GSTFileNo { get; set; }
        public string GSTUserName { get; set; }
        public string GSTPassword { get; set; }
        public string GSTRegDate { get; set; }
        public string GSTRegType { get; set; }
        public string DateofOPT { get; set; }
        public string GSTRegEmailID { get; set; }
        public string GSTRegContactNo { get; set; }
        public string MobileNo { get; set; }
        public string AddMobileNo { get; set; }
        public string EmailID { get; set; }
        public string AddEmailID { get; set; }
        public string LLIPINno { get; set; }
        public string TANNo { get; set; }
        public string CINNo { get; set; }
        public string DINNo { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Country { get; set; }
        public string Pincode { get; set; }
        public string Active { get; set; }
        public string Cby { get; set; }
        public List<object> FileList { get; set; }


    }
    public class MasterModel
    {
        public string Mode { get; set; }
        public string ID { get; set; }
        public string Name { get; set; }
        public string Active { get; set; }
        public string Cby { get; set; }
    }
    public class ImportResults
    {
        public string ID { get; set; }
        public string Total { get; set; }
        public string Saved { get; set; }
        public string UnSaved { get; set; }
        public string Msg { get; set; }
        public string FileName { get; set; }
        public string FilePath { get; set; }
    }
}