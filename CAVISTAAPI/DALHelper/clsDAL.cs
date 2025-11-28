using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace CAVISTAAPI.DALHelper
{
    public class clsDAL
    {
        SqlConnection aSqlconnection = new SqlConnection();
        SqlTransaction aSqlTransaction;
        public string Connectionstr()
        {
            string COns = Connection.GetConnectionString();
            return clsEncryptDecrypt.Decrypt(COns);
        }
        public DataTable dl_ExecuteSqlQuery(string strquery)
        {
            DataTable dt_SqlQuery = new DataTable();
            try
            {
                using (aSqlconnection = new SqlConnection(Connectionstr()))
                {
                    aSqlconnection.Open();
                    SqlDataAdapter dacmd = new SqlDataAdapter(strquery, aSqlconnection);
                    dacmd.Fill(dt_SqlQuery);
                    //aSqlconnection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt_SqlQuery;
        }
        public DataTable dl_ExecuteParamSP(string strSPName, params object[] parameters)
        {
            DataTable dt_returnvalue = new DataTable();
            int Errindex = 0;
            try
            {
                using (aSqlconnection = new SqlConnection(Connectionstr()))
                {
                    aSqlconnection.Open();
                    SqlCommand cmd = new SqlCommand(strSPName, aSqlconnection);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlCommandBuilder.DeriveParameters(cmd);
                    for (int i = 0; i < parameters.Length; i++)
                    {
                        Errindex = i;
                        cmd.Parameters[i + 1].Value = parameters[i];
                    }
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt_returnvalue);
                    //aSqlconnection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt_returnvalue;
        }
        public void dl_Transaction(int Action)
        {
            if (Action == 1)
            {
                aSqlconnection = new SqlConnection(Connectionstr());
                aSqlconnection.Open();
                aSqlTransaction = aSqlconnection.BeginTransaction();
            }
            else if (Action == 2)
            {
                aSqlTransaction.Commit();
                aSqlconnection.Close();
            }
            else if (Action == 3)
            {
                aSqlTransaction.Rollback();
                aSqlconnection.Close();
            }
        }
        public DataTable dl_ManageTrans(string strStoredProc, params object[] obj)
        {
            try
            {
                DataTable dt_returnvalue = new DataTable();

                SqlCommand cmd = new SqlCommand(strStoredProc, aSqlconnection, aSqlTransaction);
                cmd.CommandTimeout = 0;
                cmd.CommandType = CommandType.StoredProcedure;
                SqlCommandBuilder.DeriveParameters(cmd);
                for (int i = 0; i < obj.Length; i++)
                {
                    if (cmd.Parameters[i + 1].SqlDbType != SqlDbType.Structured)
                    {
                        cmd.Parameters[i + 1].Value = obj[i];
                    }
                    else
                    {
                        string strTypeName = cmd.Parameters[i + 1].TypeName;
                        int nIndex = strTypeName.IndexOf(".");
                        cmd.Parameters[i + 1].TypeName = strTypeName.Substring(nIndex + 1);
                        cmd.Parameters[i + 1].Value = obj[i];
                    }
                }
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt_returnvalue);
                return dt_returnvalue;
            }
            catch (Exception ex)
            {
                aSqlTransaction.Rollback();
                throw ex;
            }
        }
        public DataSet dl_ExecuteParamSPDataset(string strSPName, params object[] parameters)
        {
            DataSet dt_returnvalue = new DataSet();
            int Errindex = 0;
            try
            {
                using (aSqlconnection = new SqlConnection(Connectionstr()))
                {
                    aSqlconnection.Open();
                    SqlCommand cmd = new SqlCommand(strSPName, aSqlconnection);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlCommandBuilder.DeriveParameters(cmd);
                    for (int i = 0; i < parameters.Length; i++)
                    {
                        Errindex = i;
                        cmd.Parameters[i + 1].Value = parameters[i];
                    }
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt_returnvalue);
                    //aSqlconnection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt_returnvalue;
        }
    }
    public class clsEncryptDecrypt
    {
        public static string Encrypt(string clearText)
        {
            string strReturnValue = string.Empty;
            string EncryptionKey = "MAKV2SPBNI99212";
            byte[] clearBytes = Encoding.Unicode.GetBytes(clearText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    strReturnValue = Convert.ToBase64String(ms.ToArray());
                }
            }
            return strReturnValue;
        }
        //Decryption
        public static string Decrypt(string cipherText)
        {
            string EncryptionKey = "MAKV2SPBNI99212";
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }
    }
}