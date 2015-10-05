using System;
//using kundan.ExceptionHandler;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Security.Cryptography;
using System.Text;
using System.IO;
using System.Xml;


/// <summary>
/// Summary description for Encryption
/// </summary>
public class Encryption
{
    public Encryption()
    {
        //
        // TODO: Add constructor logic here
        //
    }

    #region CreateSalt
    public string CreateSalt(int size)
    {
        size = 5;
        RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider();
        byte[] buff = new byte[size];
        rng.GetBytes(buff);
        return Convert.ToBase64String(buff);
    }
    #endregion CreateSalt

    #region Create password hesh value
    /// <summary>
    /// generate encrypted value for password
    /// </summary>
    /// <param name="pwd">password</param>
    /// <param name="salt">user salt value</param>
    /// <returns>encrypted password</returns>
    public string CreatePasswordHash(string pwd, string salt)
    {
        string saltAndPwd = String.Concat(pwd, salt);
        string hashedPwd = "";
        hashedPwd = FormsAuthentication.HashPasswordForStoringInConfigFile(saltAndPwd, "SHA1");  //"sha1" or "md5"
        return hashedPwd;
    }
    #endregion Create password hesh value

}


public class Encryption64
{
    private byte[] key = { };
    private byte[] IV = { 0x12, 0x34, 0x56, 0x78, 0x90, 0xab, 0xcd, 0xef };

    public string Decrypt(string stringToDecrypt)
    {
        byte[] inputByteArray = new byte[stringToDecrypt.Length + 1];
        try
        {
            key = System.Text.Encoding.UTF8.GetBytes("!#$a54?3");

            DESCryptoServiceProvider des = new DESCryptoServiceProvider();
            stringToDecrypt = stringToDecrypt.Replace(" ", "+");
            inputByteArray = Convert.FromBase64String(stringToDecrypt);
            MemoryStream ms = new MemoryStream();
            CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(key, IV), CryptoStreamMode.Write);
            cs.Write(inputByteArray, 0, inputByteArray.Length);
            cs.FlushFinalBlock();
            System.Text.Encoding encoding = System.Text.Encoding.UTF8;
            return encoding.GetString(ms.ToArray());
        }
        catch (Exception e)
        {
            return e.Message;
        }
    }
    public string Decrypt(string stringToDecrypt, string salt)
    {
        byte[] inputByteArray = new byte[stringToDecrypt.Length + 1];
        try
        {
            //key = System.Text.Encoding.UTF8.GetBytes("!#$a54?3");
            key = System.Text.Encoding.UTF8.GetBytes(salt);
            DESCryptoServiceProvider des = new DESCryptoServiceProvider();
            stringToDecrypt = stringToDecrypt.Replace(" ", "+");
            inputByteArray = Convert.FromBase64String(stringToDecrypt);
            MemoryStream ms = new MemoryStream();
            CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(key, IV), CryptoStreamMode.Write);
            cs.Write(inputByteArray, 0, inputByteArray.Length);
            cs.FlushFinalBlock();
            System.Text.Encoding encoding = System.Text.Encoding.UTF8;
            return encoding.GetString(ms.ToArray());
        }
        catch (Exception e)
        {
            return e.Message;
        }
    }

    public string Encrypt(string stringToEncrypt)
    {
        try
        {
            key = System.Text.Encoding.UTF8.GetBytes("!#$a54?3");
            DESCryptoServiceProvider des = new DESCryptoServiceProvider();
            byte[] inputByteArray = Encoding.UTF8.GetBytes(stringToEncrypt);
            MemoryStream ms = new MemoryStream();
            CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(key, IV), CryptoStreamMode.Write);
            cs.Write(inputByteArray, 0, inputByteArray.Length);
            cs.FlushFinalBlock();
            return Convert.ToBase64String(ms.ToArray());
        }
        catch (Exception e)
        {
            return e.Message;
        }
    }

}
