using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.IO;

namespace RSI.Data.Utility
{
    /// <summary>
    /// 加解密工具
    /// </summary>
    public static class EncryptUtility
    {
        /// <summary>
        /// 加密類型
        /// </summary>
        public enum EncryptType
        {
            //無
            None,
            //可逆編碼(對稱金鑰)
            AES,
            DES,
            RC2,
            TripleDES,
            //可逆編碼(非對稱金鑰)
            RSA,
            //不可逆編碼(雜湊值)
            MD5,
            SHA1,
            SHA256,
            SHA384,
            SHA512
        }

        public static EncryptType GetEncryptTypeByName(string encryptName)
        {
            EncryptType result = EncryptType.DES;

            switch (encryptName.ToUpper())
            {
                case "NONE":
                    result = EncryptType.None;
                    break;
                case "DES":
                    result = EncryptType.DES;
                    break;
                case "AES":
                    result = EncryptType.AES;
                    break;
                case "RC2":
                    result = EncryptType.RC2;
                    break;
                case "TRIPLEDES":
                    result = EncryptType.TripleDES;
                    break;
            }

            return result;
        }

        /// <summary>
        /// 產生新的KEY
        /// </summary>
        /// <param name="type">編碼器種類</param>
        public static string GenerateKey(EncryptType type)
        {
            string result = string.Empty;

            switch (type)
            {
                //可逆編碼(對稱金鑰)
                case EncryptType.AES:
                    using (AesCryptoServiceProvider csp = new AesCryptoServiceProvider())
                    {
                        csp.GenerateKey();
                        result = Convert.ToBase64String(csp.Key);
                    }
                    break;
                case EncryptType.DES:
                    using (DESCryptoServiceProvider csp = new DESCryptoServiceProvider())
                    {                        
                        csp.GenerateKey();
                        result = Convert.ToBase64String(csp.Key);
                    }
                    break;
                case EncryptType.RC2:
                    using (RC2CryptoServiceProvider csp = new RC2CryptoServiceProvider())
                    {                        
                        csp.GenerateKey();
                        result = Convert.ToBase64String(csp.Key);
                    }
                    break;
                case EncryptType.TripleDES:
                    using (TripleDESCryptoServiceProvider csp = new TripleDESCryptoServiceProvider())
                    {                        
                        csp.GenerateKey();
                        result = Convert.ToBase64String(csp.Key);
                    }
                    break;
                //可逆編碼(非對稱金鑰)
                case EncryptType.RSA:
                    using (RSACryptoServiceProvider csp = new RSACryptoServiceProvider())
                    {
                        result = csp.ToXmlString(true);
                    }
                    break;
                default:
                    break;
            }

            return result;
        }

        /// <summary>
        /// 產生新的IV
        /// </summary>
        /// <param name="type">編碼器種類</param>
        public static string GenerateIV(EncryptType type)
        {
            string result = string.Empty;

            switch (type)
            {
                //可逆編碼(對稱金鑰)
                case EncryptType.AES:
                    using (AesCryptoServiceProvider csp = new AesCryptoServiceProvider())
                    {
                        csp.GenerateIV();
                        result = Convert.ToBase64String(csp.IV);                        
                    }
                    break;
                case EncryptType.DES:
                    using (DESCryptoServiceProvider csp = new DESCryptoServiceProvider())
                    {
                        csp.GenerateIV();
                        result = Convert.ToBase64String(csp.IV);                        
                    }
                    break;
                case EncryptType.RC2:
                    using (RC2CryptoServiceProvider csp = new RC2CryptoServiceProvider())
                    {
                        csp.GenerateIV();
                        result = Convert.ToBase64String(csp.IV);
                    }
                    break;
                case EncryptType.TripleDES:
                    using (TripleDESCryptoServiceProvider csp = new TripleDESCryptoServiceProvider())
                    {
                        csp.GenerateIV();
                        result = Convert.ToBase64String(csp.IV);
                    }
                    break;
                //可逆編碼(非對稱金鑰)
                case EncryptType.RSA:
                    using (RSACryptoServiceProvider csp = new RSACryptoServiceProvider())
                    {
                        result = "";
                    }
                    break;
                default:
                    break;
            }

            return result;
        }

        /// <summary>
        /// 加密
        /// </summary>
        /// <param name="type">編碼器種類</param>  
        /// <param name="source">加密前字串</param>
        /// <returns>加密後字串</returns>
        public static string Encrypt(EncryptType type, string source, string key, string iv)
        {
            string ret = "";
            byte[] inputByteArray = Encoding.UTF8.GetBytes(source);

            switch (type)
            {
                case EncryptType.None:
                    ret = source;
                    break;
                //可逆編碼(對稱金鑰)
                case EncryptType.AES:
                    using (AesCryptoServiceProvider csp = new AesCryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(key);
                        byte[] rgbIV = Convert.FromBase64String(iv);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateEncryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Convert.ToBase64String(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncryptType.DES:
                    using (DESCryptoServiceProvider csp = new DESCryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(key);
                        byte[] rgbIV = Convert.FromBase64String(iv);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateEncryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Convert.ToBase64String(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncryptType.RC2:
                    using (RC2CryptoServiceProvider csp = new RC2CryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(key);
                        byte[] rgbIV = Convert.FromBase64String(iv);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateEncryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Convert.ToBase64String(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncryptType.TripleDES:
                    using (TripleDESCryptoServiceProvider csp = new TripleDESCryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(key);
                        byte[] rgbIV = Convert.FromBase64String(iv);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateEncryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Convert.ToBase64String(ms.ToArray());
                            }
                        }
                    }
                    break;
                //可逆編碼(非對稱金鑰)
                case EncryptType.RSA:
                    using (RSACryptoServiceProvider csp = new RSACryptoServiceProvider())
                    {
                        csp.FromXmlString(key);
                        ret = Convert.ToBase64String(csp.Encrypt(inputByteArray, false));
                    }
                    break;
                //不可逆編碼(雜湊值)
                case EncryptType.MD5:
                    using (MD5CryptoServiceProvider csp = new MD5CryptoServiceProvider())
                    {
                        ret = BitConverter.ToString(csp.ComputeHash(inputByteArray)).Replace("-", string.Empty);
                    }
                    break;
                case EncryptType.SHA1:
                    using (SHA1CryptoServiceProvider csp = new SHA1CryptoServiceProvider())
                    {
                        ret = BitConverter.ToString(csp.ComputeHash(inputByteArray)).Replace("-", string.Empty);
                    }
                    break;
                case EncryptType.SHA256:
                    using (SHA256CryptoServiceProvider csp = new SHA256CryptoServiceProvider())
                    {
                        ret = BitConverter.ToString(csp.ComputeHash(inputByteArray)).Replace("-", string.Empty);
                    }
                    break;
                case EncryptType.SHA384:
                    using (SHA384CryptoServiceProvider csp = new SHA384CryptoServiceProvider())
                    {
                        ret = BitConverter.ToString(csp.ComputeHash(inputByteArray)).Replace("-", string.Empty);
                    }
                    break;
                case EncryptType.SHA512:
                    using (SHA512CryptoServiceProvider csp = new SHA512CryptoServiceProvider())
                    {
                        ret = BitConverter.ToString(csp.ComputeHash(inputByteArray)).Replace("-", string.Empty);
                    }
                    break;
            }
            return ret;
        }

        /// <summary>
        /// 解密
        /// </summary>
        /// <param name="type">編碼器種類</param>  
        /// <param name="source">解密前字串</param>
        /// <returns>解密後字串</returns>
        public static string Decrypt(EncryptType type, string source, string key, string iv)
        {
            string ret = "";
            byte[] inputByteArray = Convert.FromBase64String(source);

            switch (type)
            {
                case EncryptType.None:
                    ret = source;
                    break;
                //可逆編碼(對稱金鑰)
                case EncryptType.AES:
                    using (AesCryptoServiceProvider csp = new AesCryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(key);
                        byte[] rgbIV = Convert.FromBase64String(iv);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateDecryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Encoding.UTF8.GetString(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncryptType.DES:
                    using (DESCryptoServiceProvider csp = new DESCryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(key);
                        byte[] rgbIV = Convert.FromBase64String(iv);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateDecryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Encoding.UTF8.GetString(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncryptType.RC2:
                    using (RC2CryptoServiceProvider csp = new RC2CryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(key);
                        byte[] rgbIV = Convert.FromBase64String(iv);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateDecryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Encoding.UTF8.GetString(ms.ToArray());
                            }
                        }
                    }
                    break;
                case EncryptType.TripleDES:
                    using (TripleDESCryptoServiceProvider csp = new TripleDESCryptoServiceProvider())
                    {
                        byte[] rgbKey = Convert.FromBase64String(key);
                        byte[] rgbIV = Convert.FromBase64String(iv);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (CryptoStream cs = new CryptoStream(ms, csp.CreateDecryptor(rgbKey, rgbIV), CryptoStreamMode.Write))
                            {
                                cs.Write(inputByteArray, 0, inputByteArray.Length);
                                cs.FlushFinalBlock();
                                ret = Encoding.UTF8.GetString(ms.ToArray());
                            }
                        }
                    }
                    break;
                //可逆編碼(非對稱金鑰)
                case EncryptType.RSA:
                    using (RSACryptoServiceProvider csp = new RSACryptoServiceProvider())
                    {
                        csp.FromXmlString(key);
                        ret = Encoding.UTF8.GetString(csp.Decrypt(inputByteArray, false));
                    }
                    break;
            }
            return ret;
        }
    }
}
