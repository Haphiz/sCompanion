using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace sCompanion
{
    static class Tools
    {
        static int intRetVal;
        static string strRetVal;
        public static string Chr(int p_intBytes)
        {

            if ((p_intBytes < 0) || (p_intBytes > 225))
            {
                throw new ArgumentOutOfRangeException("p_intBytes", p_intBytes, "must be between 1 and 225");
            }
            byte[] byteBuffer = new byte[] { (byte)p_intBytes };


            return Encoding.GetEncoding(1252).GetString(byteBuffer);
        }
        public static int Asc(string p_strChar)
        {

            if ((p_strChar.Length == 0) || (p_strChar.Length > 1))
            {
                throw new ArgumentOutOfRangeException("p_strChar", p_strChar, "must be a single character");
            }
            char[] chrBuffer = { Convert.ToChar(p_strChar) };
            byte[] byteBuffer = Encoding.GetEncoding(1252).GetBytes(chrBuffer);

            return (int)byteBuffer[0];
        }

        public static int getCandidatesOnSheet(string code)
        {
            int[] value={0,1,2,3,4,5,6,7,8,9};
            string[] mcode = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" };

            for (int i = 0; (i <= mcode.Length); i++)
            {
                if (mcode[i] == code.Trim())
                {
                    intRetVal = value[i];
                    break;
                }
            }

            return intRetVal;

        }
        public static string ReadRegistryValue(string regKey)
        {
            RegistryKey mICParams = Registry.CurrentUser;
            mICParams = mICParams.OpenSubKey("software", true);
            foreach (string Keyname in mICParams.GetSubKeyNames())
            {

                if (Keyname == "necoscan")
                {
                    mICParams = mICParams.OpenSubKey("necoscan", true);
                    strRetVal =   mICParams.GetValue(regKey).ToString();
                    break;
                 }

            }
            mICParams.Close();
            return strRetVal;
        }
        public static void WriteRegistryValue(string regKey,string regValue)
        {
            RegistryKey mICParams = Registry.CurrentUser;
            mICParams = mICParams.OpenSubKey("software", true);
            foreach (string Keyname in mICParams.GetSubKeyNames())
            {

                if (Keyname == "necoscan")
                {
                    mICParams = mICParams.OpenSubKey("necoscan", true);
                    mICParams.SetValue(regKey,regValue);
                    break;
                }

            }
            mICParams.Close();
        }
       
        public static string Right(string str, int index)
        {
            string iRight = "";
            iRight = str.Substring((str.Length - index), index);
            if (iRight.ToUpper() == "TXT")
            {
                iRight = "000";
            }
            return iRight;
        }
    }
}
