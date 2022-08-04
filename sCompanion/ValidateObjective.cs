using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace sCompanion
{
    static class ValidateObjective
    {
        static string strRetVal;
        static bool passed = true;
        public static string ValidateEnglishOral(string resp,string subject)
        {
            strRetVal = "";
            if(resp.Substring(60,40).Length==0 && subject=="English")
            {
                if (MessageBox.Show("Are you sure the sheet is English Language", "Companion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    strRetVal = "Wrong Sheet, English Expected";
                }
                else
                {
                    strRetVal = "";
                }
      
            }
            if (resp.Substring(60, 40).Length > 0 && subject =="Oral")
            {
                if (MessageBox.Show("Are you sure the sheet is Oral English", "Companion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    strRetVal = "Wrong Sheet, Oral Expected";
                }
                else
                {
                    strRetVal = "";
                }
            }
            return strRetVal;
        }

        public static string ValidateCandidateNumber(string number)
        {
            int mexmno = number.Length - 2;
            strRetVal = "";
            if (string.IsNullOrEmpty(number.Trim()))
            {
                strRetVal = "Invalid Candidate Number: " + number;
                passed = false;
            }

            if (passed==true )
            {
                if (number.Contains(">"))
                {
                    strRetVal = "Multiple Shading: " + number;
                    passed = false;
                }
                else
                {
                    strRetVal = "";
                }
            }
            
               
            if (passed==true )
            {
                if (number.Contains(" "))
                {
                    strRetVal = "Omitted Shading: " + number;
                    passed = false;
                }
                else
                {
                    strRetVal = "";
                }
            }

            return strRetVal;
        }

        public static string CalculateCheckDigits(string number)
        {
            strRetVal="";
            string mexmno;
            int length = number.Length - 2;
            int sumdig = 0;
            string midno = number.Substring(0, length);
            mexmno = midno + "00";
            int tc = 0;
            for (tc = 0; (tc < mexmno.Length); tc++)
            {
                string tck = mexmno.Substring(tc, 1);
                if (tck == " ")
                {
                    tck = "0";
                }
                sumdig = (sumdig * 10 + Tools.Asc(tck) - 48) % 97;
            }
            string digit1 =Tools.Chr(Convert.ToInt32(((97 - sumdig) / 10)) + 65);
            string digit2 = Tools.Chr(((97 - sumdig) % 10) + 65);
            
            if (Tools.Right(number, 2) != (digit1 + digit2))
            {
                strRetVal = "Invalid Check digits " + number ;
            }
            else
            {
                strRetVal = "";
            }
            return strRetVal;

        }
    }
}
