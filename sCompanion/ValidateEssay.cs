using System;
using System.Collections.Generic;
using System.Text;

namespace sCompanion
{
    static class ValidateEssay
    {
        static string strRetVal;
        static bool passed;
        public static string ValidateRefNo(string Number)
        {
            strRetVal = "";
            if (string.IsNullOrEmpty(Number.Trim()))
            {
                strRetVal = "Invalid Ref. Number: " + Number;
                passed = false;
            }

            if (passed == true)
            {
                if (Number.Contains(">"))
                {
                    strRetVal = "Multiple Shading: " + Number;
                    passed = false;
                }
                else
                {
                    strRetVal = "";
                }
            }


            if (passed == true)
            {
                if (Number.Contains(" "))
                {
                    strRetVal = "Omitted Shading: " + Number;
                    passed = false;
                }
                else
                {
                    strRetVal = "";
                }
            }

            return strRetVal;
        }
    }
}
