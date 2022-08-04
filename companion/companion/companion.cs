using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using Sosinpw;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Diagnostics;

namespace sCompanion
{

    [Guid("6B7CEFEC-4506-4177-9CED-FD48F12B861C")]
    public interface ICompanion
    {
        
    }
       
    [Guid("E21D2B40-9421-4279-BE05-489143193032"),ClassInterface(ClassInterfaceType.None)]
    public class companion: DRS.Common.Scanning.CompanionBase,ICompanion 
    {
        
        bool startBatch;
        string comScanDir, comSubject, comState, comExamType;
        string midno;
        public companion()
        {
            
        }

        protected override void StartBatch(ref bool AllowStart, ref int BatchNumber, ref string OperatorName, ref string TrackingId, ref int ExpectSheets, ref string Comment)
        {
            base.StartBatch(ref AllowStart, ref BatchNumber, ref OperatorName, ref TrackingId, ref ExpectSheets, ref Comment);
                        
            try
            {
                RegistryKey mICParams = Registry.CurrentUser;
                mICParams = mICParams.OpenSubKey("software", true);
                foreach (string Keyname in mICParams.GetSubKeyNames())
                {

                    if (Keyname == "necoscan")
                    {
                        mICParams = mICParams.OpenSubKey("necoscan", true);
                        startBatch=Convert.ToBoolean (mICParams.GetValue("startBatch") );
                        comScanDir = mICParams.GetValue("scanDir").ToString ();
                        comSubject = mICParams.GetValue("subject").ToString();
                        comState = mICParams.GetValue("state").ToString();
                        comExamType = mICParams.GetValue("ExamType").ToString();
                        mICParams.SetValue("startBatch", false);
                        mICParams.Close();
                        break;
                        }

                }
                if (startBatch == false)
                {
                    string kstate;
                    if (comExamType.ToUpper().Contains("SSCE") || comExamType.ToUpper().Contains("JSCE"))
                    {
                        kstate = comState;
                    }
                    else
                    {
                        kstate = " ";
                    }
                    if (System.Windows.Forms.MessageBox.Show("Do you want to Continue Scanning \n" +comSubject + "  and "+comState +" ?", "NECO SCAN",
                        System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                    {
                        AllowStart = true;

                    }
                    else
                    {
                        
                        AllowStart = false;
                        Sosinp.Terminate(0, -1);
                        Process.Start(@"c:\program files\necoscan\scan\scan.exe");
                    }
                }
                
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message,"NECO SCAN" );
            }
 
        }



        protected override void ValidateData(ref bool DataValid, ref string OMRData, ref Sosinpw.SheetStatus ASheetStatus, string FormName, ref string ErrMsg)
        {
            base.ValidateData(ref DataValid, ref OMRData, ref ASheetStatus, FormName, ref ErrMsg);
                string mexmno;
                string strCandidateNum = Sosinp.get_FieldValueByName("exm_no");
                string mregno;
                int tlength= strCandidateNum.Trim().Length;
                midno = strCandidateNum.Substring(0, (tlength -2));
                int sumdig = 0;
                int tc = 0;
                
                try
                {
                    mexmno = midno + "00";
                    for (tc = 0; tc < mexmno.Length; tc++)
                    {
                        string tck = mexmno.Substring(tc, 1);
                        if (tck == " ")
                        {
                            tck = "0";
                        }
                        sumdig = (sumdig * 10 + Asc(tck) - 48) % 97;

                    }
                    string digit1 = Chr(Convert.ToInt32(((97 - sumdig) / 10)) + 65);
                    string digit2 = Chr(((97 - sumdig) % 10) + 65);
                    mregno = midno + digit1 + digit2;
                    if (mregno != strCandidateNum && strCandidateNum.Trim().Length == 10)
                    {
                        ErrMsg = "Invalid Check Digits " + strCandidateNum;
                        DataValid = false;
                    }
                    if (strCandidateNum.Trim().Length == 0)
                    {
                        ErrMsg = "Blank Candidate Number " + strCandidateNum;
                        DataValid = false;

                    }
                    if (strCandidateNum.Contains(">") && strCandidateNum.Trim().Length > 0)
                    {
                        ErrMsg = "Multiple Shading " + strCandidateNum;
                        DataValid = false;

                    }

                    if (strCandidateNum.Contains(" ") && strCandidateNum.Trim().Length > 0)
                    {
                        ErrMsg = "Omitted Shading " + strCandidateNum;
                        DataValid = false;

                    }
                    
                }

                catch (Exception ex)
                {
                    DataValid = false;
                    ErrMsg = "An Error occured, Form will be rejected: " + ex.Message;

                }
            
    
    
        }

        protected override void CloseDown()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            base.CloseDown();
        }
        internal string Chr(int p_intBytes)
        {
            
                if ((p_intBytes < 0) || (p_intBytes > 225))
                {
                    throw new ArgumentOutOfRangeException("p_intBytes", p_intBytes, "must be between 1 and 225");
                }
                byte[] byteBuffer = new byte[] { (byte)p_intBytes };
            
                            
            return Encoding.GetEncoding(1252).GetString(byteBuffer);
        }
        internal int Asc(string p_strChar)
        {
            
                if ((p_strChar.Length == 0) || (p_strChar.Length > 1))
                {
                    throw new ArgumentOutOfRangeException("p_strChar", p_strChar, "must be a single character");
                }
                char[] chrBuffer = { Convert.ToChar(p_strChar) };
                byte[] byteBuffer = Encoding.GetEncoding(1252).GetBytes(chrBuffer);
                        
            return (int)byteBuffer [0];
        }

        
    }
}
