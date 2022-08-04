using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using Sosinpw;
using System.Windows.Forms;
using System.Diagnostics;
//using TextFiles;
namespace sCompanion
{
    [Guid ("6B6C05E2-87CD-4344-B21A-43B70DC7524E")]
    public interface sCompanion
    {

    }
    [Guid("DE152C82-5DEC-4260-B7A4-85EE25787305"),ClassInterface(ClassInterfaceType.None )]
    public class companion :DRS.Common.Scanning.CompanionBase,sCompanion   
    {
        #region External Dll Calls
        //Objective OMR forms
        [DllImport("ZToolBox.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern int ChkRegN(StringBuilder buffer, string numberIn, int capacity);


        [DllImport("ZToolBox.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern int ChkExm(StringBuilder buffer, string scanSubjectCodeIn, string examIn, string examCodeIn, int capacity);
         
        [DllImport("ZToolBox.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern int ValSubj(StringBuilder buffer, int capacity, string OMRSubjectCodeIn, string scanSubjectCodeIn, string scanSubjectIn, string examIn);


        //Essay OMR Forms
        [DllImport("ZToolBox.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern int ChkRefN(StringBuilder buffer, string numberIn, int capacity);


        [DllImport("ZToolBox.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern int CndOnSht(string numberIn);

        [DllImport("ZToolBox.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern void ChkScrNSht(StringBuilder buffer, string[] scores, string numberIn, int capacity);
        #endregion



        //Global Variables declaration
        string comExamtype, comSubject, comState, comJob, scanSubject, scanSubjCode, comUserName, msg, scanType;
        int SubjCodeValidCount, SubjCodeValidCountCheck, MaxSubjectCountAllowed, SheetPassCount, comExpSheets, comAnsSheet;
        bool comStartBatch, comLogOut;
        
        protected override void ValidateData(ref bool DataValid, ref string OMRData, ref SheetStatus ASheetStatus, string FormName, ref string ErrMsg)
        {

            //This method validates forms.
            base.ValidateData(ref DataValid, ref OMRData, ref ASheetStatus, FormName, ref ErrMsg);
            try
            {
                
#region OBJ Candidate Number Validation//THis code section validates all OMR reg_no fields
                //string ErrorMsg = "";
                if (FormName == "UNID")//Checking for unidentified form using the the formname reference.
                {
                    ErrMsg = "Unidentified Sheet : ";
                    DataValid = false;
                    return;
                }
                string strCandidateNum = Sosinp.get_FieldValueByName("exm_no"); //Reading Reg_no from scanner read head.
                if (comJob == "Obj")//Validateing Objective OMR forms statrts here.
                {
                    StringBuilder sb = new StringBuilder(255);
                    ChkRegN(sb, strCandidateNum, sb.Capacity);
                    //MessageBox.Show(sb.ToString().Trim());
                    if (!String.IsNullOrEmpty(sb.ToString().Trim()))
                    {
                        ErrMsg = sb.ToString().Trim();// "Invalid Candidate Number:" + strCandidateNum;
                        DataValid = false;
                        sb.Remove(0, sb.Length);
                        return;
                    }
                    
#endregion
                    #region Validating SSCE Subject codes
                    if (!comExamtype.Substring(0, 4).Equals("NCEE"))
                    {
                        string strsubj = "";
                        if(comAnsSheet == 3 )
                        {
                            var Xreg_no = OMRData.Substring(0, 10);
                            var XcandNo = "9999";
                            var Xsubj = scanSubjCode;
                            var XResponses = OMRData.Substring(10, OMRData.Length - 10);
                            var TempData = $"{Xreg_no}{XcandNo}{Xsubj}{XResponses}";
                            OMRData = TempData;
                            TempData = Xreg_no = XcandNo = Xsubj = XResponses = string.Empty;
                            var Subj = $"{Sosinp.get_FieldValueByName("Subj1")}{Sosinp.get_FieldValueByName("Subj2")}{Sosinp.get_FieldValueByName("Subj3")}";
                            if (comSubject == Subj)
                            {
                                strsubj = scanSubjCode;
                            }
                        }
                        else
                        {
                            strsubj = Sosinp.get_FieldValueByName("subject_code");
                        }
                        if (comAnsSheet == 0 )
                        {
                            string strexm = Sosinp?.get_FieldValueByName("Exm");
                            if (comSubject.Contains("l2"))
                            {
                                strexm = "1";
                            }
                            
                            sb.Capacity = 255;
                            ChkExm(sb, scanSubjCode, comExamtype.Substring(0, 4), strexm, sb.Capacity);
                            if (!String.IsNullOrEmpty(sb.ToString().Trim()))
                            {
                                ErrMsg = sb.ToString().Trim();// "Invalid Candidate Number:" + strCandidateNum;
                                DataValid = false;
                                sb.Remove(0, sb.Length);
                                return;
                            }
                        }



                        sb.Capacity = 255;
                        ValSubj(sb, sb.Capacity, strsubj, scanSubjCode, scanSubject, comExamtype.Substring(0, 4));
                        if (!String.IsNullOrEmpty(sb.ToString().Trim()))
                        {
                            ErrMsg = sb.ToString().Trim();// "Invalid Candidate Number:" + strCandidateNum;
                            DataValid = false;
                            sb.Remove(0, sb.Length);
                            return;
                        }
                    }
                    
                    /*if (comExamtype.Contains("SSCE"))//Validating SSCE Subject codes.
                    {
                        ErrorMsg = "";

                        string strexm = Sosinp.get_FieldValueByName("Exm");

                        if(string.IsNullOrEmpty(strexm))
                        {
                            ErrorMsg = "Invalid Examination " + strexm;
                            return;
                        }
                        
                        if (scanSubject == "Oral English")
                        {
                            ErrorMsg = ValidateOralEnglish(strsubj);
                            if (!String.IsNullOrEmpty(ErrorMsg))
                            {
                                ErrMsg = ErrorMsg;
                                DataValid = false;
                                return;
                            }
                        }*/

                        

                        /* if (scanSubject == "English Language")
                         {
                             ErrorMsg = "";
                             ErrorMsg = ValidateEnglishLanguage(strsubj);
                             if (!String.IsNullOrEmpty(ErrorMsg))
                             {
                                 ErrMsg = ErrorMsg;
                                 DataValid = false;
                                 return;
                             }
                         }*/
                        /*if (scanSubject != "Oral English")
                        {
                            ErrorMsg = "";
                            ErrorMsg = ValidateOtherSSCESubjects(strsubj, scanSubjCode);
                            if (!String.IsNullOrEmpty(ErrorMsg))
                            {
                                ErrMsg = ErrorMsg;
                                DataValid = false;
                                return;
                            }
                        }*/
                    //}
#endregion
#region Validating BECE Subject code
                    /*if (comExamtype.Contains("BECE"))//Validating BECE Subject codes.
                    {

                        if (!Tools.ReadRegistryValue("scanDir").Contains("Resit"))
                        {
                            string strexm = Sosinp.get_FieldValueByName("Exm");
                            if (string.IsNullOrEmpty(strexm))
                            {
                                ErrorMsg = "Invalid Examination " + strexm;
                                return;
                            }


                            string strsubj = Sosinp.get_FieldValueByName("subject_code");
                            

                            if (string.IsNullOrEmpty(strexm))
                            {
                                ErrorMsg = "Invalid Examination " + strexm;
                            }
                            ErrorMsg = "";
                            ErrorMsg = ValidateBECESubjects(strsubj, scanSubjCode);
                            if (!String.IsNullOrEmpty(ErrorMsg))
                            {
                                ErrMsg = ErrorMsg;
                                DataValid = false;
                                return;
                            }
                        }
                    }*/
                }
                #endregion

                #region EMS Validation
                
                if (comJob == "Essay" && scanType == "Blank")
                {
                    //EMS Ref No. Validation
                    if (string.IsNullOrEmpty(strCandidateNum))
                    {
                        ErrMsg = "INVALID Ref. No.: " + strCandidateNum;
                        DataValid = false;
                        return;
                    }
                }

                if (comJob == "Essay" && scanType!="Blank")
                {
                    StringBuilder sb = new StringBuilder(255);
                    ChkRefN(sb, strCandidateNum, sb.Capacity);
                    if (!String.IsNullOrEmpty(sb.ToString().Trim()))
                    {
                        ErrMsg = sb.ToString().Trim();// "Invalid Candidate Number:" + strCandidateNum;
                        DataValid = false;
                        sb.Remove(0, sb.Length);
                        return;
                    }

                    //EMS Ref No. Validation
                   /* if (string.IsNullOrEmpty(strCandidateNum))
                    {
                        ErrMsg = "INVALID Ref. No.: " + strCandidateNum;
                        DataValid = false;
                        return;
                    }*/

                    if(comExamtype.Contains("SSCE"))
                    {
                        //MessageBox.Show(OMRData);
                        var arrOMRData = OMRData.Split(' ');
                        string tempOMRData = "";
                        foreach (var s in arrOMRData)
                        {
                            var st = s;
                            if (s == "00")
                            {
                                st = "ZR";
                            }
                            tempOMRData += st + " ";

                        }
                        tempOMRData = tempOMRData.Remove(tempOMRData.LastIndexOf(' '), 1);
                        //MessageBox.Show(OMRData+"\n"+tempOMRData);
                        OMRData = tempOMRData;
                    }

                    if (comExamtype.Contains("BECE"))
                    {
                        //MessageBox.Show(OMRData);
                        var strOMRDataFieldName = "";
                        var arrOMRData = new string[21];// OMRData.Split(' ');
                        arrOMRData[0] = strCandidateNum;
                        int kountData = 1;
                        for (int i = 0; (i < 20); i++)
                        {
                            if (kountData < 10)
                            {
                                strOMRDataFieldName = "cand_0" + kountData.ToString();
                            }
                            else
                            {
                                strOMRDataFieldName = "cand_" + kountData.ToString();
                            }
                            arrOMRData[kountData] = Sosinp.get_FieldValueByName(strOMRDataFieldName);
                            //ScoresOnSheet[i] = strMark;
                            //MessageBox.Show(arrOMRData[kountData].ToString());
                            kountData++;
                        }
                        string tempOMRData = "";
                        foreach (var s in arrOMRData)
                        {
                            var st = s;
                            if (s == "00")
                            {
                                st = "ZR";
                            }
                            tempOMRData += st;

                        }
                        //tempOMRData = tempOMRData.Remove(tempOMRData.LastIndexOf(' '), 1);
                        OMRData = tempOMRData;
                        //MessageBox.Show(OMRData.ToString());
                    }

                   

                    /*for (int i = 0; i < strCandidateNum.Length; i++)
                    {
                        if (strCandidateNum.Substring(i,1)==">")
                        {
                            ErrMsg = "Multiple Shading: " + strCandidateNum;
                            DataValid = false;
                            return;
                        }
                    }

                    for (int i = 0; i < strCandidateNum.Length; i++)
                    {
                        if (strCandidateNum.Substring(i, 1) == " ")
                        {
                            ErrMsg = "Omitted Shading: " + strCandidateNum;
                            DataValid = false;
                            return;
                        }
                    }*/



                    /* string code = Tools.Right(strCandidateNum, 2);
                    // MessageBox.Show(strCandidateNum);
                     //MessageBox.Show(code);
                     string num1 = code.Substring(0, 1);
                     string num2 = code.Substring(1, 1);
                    // MessageBox.Show(num2);
                     int XX1 = Tools.getCandidatesOnSheet(num1);
                     int XX2= Tools.getCandidatesOnSheet(num2);

                     int CandidatesOnSheet = Convert.ToInt32(XX1.ToString()+XX2.ToString());

                     //MessageBox.Show(CandidatesOnSheet.ToString());


                     if (CandidatesOnSheet == 0 || CandidatesOnSheet > 20)
                     {
                         ErrMsg = "INVALID Check Digits! " + strCandidateNum;
                         DataValid = false;
                         return;
                     }*/

                     sb.Capacity = 255;
                    CndOnSht(strCandidateNum);
                    //MessageBox.Show(sb.ToString().Trim());
                    if (!String.IsNullOrEmpty(sb.ToString().Trim()))
                    {
                        ErrMsg = sb.ToString().Trim();// "Invalid Candidate Number:" + strCandidateNum;
                        DataValid = false;
                        sb.Remove(0, sb.Length);
                        return;
                    }

                    // string SystemubjectCode = Tools.ReadRegistryValue("BatchNumber").ToString().Substring(0, 2);
                    string SystemubjectCode ="";
                    string temsubjCode = Tools.ReadRegistryValue("SubjCode").ToString();
                    //if (temsubjCode == "8304")
                    //    SystemubjectCode = "31";// Tools.ReadRegistryValue("SubjCode").ToString().Substring(1, 2);
                   // else
                    //    SystemubjectCode = Tools.ReadRegistryValue("SubjCode");//.ToString().Substring(1,2);
                        

                    string SheetSubjectCode = "";// strCandidateNum.Substring(7, 2);
                    if (comExamtype.Contains("SSCE"))
                    {
                        SheetSubjectCode = strCandidateNum.Substring(6, 4);
                        SystemubjectCode = Tools.ReadRegistryValue("SubjCode").ToString();


                    }

                    if (comExamtype.Contains("BECE"))
                    {
                        SheetSubjectCode = strCandidateNum.Substring(10, 4);

                       /* if (temsubjCode == "8304")
                            SystemubjectCode = "31";
                        else*/
                           SystemubjectCode = Tools.ReadRegistryValue("SubjCode").ToString();
                    }
                        


                    //if (SheetSubjectCode.Substring(0, 1) == "0")
                    //   SheetSubjectCode = "9" + SheetSubjectCode.Substring(1, 1);
                    if (SystemubjectCode != SheetSubjectCode)
                       // MessageBox.Show(SystemubjectCode.ToString());
                       // MessageBox.Show(SheetSubjectCode.ToString());
                    {
                        ErrMsg = "INVALID Subject or Paper: " + SheetSubjectCode;
                        DataValid = false;
                        return;
                    }
                    string[] ScoresOnSheet = new string[20];
                    string[] FieldNames = new string[20];
                    string strFieldName, strMark;
                    int k = 1;
                    for (int i = 0; (i < 20); i++)
                    {
                        if (k < 10)
                        {
                            strFieldName = "cand_0" + k.ToString();
                        }
                        else
                        {
                            strFieldName = "cand_" + k.ToString();
                        }
                        strMark = Sosinp.get_FieldValueByName(strFieldName);
                        ScoresOnSheet[i] = strMark;
                        k++;
                    }

                    /*int k1 = 0;
                    for (int i = 0; (i < 20); i++)
                    {
                        k1++;
                        if (string.IsNullOrEmpty(ScoresOnSheet[i].Trim()) && k1 < CandidatesOnSheet)
                        {
                            ErrMsg = "Incomplete entry! " + " No score for Cand " + k1;
                            DataValid = false;
                            return;
                        }
                        
                    }
                    int k2 = 0;
                    for (int i = 0; (i < 20); i++)
                    {
                        k2++;
                        if (ScoresOnSheet[i].Substring(0, 1) == "N" && ScoresOnSheet[i].Substring(1, 1) != "S")
                        {
                            ErrMsg = "INVALID score! Cand " + k2;
                            DataValid = false;
                            return;
                        }
                        
                    }
                    int k3 = 0;
                    for (int i = 0; (i < 20); i++)
                    {
                        k3++;
                        if (ScoresOnSheet[i].Substring(0, 1) == "A" && ScoresOnSheet[i].Substring(1, 1) != "B")
                        {
                            ErrMsg = "INVALID score! Cand " + k3;
                            DataValid = false;
                            return;
                        }
                        
                    }

                    int k4 = 0;
                    for (int i = 0; (i < 20); i++)
                    {
                        k4++;
                        if (ScoresOnSheet[i].Substring(1, 1) == "B" && ScoresOnSheet[i].Substring(0, 1) != "A")
                        {
                            ErrMsg = "INVALID score! Cand " + k4;
                            DataValid = false;
                            return;
                        }
                        
                    }

                    int k5 = 0;
                    for (int i = 0; (i < 20); i++)
                    {
                        k5++;
                        if (ScoresOnSheet[i].Substring(1, 1) == "S" && ScoresOnSheet[i].Substring(0, 1) != "N")
                        {
                            ErrMsg = "INVALID score! Cand " + k5;
                            DataValid = false;
                            return;
                        }
                        
                    }*/



                    /*for (int i = 0; (i < 20); i++)
                    {
                        if (!string.IsNullOrEmpty(ScoresOnSheet[i].Trim()) && i > CandidatesOnSheet)
                        {
                            ErrMsg = "Candidate " + i + " Does not exist, " + CandidatesOnSheet + " Candidates records Expected!";
                            DataValid = false;
                            return;
                        }
                    }
                   // MessageBox.Show(CandidatesOnSheet.ToString());
                    int k6 = 0;
                    for (int i = 0; (i < CandidatesOnSheet); i++)
                    {
                       // MessageBox.Show(ScoresOnSheet[i].Trim());
                        k6++;
                        if (ScoresOnSheet[i].Trim().Length < 2)
                        {
                            ErrMsg = "Omitted Shading: " + " Cand " + k6 + ", " + ScoresOnSheet[i].ToString();
                            DataValid = false;
                            return;
                        }
                        
                    }

                    int k7 = 0;
                    for (int i = 0; (i < CandidatesOnSheet); i++)
                    {
                        k7++;
                        if (ScoresOnSheet[i].Trim().Contains(">") || ScoresOnSheet[i].Trim().Contains(">>"))
                        {
                            ErrMsg = "Multiple Shading: " + " Cand " + k7 + ", " + ScoresOnSheet[i].ToString();
                            DataValid = false;
                            
                            return;
                        }
                    }*/

                    sb.Capacity = 255;
                    ChkScrNSht(sb, ScoresOnSheet, strCandidateNum, sb.Capacity);
                    //MessageBox.Show(sb.ToString().Trim());
                    if (!String.IsNullOrEmpty(sb.ToString().Trim()))
                    {
                        ErrMsg = sb.ToString().Trim();// "Invalid Candidate Number:" + strCandidateNum;
                        DataValid = false;
                        sb.Remove(0, sb.Length);
                        return;
                    }

                }

            }
#endregion
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred : " + ex.Message + "\n" + "Form will be rejected");
                ErrMsg = "Error " + ex.Message;
                DataValid = false;
            }

        }


       /* protected override void IdentifySheet(ref int FormIndex, string OMRData)
        {
            base.IdentifySheet(ref FormIndex, OMRData);
            string CandidateNo;
            string SubjectCode;
            MessageBox.Show(OMRData);
            CandidateNo = OMRData.Substring(0, 10);
            SubjectCode = OMRData.Substring(10, 4);
            MessageBox.Show(CandidateNo);
            //if (CandidateNo.Length == 10)
            //{
                for(int i=0; i< Sosinp.FormTypeCount-1; i++)
                {
                    MessageBox.Show(Sosinp.FormTypes[i].FormName);
                    if(Sosinp.FormTypes[i].FormName=="SSCESCAN" || Sosinp.FormTypes[i].FormName == "BECESCAN")
                    {
                        FormIndex = i;
                        break;
                    }
                }
            //}
        }*/

       
        protected override void StartBatch(ref bool AllowStart, ref int BatchNumber, ref string OperatorName, ref string TrackingId, ref int ExpectSheets, ref string Comment)
        {
            base.StartBatch(ref AllowStart, ref BatchNumber, ref OperatorName, ref TrackingId, ref ExpectSheets, ref Comment);
            try
            {
                Sosinp.UsesOnIdentify = true;
                Sosinp.StatusBarMessage = Tools.ReadRegistryValue("UID").ToString();
                Sosinp.StopWhenReached = true;

                comExamtype = Tools.ReadRegistryValue("ExamType").ToString();
                comSubject = Tools.ReadRegistryValue("shortsubj").ToString();
                comState = Tools.ReadRegistryValue("state").ToString();
                comUserName = Tools.ReadRegistryValue("UID").ToString();
                comJob = Tools.ReadRegistryValue("Job").ToString();
                scanSubject = Tools.ReadRegistryValue("subject");
                scanSubjCode = Tools.ReadRegistryValue("subjcode");
                comStartBatch = Convert.ToBoolean(Tools.ReadRegistryValue("startbatch").ToString());
                comLogOut = Convert.ToBoolean(Tools.ReadRegistryValue("Logout").ToString());
                comExpSheets = Convert.ToInt32(Tools.ReadRegistryValue("ExpectedSheets"));
                scanType = Tools.ReadRegistryValue("scanType");
                ExpectSheets = comExpSheets;
                OperatorName = comUserName;
                Tools.WriteRegistryValue("startbatch", "false");
                SubjCodeValidCount = 0;
                SubjCodeValidCountCheck = 20;
                MaxSubjectCountAllowed = 1;
                SheetPassCount = 0;
                comAnsSheet = Convert.ToInt32(Tools.ReadRegistryValue("AnswerSheet").ToString());
                if (comUserName.ToLower() == "default")
                {
                    AllowStart = false;
                    Tools.WriteRegistryValue("startbatch", "false");
                    Tools.WriteRegistryValue("Loguot", "true");
                    Sosinp.Terminate(0, -1);
                }
                else
                {

                    if (comLogOut == false)
                    {
                        if (comStartBatch == false)
                        {
                            if (comExamtype.Substring(0, 4) == "SSCE" || comExamtype.Substring(0, 4) == "BECE")
                            {
                                msg = "Do you want to Continue Scanning " + comSubject + "  and " + comState;

                            }
                            else
                            {
                                msg = "Do you want to Continue Scanning " + comSubject;
                            }
                            if (MessageBox.Show(msg, "Companion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            {
                                AllowStart = true;

                            }
                            else
                            {
                                AllowStart = false;
                                Process.Start(@"c:\program files\necoscan\scan\scan.exe");
                                Sosinp.Terminate(0, -1);
                            }
                        }
                    }
                    else
                    {
                        AllowStart = false;
                        LogoutCurrentUser();

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred : " + ex.Message + "\n" + "Cannot start this batch");
                AllowStart = false;
                Sosinp.Terminate(0, -1);
            }
            ExpectSheets = comExpSheets;
        }

        protected override void EndBatch(bool Discard)
        {
            base.EndBatch(Discard);
            try
            {
               //string l=base.Sosinp.Statistics.Total.ToString();  
                if (Tools.ReadRegistryValue("BatchNumber").ToString() != "000000.000")
                {
                    //frmDialogue f = new frmDialogue(Tools.ReadRegistryValue("BatchNumber").ToString());
                    //f.Show();
                    
                    MessageBox.Show(Tools.ReadRegistryValue("BatchNumber").ToString(), "Scan file", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (MessageBox.Show("Do you want to sign out now", "Log out", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    if (MessageBox.Show("Loging out current user \n" + " Click OK to end session ", "Loging out", MessageBoxButtons.OKCancel, MessageBoxIcon.Information ) == DialogResult.OK)
                    {
                        Tools.WriteRegistryValue("UID", "default");
                        Tools.WriteRegistryValue("Logout", "true");
                        Sosinp.Terminate(0, -1);
                        AllowLogout = true;
                    }
                    else
                    {
                        AllowLogout = false;
                    }
                }
                else
                {
                    if (Tools.Right(Tools.ReadRegistryValue("BatchNumber").ToString(), 3) == "999")
                    {
                        MessageBox.Show("Batch Counter has exceeded allowed figure for this subject (999)\n" +
                                    "please use another system to continue scanning the state and subject");
                        if (MessageBox.Show("Logging out current user \n" + " Click OK to end session ", "Loging out", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                        {
                            Tools.WriteRegistryValue("UID", "default");
                            Tools.WriteRegistryValue("Logout", "true");
                            Sosinp.Terminate(0, -1);
                            
                        }
                    }
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred : " + ex.Message + "\n" + "Cannot end batch");
                Sosinp.Terminate(0, -1);
            }
        }
        bool AllowLogout = false;
        protected override void CloseDown()
        {
            base.CloseDown();
            try
            {
                //ManipulateTextFiles.WriteToTextFile(@"c:\program files\necoscan\scan\scan.ref", "LogOut", "true");
                //ManipulateTextFiles.WriteToTextFile(@"c:\program files\necoscan\scan\scan.ref", "Uid", "Default");
                //Tools.WriteRegistryValue("UID", "default");
                //Tools.WriteRegistryValue("Logout", "true");
                //Sosinp.Terminate(0, -1);
                // if (MessageBox.Show("Do you want to sign out now", "Log out", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                // {
                if (AllowLogout == true)
                {
                    Tools.WriteRegistryValue("UID", "default");
                    Tools.WriteRegistryValue("Logout", "true");
                    Tools.WriteRegistryValue("OperatorId", "");
                }
                    
               // }

            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred : " + ex.Message + "\n" + "Closing down program");
                Sosinp.Terminate (0, -1);
            }
        }
        internal string CalculateCheckDigits(string number)
        {
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
            string digit1 = Tools.Chr(Convert.ToInt32(((97 - sumdig) / 10)) + 65);
            string digit2 = Tools.Chr(((97 - sumdig) % 10) + 65);
            return digit1+digit2;

        }

        internal void LogoutCurrentUser()
        {
            //MessageBox.Show("Logging out, click OK to terminate program...","Terminate SosInp");
            Sosinp.Terminate(0, -1);
        }

        internal string ValidateOralEnglish(string code)
        {
            string ErrorMsg = "";
            if (SubjCodeValidCount < SubjCodeValidCountCheck)
            {
                if (code == "1014")
                {
                    SubjCodeValidCount = SubjCodeValidCount + MaxSubjectCountAllowed;//increment on every valid subject code encountered
                }
                else
                {
                    
                    ErrorMsg = "Invalid Subject code: " + code + ": - 1014 Expected";
                    SubjCodeValidCount = SubjCodeValidCount - 1; //subtract on every wrong subject code encountered
                }
            }
            else
            {
                SheetPassCount = SheetPassCount + 1;
                if (SheetPassCount == 20)
                {
                    SubjCodeValidCount = 1;
                    SubjCodeValidCountCheck = 5;
                    SheetPassCount = 0;
                }
            }
            return ErrorMsg;
        }

        internal string ValidateEnglishLanguage(string code)
        {
            string ErrorMsg = "";
            //if (ScanSubjectCode.Substring(0, 3) != OMRSubjectCode.Substring(0, 3))
            //{
            //    ErrorMsg = "Invalid Subject code: " + code + ": - 1013 Expected";
                //SubjCodeValidCount = SubjCodeValidCount - 1; //subtract on every wrong subject code encountered
           // }
            /*if (SubjCodeValidCount < SubjCodeValidCountCheck)
            {
                if (code == "1011" || code == "1012" || code=="1013")
                {
                    SubjCodeValidCount = SubjCodeValidCount + MaxSubjectCountAllowed;//increment on every valid subject code encountered
                }
                else
                {
                    
                    ErrorMsg = "Invalid Subject code: " + code + ": - 1011, 1012 or 1013 Expected";
                    SubjCodeValidCount = SubjCodeValidCount - 1; //subtract on every wrong subject code encountered
                }
            }
            else
            {
                SheetPassCount = SheetPassCount + 1;
                if (SheetPassCount == 20)
                {
                    SubjCodeValidCount = 1;
                    SubjCodeValidCountCheck = 5;
                    SheetPassCount = 0;
                }
            }*/
            return ErrorMsg;
        }

        internal string ValidateOtherSSCESubjects(string OMRSubjectCode, string ScanSubjectCode)
        {
            string ErrorMsg = "";
            if (ScanSubjectCode.Substring(0, 3) != OMRSubjectCode.Substring(0, 3))
            {
                ErrorMsg = "Invalid Subject code: " + OMRSubjectCode + ": - " + ScanSubjectCode + " Expected";
                SubjCodeValidCount = SubjCodeValidCount - 1; //subtract on every wrong subject code encountered
            }
            /* if (SubjCodeValidCount < SubjCodeValidCountCheck)
             {
                 if (ScanSubjectCode.Substring(0, 3) != OMRSubjectCode.Substring(0, 3))
                 {
                     ErrorMsg = "Invalid Subject code: " + OMRSubjectCode +": - "+ScanSubjectCode + " Expected";
                     SubjCodeValidCount = SubjCodeValidCount - 1; //subtract on every wrong subject code encountered
                 }
                 else
                 {
                     SubjCodeValidCount = SubjCodeValidCount + MaxSubjectCountAllowed;//increment on every valid subject code encountered
                 }
             }
             else
             {
                 SheetPassCount = SheetPassCount + 1;
                 if (SheetPassCount == 20)
                 {
                     SubjCodeValidCount = 1;
                     SubjCodeValidCountCheck = 5;
                     SheetPassCount = 0;
                 }
             }*/
            return ErrorMsg;
        }

        internal string ValidateBECESubjects(string OMRSubjectCode, string ScanSubjectCode)
        {
            string ErrorMsg = "";
            if (ScanSubjectCode.Trim() != OMRSubjectCode.Trim())
            {
                ErrorMsg = "Invalid Subject code: " + OMRSubjectCode + ": - " + ScanSubjectCode + " Expected";
                SubjCodeValidCount = SubjCodeValidCount - 1; //subtract on every wrong subject code encountered
            }
            /*if (SubjCodeValidCount < SubjCodeValidCountCheck)
            {
                if (ScanSubjectCode.Trim() != OMRSubjectCode.Trim())
                {
                    ErrorMsg = "Invalid Subject code: " + OMRSubjectCode + ": - " + ScanSubjectCode + " Expected";
                    SubjCodeValidCount = SubjCodeValidCount - 1; //subtract on every wrong subject code encountered
                }
                else
                {
                    SubjCodeValidCount = SubjCodeValidCount + MaxSubjectCountAllowed;//increment on every valid subject code encountered
                }
            }
            else
            {
                SheetPassCount = SheetPassCount + 1;
                if (SheetPassCount == 7)
                {
                    SubjCodeValidCount = 1;
                    SubjCodeValidCountCheck = 3;
                    SheetPassCount = 0;
                    MaxSubjectCountAllowed = 2;
                }
            }*/
            return ErrorMsg;
        }
    }
}
