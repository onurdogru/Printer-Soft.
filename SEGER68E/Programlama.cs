using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using System.IO;

namespace SEGER68E
{
    public class ProgramlamaForm
    {
        public ProgAyarForm ProgAyarFrm;
        public Main MainFrm;
        public ProcessForm ProcessFrm;

        /*PROGRAMLAMA + BARKOD DEĞİŞKENLERİ*/
        string[] batchFileFeedback = new string[4];

        int versions_number = 0;
        string[] versionsBarcodName = new string[21];
        string[] versionsSlaveName = new string[21];
        int[] stepProgJob = new int[21];
        string computerBatchFileAdress;
        int isProgrammingState = 0;

        string companyNo;
        string productDate;
        string indexNo;
        string productNo;
        string cardType;
        string gerberVer;
        string bomVer;
        string ictRev;
        string fctRev;
        string softwareVer;
        string softwareRev;
        string uniqId;

        public void programlamaInit()
        {
            computerBatchFileAdress = Prog_Ayarlar.Default.Logdosyayolu;

            companyNo = Prog_Ayarlar.Default.companyNo;
            cardType = Prog_Ayarlar.Default.cardNo;
            gerberVer = Prog_Ayarlar.Default.gerberVer;
            bomVer = Prog_Ayarlar.Default.BOMVer;
            ictRev = Prog_Ayarlar.Default.ICTRev;
            fctRev = Prog_Ayarlar.Default.FCTRev;
            softwareVer = Prog_Ayarlar.Default.softwareVer;
            softwareRev = Prog_Ayarlar.Default.softwareRev;

            this.versions_number = Convert.ToInt16(Prog_Ayarlar.Default.barcodeNum);
            this.stepProgJob[1] = Convert.ToInt16(Prog_Ayarlar.Default.step1Job);
            this.stepProgJob[2] = Convert.ToInt16(Prog_Ayarlar.Default.step2Job);
            this.stepProgJob[3] = Convert.ToInt16(Prog_Ayarlar.Default.step3Job);
            this.stepProgJob[4] = Convert.ToInt16(Prog_Ayarlar.Default.step4Job);
            this.stepProgJob[5] = Convert.ToInt16(Prog_Ayarlar.Default.step5Job);
            this.stepProgJob[6] = Convert.ToInt16(Prog_Ayarlar.Default.step6Job);
            this.stepProgJob[7] = Convert.ToInt16(Prog_Ayarlar.Default.step7Job);
            this.stepProgJob[8] = Convert.ToInt16(Prog_Ayarlar.Default.step8Job);
            this.stepProgJob[9] = Convert.ToInt16(Prog_Ayarlar.Default.step9Job);
            this.stepProgJob[10] = Convert.ToInt16(Prog_Ayarlar.Default.step10Job);
            this.versionsBarcodName[1] = Prog_Ayarlar.Default.barcode1;
            this.versionsBarcodName[2] = Prog_Ayarlar.Default.barcode2;
            this.versionsBarcodName[3] = Prog_Ayarlar.Default.barcode3;
            this.versionsBarcodName[4] = Prog_Ayarlar.Default.barcode4;
            this.versionsBarcodName[5] = Prog_Ayarlar.Default.barcode5;
            this.versionsBarcodName[6] = Prog_Ayarlar.Default.barcode6;
            this.versionsBarcodName[7] = Prog_Ayarlar.Default.barcode7;
            this.versionsBarcodName[8] = Prog_Ayarlar.Default.barcode8;
            this.versionsBarcodName[9] = Prog_Ayarlar.Default.barcode9;
            this.versionsBarcodName[10] = Prog_Ayarlar.Default.barcode10;
            this.versionsSlaveName[1] = Prog_Ayarlar.Default.Sbarcode1;
            this.versionsSlaveName[2] = Prog_Ayarlar.Default.Sbarcode2;
            this.versionsSlaveName[3] = Prog_Ayarlar.Default.Sbarcode3;
            this.versionsSlaveName[4] = Prog_Ayarlar.Default.Sbarcode4;
            this.versionsSlaveName[5] = Prog_Ayarlar.Default.Sbarcode5;
            this.versionsSlaveName[6] = Prog_Ayarlar.Default.Sbarcode6;
            this.versionsSlaveName[7] = Prog_Ayarlar.Default.Sbarcode7;
            this.versionsSlaveName[8] = Prog_Ayarlar.Default.Sbarcode8;
            this.versionsSlaveName[9] = Prog_Ayarlar.Default.Sbarcode9;
            this.versionsSlaveName[10] = Prog_Ayarlar.Default.Sbarcode10;

            batchFileFeedback[0] = Prog_Ayarlar.Default.successBatch;
            batchFileFeedback[1] = Prog_Ayarlar.Default.error1Batch;
            batchFileFeedback[2] = Prog_Ayarlar.Default.error2Batch;
            batchFileFeedback[3] = Prog_Ayarlar.Default.error3Batch;
        }

        public bool BarcodeControl(object data)
        {
            bool result = false;
            this.MainFrm.ConsoleClean();

            string barcode = (string)data;
            string company_no = string.Empty;
            string sap_no = string.Empty;
            string product_date = string.Empty;
            string index_no = string.Empty;

            // if barcode is true
            if (BarcodeCheck(barcode))
            {
                this.MainFrm.ConsoleAppendLine("Barkod: " + barcode, Color.Black);
                this.MainFrm.ConsoleNewLine();
                this.MainFrm.ConsoleAppendLine("Barkod Kriterlere Uygun.", Color.Green);
                this.MainFrm.ConsoleNewLine();

                //barcode = "91358531010045201234560000000123456712070101010104";  //Örnek
                company_no = barcode.Substring(0, 6);
                this.MainFrm.sap_no[this.MainFrm.barcodeCounter] = barcode.Substring(6, 6);
                product_date = barcode.Substring(14, 8);
                index_no = barcode.Substring(23, 6);

                this.MainFrm.ConsoleAppendLine("Firma Kodu: " + company_no, Color.Black);
                this.MainFrm.ConsoleAppendLine("SAP Kodu: " + this.MainFrm.sap_no[this.MainFrm.barcodeCounter], Color.Black);
                this.MainFrm.ConsoleAppendLine("Üretim Tarihi: " + product_date, Color.Black);
                this.MainFrm.ConsoleAppendLine("İndex No: " + index_no, Color.Black);
                this.MainFrm.ConsoleNewLine();
                if (this.MainFrm.urunlerRead(barcode))
                {
                    this.MainFrm.urun_id[this.MainFrm.barcodeCounter] = this.MainFrm.urun_id_carrier;
                    result = true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                if (barcode.Equals(""))
                {
                    this.MainFrm.ConsoleAppendLine("Barkod: " + barcode, Color.Black);
                    this.MainFrm.ConsoleNewLine();
                    this.MainFrm.ConsoleAppendLine("Barkod Boş Bırakılamaz! FCT yapılamadı!", Color.Red);
                    this.MainFrm.ConsoleNewLine();
                }
                else
                {
                    this.MainFrm.ConsoleAppendLine("Barkod: " + barcode, Color.Black);
                    this.MainFrm.ConsoleNewLine();
                    this.MainFrm.ConsoleAppendLine("Yanlış barkod! Barkod kriterlere uygun değil! FCT yapılamadı!", Color.Red);
                    this.MainFrm.ConsoleNewLine();
                }
                result = false;
            }

            // Assign Last Barcode with Current
            this.MainFrm.tbBarcodeLast.Invoke(new Action(delegate () { this.MainFrm.tbBarcodeLast.Text = barcode; }));

            this.MainFrm.tbBarcodeCurrent.Invoke(new Action(delegate ()
            {
                // Clean Current Barcode
                this.MainFrm.tbBarcodeCurrent.Text = "";
                // Focus Barcode textbox
                this.MainFrm.tbBarcodeCurrent.Focus();
                // Select all text in textbox
                this.MainFrm.tbBarcodeCurrent.SelectionStart = 0;
                this.MainFrm.tbBarcodeCurrent.SelectionLength = this.MainFrm.tbBarcodeCurrent.Text.Length;
            }));
            return result;
        }

        /* Gelen Barcode Kontrol Edilir*/
        private bool BarcodeCheck(String barcode)
        {
            // Barcode length should be between 26-150.
            if (barcode.Length > 29) return false;

            string company_no = barcode.Substring(0, 6);
            string sap_no = barcode.Substring(6, 6);

            // check company no
            if (!company_no.Equals(companyNo))
            {
                this.MainFrm.ConsoleAppendLine("Firma Kodu !", Color.Red);
                return false;
            }
            /*
            // check sap no
            if (sap_no.IndexOf(sap_no) != 0)
            {
                this.MainFrm.ConsoleAppendLine("SAP Kodu !", Color.Red);
                return false;
            }*/


            // There is no problem
            return true;
        }

        public bool ProgramProduct(string product_no)
        {
            bool result = false;
            String batchPath = String.Empty;

            for (int i = 1; i <= versions_number; i++)
            {
                if (versionsBarcodName[i] == product_no)
                {
                    if (stepProgJob[versions_number] == 1)
                    {
                        batchPath = computerBatchFileAdress + product_no + ".bat";    // C:\Users\serkan.baki\Desktop\MIND-BATCH-FILES\
                        result = RunBatch(batchPath, this.MainFrm.projectNameTxt.Text);
                        if (result == false) return result;
                    }
                    break;
                }
            }
            return result;
        }

        /*Seçilen .bat Çalıştırılır- Kontrol Edilir ve Kapatılır*/
        private bool RunBatch(string batch_path, string batch_name)
        {
            bool result = false;

            Process processBatch = new Process();
            processBatch.StartInfo.UseShellExecute = false;
            processBatch.StartInfo.RedirectStandardOutput = true;
            processBatch.StartInfo.CreateNoWindow = true;
            processBatch.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            processBatch.StartInfo.FileName = batch_path;
            //processBatch.StartInfo.Arguments = string.Format("");
            processBatch.Start();

            StreamReader strmReader = processBatch.StandardOutput;
            string batchTempRow = string.Empty;
            // get all lines of batch
            while ((batchTempRow = strmReader.ReadLine()) != null)
            {
                // Write batch operation to the console
                this.MainFrm.ConsoleAppendLine(batchTempRow, Color.Black);

                // check programming is successful.
                // if succesfully finished.
                if (Prog_Ayarlar.Default.chBoxSuccess && ((batchTempRow.IndexOf("pause", StringComparison.OrdinalIgnoreCase) >= 0) || (batchTempRow.IndexOf(batchFileFeedback[0], StringComparison.OrdinalIgnoreCase) >= 0)))  // color ae
                {
                    isProgrammingState = 2;
                    this.MainFrm.ConsoleNewLine();
                    this.MainFrm.ConsoleAppendLine(batch_name + " Programlama İşlemi Başarıyla Tamamlanmıştır!", Color.Green);
                    result = true;
                    break;
                }
                else if (Prog_Ayarlar.Default.chBoxError1 && ((batchTempRow.IndexOf("pause", StringComparison.OrdinalIgnoreCase) >= 0) || (batchTempRow.IndexOf(batchFileFeedback[1], StringComparison.OrdinalIgnoreCase) >= 0))) //Could not start CPU core.
                {
                    isProgrammingState = 1;
                    this.MainFrm.ConsoleNewLine();
                    this.MainFrm.ConsoleAppendLine(batch_name + " Programlama İşlemi Başarısız1.", Color.Red);  // Programlama Soketi Düzgün Takılı Değil!
                    result = false;
                    break;
                }
                else if (Prog_Ayarlar.Default.chBoxError2 && ((batchTempRow.IndexOf("pause", StringComparison.OrdinalIgnoreCase) >= 0) || (batchTempRow.IndexOf(batchFileFeedback[2], StringComparison.OrdinalIgnoreCase) >= 0)))  // Cannot connect to target.
                {
                    isProgrammingState = 1;
                    this.MainFrm.ConsoleNewLine();
                    this.MainFrm.ConsoleAppendLine(batch_name + " Programlama İşlemi Başarısız2.", Color.Red); // Programlama Soketi Takılı Değil!
                    result = false;
                    break;
                }
                else if (Prog_Ayarlar.Default.chBoxError3 && ((batchTempRow.IndexOf("pause", StringComparison.OrdinalIgnoreCase) >= 0) || (batchTempRow.IndexOf(batchFileFeedback[3], StringComparison.OrdinalIgnoreCase) >= 0))) //FAILED
                {
                    isProgrammingState = 1;
                    this.MainFrm.ConsoleNewLine();
                    this.MainFrm.ConsoleAppendLine(batch_name + " Programlama İşlemi Başarısız3.", Color.Red);  //  USB Takılı Değil!
                    result = false;
                    break;
                }
            }

            // if batch didn't closed kill it.
            if (!processBatch.HasExited)
            {
                processBatch.Kill();
            }

            return result;
        }

        /*Gelen Barkodu Ayıklanır*/
        private String GetStringBetweenTwoSubStrings(String source, String sub1, String sub2)
        {
            int pFrom = source.IndexOf(sub1) + sub1.Length;
            int pTo = source.LastIndexOf(sub2);
            String result = source.Substring(pFrom, pTo - pFrom);
            return result;
        }
    }
}
