// Decompiled with JetBrains decompiler
// Type: EsdTurnikesi.AyarForm
// Assembly: EsdTurnikesi, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: C8099926-BBEB-495E-ADF6-36B4F5F75BE8
// Assembly location: C:\Users\serkan.baki\Desktop\esd-rar\ESD\Release\EsdTurnikesi.exe

using System;
using System.ComponentModel;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Text;

namespace SEGER68E
{
    public class ProgAyarForm : Form
    {
        public Form MainFrm;
        private IContainer components;

        private Button btnKaydet;
        private ComboBox barcodeNum;
        private TextBox barcode1;
        private TextBox barcode2;
        private TextBox barcode4;
        private TextBox barcode3;
        private TextBox barcode8;
        private TextBox barcode7;
        private TextBox barcode6;
        private TextBox barcode5;
        private TextBox barcode10;
        private TextBox barcode9;
        private Label label12;
        private TextBox txtBatchDosya;
        private Button btnLogsec;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private Label label6;
        private Label label7;
        private Label label8;
        private Label label9;
        private Label label10;
        private Label label11;
        private TextBox companyNo;
        private Label label23;
        private TextBox SAPNo;
        private Label label28;
        private ToolTip toolTip1;
        private Button btnINIsec;
        private TextBox txtINIdosya;
        private Label label220;
        private Button btnOkuIni;
        private Button btnKaydetIni;
        private GroupBox groupBox2;
        private GroupBox groupBox3;
        private PictureBox infoPicture2;
        private ComboBox step10Job;
        private Label label68;
        private ComboBox step9Job;
        private Label label69;
        private ComboBox step8Job;
        private Label label70;
        private ComboBox step7Job;
        private Label label71;
        private ComboBox step6Job;
        private Label label72;
        private ComboBox step5Job;
        private Label label37;
        private ComboBox step4Job;
        private Label label38;
        private ComboBox step3Job;
        private Label label39;
        private ComboBox step2Job;
        private Label label40;
        private ComboBox step1Job;
        private Label label41;
        private ToolTip toolTip2;
        private Label label52;
        private Label label53;
        private Label label54;
        private Label label55;
        private Label label56;
        private Label label57;
        private Label label58;
        private Label label59;
        private Label label60;
        private Label label61;
        private TextBox Sbarcode10;
        private TextBox Sbarcode9;
        private TextBox Sbarcode8;
        private TextBox Sbarcode7;
        private TextBox Sbarcode6;
        private TextBox Sbarcode5;
        private TextBox Sbarcode4;
        private TextBox Sbarcode3;
        private TextBox Sbarcode2;
        private TextBox Sbarcode1;
        private PictureBox infoPicture3;
        private PictureBox infoPicture4;
        private ToolTip toolTip3;
        private ToolTip toolTip4;
        private CheckBox chBoxSuccess;
        private CheckBox chBoxError1;
        private CheckBox chBoxError2;
        private CheckBox chBoxError3;
        private GroupBox groupBox4;
        private TextBox txtError3Batch;
        private TextBox txtError2Batch;
        private TextBox txtError1Batch;
        private GroupBox groupBox1;
        private TextBox softwareRev;
        private Label label31;
        private Label label29;
        private TextBox FCTRev;
        private TextBox softwareVer;
        private Label label30;
        private Label label26;
        private TextBox BOMVer;
        private TextBox ICTRev;
        private Label label27;
        private Label label24;
        private TextBox cardNo;
        private TextBox gerberVer;
        private Label label25;
        private Label label13;
        private TextBox kDutyCycleMin;
        private GroupBox groupBox5;
        private TextBox kFrekansMin;
        private Label label15;
        private TextBox kFrekansMax;
        private Label label16;
        private Label label14;
        private TextBox kDutyCycleMax;
        private TextBox kTP12Min;
        private Label label17;
        private TextBox kTP12Max;
        private Label label18;
        private GroupBox groupBox6;
        private TextBox iTP12Min;
        private Label label19;
        private TextBox iTP12Max;
        private Label label20;
        private TextBox iFrekansMin;
        private Label label21;
        private TextBox iDutyCycleMin;
        private TextBox iFrekansMax;
        private Label label22;
        private Label label33;
        private Label label34;
        private TextBox iDutyCycleMax;
        private TextBox iTP13Min;
        private Label label32;
        private TextBox iTP13Max;
        private Label label35;
        private TextBox kTP13Min;
        private Label label36;
        private TextBox kTP13Max;
        private Label label42;
        private Label label65;
        private Label label66;
        private Label label67;
        private Label label73;
        private Label label51;
        private Label label62;
        private Label label63;
        private Label label64;
        private Label label49;
        private Label label50;
        private Label label48;
        private Label label47;
        private Label label46;
        private Label label45;
        private Label label44;
        private Label label43;
        private TextBox txtSuccessBatch;

        public ProgAyarForm()
        {
            this.InitializeComponent();
        }

        public class INIKaydet
        {
            [DllImport("kernel32")]
            private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

            [DllImport("kernel32")]
            private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

            public INIKaydet(string dosyaYolu)
            {
                DOSYAYOLU = dosyaYolu;
            }
            private string DOSYAYOLU = String.Empty;
            public string Varsayilan { get; set; }
            public string Oku(string bolum, string ayaradi)
            {
                Varsayilan = Varsayilan ?? string.Empty;
                StringBuilder StrBuild = new StringBuilder(256);
                GetPrivateProfileString(bolum, ayaradi, Varsayilan, StrBuild, 255, DOSYAYOLU);
                return StrBuild.ToString();
            }
            public long Yaz(string bolum, string ayaradi, string deger)
            {
                return WritePrivateProfileString(bolum, ayaradi, deger, DOSYAYOLU);
            }
        }

        private void AyarForm_Load(object sender, EventArgs e)
        {
            toolTip_Load();

            this.barcodeNum.Text = Prog_Ayarlar.Default.barcodeNum;

            this.step1Job.Text = Prog_Ayarlar.Default.step1Job;
            this.step2Job.Text = Prog_Ayarlar.Default.step2Job;
            this.step3Job.Text = Prog_Ayarlar.Default.step3Job;
            this.step4Job.Text = Prog_Ayarlar.Default.step4Job;
            this.step5Job.Text = Prog_Ayarlar.Default.step5Job;
            this.step6Job.Text = Prog_Ayarlar.Default.step6Job;
            this.step7Job.Text = Prog_Ayarlar.Default.step7Job;
            this.step8Job.Text = Prog_Ayarlar.Default.step8Job;
            this.step9Job.Text = Prog_Ayarlar.Default.step9Job;
            this.step10Job.Text = Prog_Ayarlar.Default.step10Job;

            this.barcode1.Text = Prog_Ayarlar.Default.barcode1;
            this.barcode2.Text = Prog_Ayarlar.Default.barcode2;
            this.barcode3.Text = Prog_Ayarlar.Default.barcode3;
            this.barcode4.Text = Prog_Ayarlar.Default.barcode4;
            this.barcode5.Text = Prog_Ayarlar.Default.barcode5;
            this.barcode6.Text = Prog_Ayarlar.Default.barcode6;
            this.barcode7.Text = Prog_Ayarlar.Default.barcode7;
            this.barcode8.Text = Prog_Ayarlar.Default.barcode8;
            this.barcode9.Text = Prog_Ayarlar.Default.barcode9;
            this.barcode10.Text = Prog_Ayarlar.Default.barcode10;

            this.Sbarcode1.Text = Prog_Ayarlar.Default.Sbarcode1;
            this.Sbarcode2.Text = Prog_Ayarlar.Default.Sbarcode2;
            this.Sbarcode3.Text = Prog_Ayarlar.Default.Sbarcode3;
            this.Sbarcode4.Text = Prog_Ayarlar.Default.Sbarcode4;
            this.Sbarcode5.Text = Prog_Ayarlar.Default.Sbarcode5;
            this.Sbarcode6.Text = Prog_Ayarlar.Default.Sbarcode6;
            this.Sbarcode7.Text = Prog_Ayarlar.Default.Sbarcode7;
            this.Sbarcode8.Text = Prog_Ayarlar.Default.Sbarcode8;
            this.Sbarcode9.Text = Prog_Ayarlar.Default.Sbarcode9;
            this.Sbarcode10.Text = Prog_Ayarlar.Default.Sbarcode10;

            this.txtBatchDosya.Text = Prog_Ayarlar.Default.Logdosyayolu;
            this.companyNo.Text = Prog_Ayarlar.Default.companyNo;
            this.SAPNo.Text = Prog_Ayarlar.Default.SAPNo;
            this.cardNo.Text = Prog_Ayarlar.Default.cardNo;
            this.gerberVer.Text = Prog_Ayarlar.Default.gerberVer;
            this.BOMVer.Text = Prog_Ayarlar.Default.BOMVer;
            this.ICTRev.Text = Prog_Ayarlar.Default.ICTRev;
            this.FCTRev.Text = Prog_Ayarlar.Default.FCTRev;
            this.softwareVer.Text = Prog_Ayarlar.Default.softwareVer;
            this.softwareRev.Text = Prog_Ayarlar.Default.softwareRev;

            this.txtINIdosya.Text = Prog_Ayarlar.Default.iniDosyaYolu;

            this.chBoxSuccess.Checked = Prog_Ayarlar.Default.chBoxSuccess;
            if (chBoxSuccess.Checked == true)
                txtSuccessBatch.Enabled = true;
            else
                txtSuccessBatch.Enabled = false;

            this.chBoxError1.Checked = Prog_Ayarlar.Default.chBoxError1;
            if (chBoxError1.Checked == true)
                txtError1Batch.Enabled = true;
            else
                txtError1Batch.Enabled = false;

            this.chBoxError2.Checked = Prog_Ayarlar.Default.chBoxError2;
            if (chBoxError2.Checked == true)
                txtError2Batch.Enabled = true;
            else
                txtError2Batch.Enabled = false;

            this.chBoxError3.Checked = Prog_Ayarlar.Default.chBoxError3;
            if (chBoxError3.Checked == true)
                txtError3Batch.Enabled = true;
            else
                txtError3Batch.Enabled = false;

            this.txtSuccessBatch.Text = Prog_Ayarlar.Default.successBatch;
            this.txtError1Batch.Text = Prog_Ayarlar.Default.error1Batch;
            this.txtError2Batch.Text = Prog_Ayarlar.Default.error2Batch;
            this.txtError3Batch.Text = Prog_Ayarlar.Default.error3Batch;

            this.kDutyCycleMin.Text = Prog_Ayarlar.Default.kDutyCycleMin;
            this.kFrekansMin.Text = Prog_Ayarlar.Default.kFrekansMin;
            this.kTP12Min.Text = Prog_Ayarlar.Default.kTP12Min;
            this.kTP13Min.Text = Prog_Ayarlar.Default.kTP13Min;
            this.kDutyCycleMax.Text = Prog_Ayarlar.Default.kDutyCycleMax;
            this.kFrekansMax.Text = Prog_Ayarlar.Default.kFrekansMax;
            this.kTP12Max.Text = Prog_Ayarlar.Default.kTP12Max;
            this.kTP13Max.Text = Prog_Ayarlar.Default.kTP13Max;

            this.iDutyCycleMin.Text = Prog_Ayarlar.Default.iDutyCycleMin;
            this.iFrekansMin.Text = Prog_Ayarlar.Default.iFrekansMin;
            this.iTP12Min.Text = Prog_Ayarlar.Default.iTP12Min;
            this.iTP13Min.Text = Prog_Ayarlar.Default.iTP13Min;
            this.iDutyCycleMax.Text = Prog_Ayarlar.Default.iDutyCycleMax;
            this.iFrekansMax.Text = Prog_Ayarlar.Default.iFrekansMax;
            this.iTP12Max.Text = Prog_Ayarlar.Default.iTP12Max;
            this.iTP13Max.Text = Prog_Ayarlar.Default.iTP13Max;

        }
        private void toolTip_Load()
        {
            string[] toolTipTitle = new string[10];
            string[] toolTipTool = new string[10];
            toolTipTitle[1] = "Lütfen Gelecek Feedback Verisini Giriniz";
            toolTipTitle[2] = "Lütfen Programlama Türünü Belirleyiniz";
            toolTipTitle[3] = "Lütfen Barcode-Master Verisini Giriniz";
            toolTipTitle[4] = "Lütfen Slave Verisini Giriniz";

            toolTipTool[1] = "Örnek : data[2] = {180,75}; \n2,180,75";
            toolTipTool[2] = "1-) Sadece Master \n2-) Master ve Slave";
            toolTipTool[3] = "Örnek : 167000192 ya da 167000204m \nDikkat Barcode ve Master Batch File Aynı Olmalı";
            toolTipTool[4] = "Örnek : 167000204s";

            toolTip2.Active = true;
            toolTip2.AutoPopDelay = 5000;
            toolTip2.InitialDelay = 1000;
            toolTip2.ReshowDelay = 5000;
            toolTip2.IsBalloon = true;
            toolTip2.UseAnimation = true;
            toolTip2.UseFading = true;
            toolTip2.ShowAlways = true;
            toolTip2.ToolTipIcon = ToolTipIcon.Info;
            toolTip2.ToolTipTitle = toolTipTitle[2];
            toolTip2.SetToolTip(infoPicture2, toolTipTool[2]);

            toolTip3.Active = true;
            toolTip3.AutoPopDelay = 5000;
            toolTip3.InitialDelay = 1000;
            toolTip3.ReshowDelay = 5000;
            toolTip3.IsBalloon = true;
            toolTip3.UseAnimation = true;
            toolTip3.UseFading = true;
            toolTip3.ShowAlways = true;
            toolTip3.ToolTipIcon = ToolTipIcon.Info;
            toolTip3.ToolTipTitle = toolTipTitle[3];
            toolTip3.SetToolTip(infoPicture3, toolTipTool[3]);

            toolTip4.Active = true;
            toolTip4.AutoPopDelay = 5000;
            toolTip4.InitialDelay = 1000;
            toolTip4.ReshowDelay = 5000;
            toolTip4.IsBalloon = true;
            toolTip4.UseAnimation = true;
            toolTip4.UseFading = true;
            toolTip4.ShowAlways = true;
            toolTip4.ToolTipIcon = ToolTipIcon.Info;
            toolTip4.ToolTipTitle = toolTipTitle[4];
            toolTip4.SetToolTip(infoPicture4, toolTipTool[4]);
        }

        private void btnKaydet_Click(object sender, EventArgs e)
        {
            try
            {

                Prog_Ayarlar.Default.barcodeNum = this.barcodeNum.Text;

                Prog_Ayarlar.Default.step1Job = this.step1Job.Text;
                Prog_Ayarlar.Default.step2Job = this.step2Job.Text;
                Prog_Ayarlar.Default.step3Job = this.step3Job.Text;
                Prog_Ayarlar.Default.step4Job = this.step4Job.Text;
                Prog_Ayarlar.Default.step5Job = this.step5Job.Text;
                Prog_Ayarlar.Default.step6Job = this.step6Job.Text;
                Prog_Ayarlar.Default.step7Job = this.step7Job.Text;
                Prog_Ayarlar.Default.step8Job = this.step8Job.Text;
                Prog_Ayarlar.Default.step9Job = this.step9Job.Text;
                Prog_Ayarlar.Default.step10Job = this.step10Job.Text;


                Prog_Ayarlar.Default.barcode1 = this.barcode1.Text;
                Prog_Ayarlar.Default.barcode2 = this.barcode2.Text;
                Prog_Ayarlar.Default.barcode3 = this.barcode3.Text;
                Prog_Ayarlar.Default.barcode4 = this.barcode4.Text;
                Prog_Ayarlar.Default.barcode5 = this.barcode5.Text;
                Prog_Ayarlar.Default.barcode6 = this.barcode6.Text;
                Prog_Ayarlar.Default.barcode7 = this.barcode7.Text;
                Prog_Ayarlar.Default.barcode8 = this.barcode8.Text;
                Prog_Ayarlar.Default.barcode9 = this.barcode9.Text;
                Prog_Ayarlar.Default.barcode10 = this.barcode10.Text;

                Prog_Ayarlar.Default.Sbarcode1 = this.Sbarcode1.Text;
                Prog_Ayarlar.Default.Sbarcode2 = this.Sbarcode2.Text;
                Prog_Ayarlar.Default.Sbarcode3 = this.Sbarcode3.Text;
                Prog_Ayarlar.Default.Sbarcode4 = this.Sbarcode4.Text;
                Prog_Ayarlar.Default.Sbarcode5 = this.Sbarcode5.Text;
                Prog_Ayarlar.Default.Sbarcode6 = this.Sbarcode6.Text;
                Prog_Ayarlar.Default.Sbarcode7 = this.Sbarcode7.Text;
                Prog_Ayarlar.Default.Sbarcode8 = this.Sbarcode8.Text;
                Prog_Ayarlar.Default.Sbarcode9 = this.Sbarcode9.Text;
                Prog_Ayarlar.Default.Sbarcode10 = this.Sbarcode10.Text;

                Prog_Ayarlar.Default.Logdosyayolu = this.txtBatchDosya.Text;
                Prog_Ayarlar.Default.companyNo = this.companyNo.Text;
                Prog_Ayarlar.Default.SAPNo = this.SAPNo.Text;
                Prog_Ayarlar.Default.cardNo = this.cardNo.Text;
                Prog_Ayarlar.Default.gerberVer = this.gerberVer.Text;
                Prog_Ayarlar.Default.BOMVer = this.BOMVer.Text;
                Prog_Ayarlar.Default.ICTRev = this.ICTRev.Text;
                Prog_Ayarlar.Default.FCTRev = this.FCTRev.Text;
                Prog_Ayarlar.Default.softwareVer = this.softwareVer.Text;
                Prog_Ayarlar.Default.softwareRev = this.softwareRev.Text;

                Prog_Ayarlar.Default.iniDosyaYolu = this.txtINIdosya.Text;

                Prog_Ayarlar.Default.chBoxSuccess = this.chBoxSuccess.Checked;
                Prog_Ayarlar.Default.chBoxError1 = this.chBoxError1.Checked;
                Prog_Ayarlar.Default.chBoxError2 = this.chBoxError2.Checked;
                Prog_Ayarlar.Default.chBoxError3 = this.chBoxError3.Checked;
                Prog_Ayarlar.Default.successBatch = this.txtSuccessBatch.Text;
                Prog_Ayarlar.Default.error1Batch = this.txtError1Batch.Text;
                Prog_Ayarlar.Default.error2Batch = this.txtError2Batch.Text;
                Prog_Ayarlar.Default.error3Batch = this.txtError3Batch.Text;

                Prog_Ayarlar.Default.kDutyCycleMin = this.kDutyCycleMin.Text;
                Prog_Ayarlar.Default.kFrekansMin = this.kFrekansMin.Text;
                Prog_Ayarlar.Default.kTP12Min = this.kTP12Min.Text;
                Prog_Ayarlar.Default.kTP13Min = this.kTP13Min.Text;
                Prog_Ayarlar.Default.kDutyCycleMax = this.kDutyCycleMax.Text;              
                Prog_Ayarlar.Default.kFrekansMax = this.kFrekansMax.Text;
                Prog_Ayarlar.Default.kTP12Max = this.kTP12Max.Text;
                Prog_Ayarlar.Default.kTP13Max = this.kTP13Max.Text;

                Prog_Ayarlar.Default.iDutyCycleMin = this.iDutyCycleMin.Text;
                Prog_Ayarlar.Default.iFrekansMin = this.iFrekansMin.Text;
                Prog_Ayarlar.Default.iTP12Min = this.iTP12Min.Text;
                Prog_Ayarlar.Default.iTP13Min = this.iTP13Min.Text;
                Prog_Ayarlar.Default.iDutyCycleMax = this.iDutyCycleMax.Text;
                Prog_Ayarlar.Default.iFrekansMax = this.iFrekansMax.Text;
                Prog_Ayarlar.Default.iTP12Max = this.iTP12Max.Text;
                Prog_Ayarlar.Default.iTP13Max = this.iTP13Max.Text;

                Prog_Ayarlar.Default.Save();

                CustomMessageBox.ShowMessage("Bütün Ayarlar Başarıyla Kaydedildi.", Ayarlar.Default.projectName, MessageBoxButtons.OK, CustomMessageBoxIcon.Information, Color.Yellow);
                this.Close();

                Application.Restart();
            }
            catch (Exception ex)
            {
                CustomMessageBox.ShowMessage("Ayarlar Kayıt Hatası: " + ex.ToString(), Ayarlar.Default.projectName, MessageBoxButtons.OK, CustomMessageBoxIcon.Error, Color.Red);
            }
        }

        private void barcodeNum_SelectedIndexChanged(object sender, EventArgs e)
        {
            step1Job.Enabled = false;
            step2Job.Enabled = false;
            step3Job.Enabled = false;
            step4Job.Enabled = false;
            step5Job.Enabled = false;
            step6Job.Enabled = false;
            step7Job.Enabled = false;
            step8Job.Enabled = false;
            step9Job.Enabled = false;
            step10Job.Enabled = false;
            barcode1.Enabled = false;
            barcode2.Enabled = false;
            barcode3.Enabled = false;
            barcode4.Enabled = false;
            barcode5.Enabled = false;
            barcode6.Enabled = false;
            barcode7.Enabled = false;
            barcode8.Enabled = false;
            barcode9.Enabled = false;
            barcode10.Enabled = false;
            Sbarcode1.Enabled = false;
            Sbarcode2.Enabled = false;
            Sbarcode3.Enabled = false;
            Sbarcode4.Enabled = false;
            Sbarcode5.Enabled = false;
            Sbarcode6.Enabled = false;
            Sbarcode7.Enabled = false;
            Sbarcode8.Enabled = false;
            Sbarcode9.Enabled = false;
            Sbarcode10.Enabled = false;

            if (Convert.ToInt32(barcodeNum.Text) >= 1)
            {
                step1Job.Enabled = true;
                barcode1.Enabled = true;
                Sbarcode1.Enabled = true;
                this.Invoke(new EventHandler(step1Job_SelectedIndexChanged));
            }
            if (Convert.ToInt32(barcodeNum.Text) >= 2)
            {
                step2Job.Enabled = true;
                barcode2.Enabled = true;
                Sbarcode2.Enabled = true;
                this.Invoke(new EventHandler(step2Job_SelectedIndexChanged));
            }
            if (Convert.ToInt32(barcodeNum.Text) >= 3)
            {
                step3Job.Enabled = true;
                barcode3.Enabled = true;
                Sbarcode3.Enabled = true;
                this.Invoke(new EventHandler(step3Job_SelectedIndexChanged));
            }
            if (Convert.ToInt32(barcodeNum.Text) >= 4)
            {
                step4Job.Enabled = true;
                barcode4.Enabled = true;
                Sbarcode4.Enabled = true;
                this.Invoke(new EventHandler(step4Job_SelectedIndexChanged));
            }
            if (Convert.ToInt32(barcodeNum.Text) >= 5)
            {
                step5Job.Enabled = true;
                barcode5.Enabled = true;
                Sbarcode5.Enabled = true;
                this.Invoke(new EventHandler(step5Job_SelectedIndexChanged));
            }
            if (Convert.ToInt32(barcodeNum.Text) >= 6)
            {
                step6Job.Enabled = true;
                barcode6.Enabled = true;
                Sbarcode6.Enabled = true;
                this.Invoke(new EventHandler(step6Job_SelectedIndexChanged));
            }
            if (Convert.ToInt32(barcodeNum.Text) >= 7)
            {
                step7Job.Enabled = true;
                barcode7.Enabled = true;
                Sbarcode7.Enabled = true;
                this.Invoke(new EventHandler(step7Job_SelectedIndexChanged));
            }
            if (Convert.ToInt32(barcodeNum.Text) >= 8)
            {
                step8Job.Enabled = true;
                barcode8.Enabled = true;
                Sbarcode8.Enabled = true;
                this.Invoke(new EventHandler(step8Job_SelectedIndexChanged));
            }
            if (Convert.ToInt32(barcodeNum.Text) >= 9)
            {
                step9Job.Enabled = true;
                barcode9.Enabled = true;
                Sbarcode9.Enabled = true;
                this.Invoke(new EventHandler(step9Job_SelectedIndexChanged));
            }
            if (Convert.ToInt32(barcodeNum.Text) >= 10)
            {
                step10Job.Enabled = true;
                barcode10.Enabled = true;
                Sbarcode10.Enabled = true;
                this.Invoke(new EventHandler(step10Job_SelectedIndexChanged));
            }
        }


        private void step1Job_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sbarcode1.Enabled = true;
            if (step1Job.Text == "1")
                Sbarcode1.Enabled = false;
        }

        private void step2Job_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sbarcode2.Enabled = true;
            if (step2Job.Text == "1")
                Sbarcode2.Enabled = false;
        }

        private void step3Job_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sbarcode3.Enabled = true;
            if (step3Job.Text == "1")
                Sbarcode3.Enabled = false;
        }

        private void step4Job_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sbarcode4.Enabled = true;
            if (step4Job.Text == "1")
                Sbarcode4.Enabled = false;
        }

        private void step5Job_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sbarcode5.Enabled = true;
            if (step5Job.Text == "1")
                Sbarcode5.Enabled = false;
        }

        private void step6Job_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sbarcode6.Enabled = true;
            if (step6Job.Text == "1")
                Sbarcode6.Enabled = false;
        }

        private void step7Job_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sbarcode7.Enabled = true;
            if (step7Job.Text == "1")
                Sbarcode7.Enabled = false;
        }

        private void step8Job_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sbarcode8.Enabled = true;
            if (step8Job.Text == "1")
                Sbarcode8.Enabled = false;
        }

        private void step9Job_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sbarcode9.Enabled = true;
            if (step9Job.Text == "1")
                Sbarcode9.Enabled = false;
        }

        private void step10Job_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sbarcode10.Enabled = true;
            if (step10Job.Text == "1")
                Sbarcode10.Enabled = false;
        }

    

        private void chBoxSuccess_CheckedChanged(object sender, EventArgs e)
        {
            txtSuccessBatch.Enabled = chBoxSuccess.Checked;
        }

        private void chBoxError1_CheckedChanged(object sender, EventArgs e)
        {
            txtError1Batch.Enabled = chBoxError1.Checked;
        }

        private void chBoxError2_CheckedChanged(object sender, EventArgs e)
        {
            txtError2Batch.Enabled = chBoxError2.Checked;
        }

        private void chBoxError3_CheckedChanged(object sender, EventArgs e)
        {
            txtError3Batch.Enabled = chBoxError3.Checked;
        }

        private void btnKaydetIni_Click(object sender, EventArgs e)
        {
            if (txtINIdosya.Text != "")
            {
                INIKaydet ini = new INIKaydet(txtINIdosya.Text);  // @"\Programlama.ini"

                ini.Yaz("barcodeNum", "Metin Kutusu", Convert.ToString(barcodeNum.Text));

                ini.Yaz("step1job", "Metin Kutusu", Convert.ToString(step1Job.Text));
                ini.Yaz("step2job", "Metin Kutusu", Convert.ToString(step2Job.Text));
                ini.Yaz("step3job", "Metin Kutusu", Convert.ToString(step3Job.Text));
                ini.Yaz("step4job", "Metin Kutusu", Convert.ToString(step4Job.Text));
                ini.Yaz("step5job", "Metin Kutusu", Convert.ToString(step5Job.Text));
                ini.Yaz("step6job", "Metin Kutusu", Convert.ToString(step6Job.Text));
                ini.Yaz("step7job", "Metin Kutusu", Convert.ToString(step7Job.Text));
                ini.Yaz("step8job", "Metin Kutusu", Convert.ToString(step8Job.Text));
                ini.Yaz("step9job", "Metin Kutusu", Convert.ToString(step9Job.Text));
                ini.Yaz("step10job", "Metin Kutusu", Convert.ToString(step10Job.Text));

                ini.Yaz("barcode1", "Metin Kutusu", Convert.ToString(barcode1.Text));
                ini.Yaz("barcode2", "Metin Kutusu", Convert.ToString(barcode2.Text));
                ini.Yaz("barcode3", "Metin Kutusu", Convert.ToString(barcode3.Text));
                ini.Yaz("barcode4", "Metin Kutusu", Convert.ToString(barcode4.Text));
                ini.Yaz("barcode5", "Metin Kutusu", Convert.ToString(barcode5.Text));
                ini.Yaz("barcode6", "Metin Kutusu", Convert.ToString(barcode6.Text));
                ini.Yaz("barcode7", "Metin Kutusu", Convert.ToString(barcode7.Text));
                ini.Yaz("barcode8", "Metin Kutusu", Convert.ToString(barcode8.Text));
                ini.Yaz("barcode9", "Metin Kutusu", Convert.ToString(barcode9.Text));
                ini.Yaz("barcode10", "Metin Kutusu", Convert.ToString(barcode10.Text));

                ini.Yaz("Sbarcode1", "Metin Kutusu", Convert.ToString(Sbarcode1.Text));
                ini.Yaz("Sbarcode2", "Metin Kutusu", Convert.ToString(Sbarcode2.Text));
                ini.Yaz("Sbarcode3", "Metin Kutusu", Convert.ToString(Sbarcode3.Text));
                ini.Yaz("Sbarcode4", "Metin Kutusu", Convert.ToString(Sbarcode4.Text));
                ini.Yaz("Sbarcode5", "Metin Kutusu", Convert.ToString(Sbarcode5.Text));
                ini.Yaz("Sbarcode6", "Metin Kutusu", Convert.ToString(Sbarcode6.Text));
                ini.Yaz("Sbarcode7", "Metin Kutusu", Convert.ToString(Sbarcode7.Text));
                ini.Yaz("Sbarcode8", "Metin Kutusu", Convert.ToString(Sbarcode8.Text));
                ini.Yaz("Sbarcode9", "Metin Kutusu", Convert.ToString(Sbarcode9.Text));
                ini.Yaz("Sbarcode10", "Metin Kutusu", Convert.ToString(Sbarcode10.Text));

                ini.Yaz("batchDosya", "Metin Kutusu", Convert.ToString(txtBatchDosya.Text));
                ini.Yaz("companyNo", "Metin Kutusu", Convert.ToString(companyNo.Text));
                ini.Yaz("SAPNo", "Metin Kutusu", Convert.ToString(SAPNo.Text));
                ini.Yaz("cardNo", "Metin Kutusu", Convert.ToString(cardNo.Text));
                ini.Yaz("gerberVer", "Metin Kutusu", Convert.ToString(gerberVer.Text));
                ini.Yaz("BOMVer", "Metin Kutusu", Convert.ToString(BOMVer.Text));
                ini.Yaz("ICTRev", "Metin Kutusu", Convert.ToString(ICTRev.Text));
                ini.Yaz("FCTRev", "Metin Kutusu", Convert.ToString(FCTRev.Text));
                ini.Yaz("softwareVer", "Metin Kutusu", Convert.ToString(softwareVer.Text));
                ini.Yaz("softwareRev", "Metin Kutusu", Convert.ToString(softwareRev.Text));

                ini.Yaz("successChecked", "Metin Kutusu", Convert.ToString(chBoxSuccess.Checked));
                ini.Yaz("error1Checked", "Metin Kutusu", Convert.ToString(chBoxError1.Checked));
                ini.Yaz("error2Checked", "Metin Kutusu", Convert.ToString(chBoxError2.Checked));
                ini.Yaz("error3Checked", "Metin Kutusu", Convert.ToString(chBoxError3.Checked));
                ini.Yaz("successBatch", "Metin Kutusu", Convert.ToString(txtSuccessBatch.Text));
                ini.Yaz("error1Batch", "Metin Kutusu", Convert.ToString(txtError1Batch.Text));
                ini.Yaz("error2Batch", "Metin Kutusu", Convert.ToString(txtError2Batch.Text));
                ini.Yaz("error3Batch", "Metin Kutusu", Convert.ToString(txtError3Batch.Text));

                ini.Yaz("kDutyCycleMin", "Metin Kutusu", Convert.ToString(kDutyCycleMin.Text));
                ini.Yaz("kFrekansMin", "Metin Kutusu", Convert.ToString(kFrekansMin.Text));
                ini.Yaz("kTP12Min", "Metin Kutusu", Convert.ToString(kTP12Min.Text));
                ini.Yaz("kTP13Min", "Metin Kutusu", Convert.ToString(kTP13Min.Text));
                ini.Yaz("kDutyCycleMax", "Metin Kutusu", Convert.ToString(kDutyCycleMax.Text));
                ini.Yaz("kFrekansMax", "Metin Kutusu", Convert.ToString(kFrekansMax.Text));
                ini.Yaz("kTP12Max", "Metin Kutusu", Convert.ToString(kTP12Max.Text));
                ini.Yaz("kTP13Max", "Metin Kutusu", Convert.ToString(kTP13Max.Text));

                ini.Yaz("iDutyCycleMin", "Metin Kutusu", Convert.ToString(iDutyCycleMin.Text));
                ini.Yaz("iFrekansMin", "Metin Kutusu", Convert.ToString(iFrekansMin.Text));
                ini.Yaz("iTP12Min", "Metin Kutusu", Convert.ToString(iTP12Min.Text));
                ini.Yaz("iTP13Min", "Metin Kutusu", Convert.ToString(iTP13Min.Text));
                ini.Yaz("iDutyCycleMax", "Metin Kutusu", Convert.ToString(iDutyCycleMax.Text));
                ini.Yaz("iFrekansMax", "Metin Kutusu", Convert.ToString(iFrekansMax.Text));
                ini.Yaz("iTP12Max", "Metin Kutusu", Convert.ToString(iTP12Max.Text));
                ini.Yaz("iTP13Max", "Metin Kutusu", Convert.ToString(iTP13Max.Text));

                CustomMessageBox.ShowMessage("Bütün Ayarlar Dosyaya Başarıyla Kaydedildi.", Ayarlar.Default.projectName, MessageBoxButtons.OK, CustomMessageBoxIcon.Information, Color.Yellow);
            }
            else
            {
                CustomMessageBox.ShowMessage("Dosya Yolu Boş Kalamaz", Ayarlar.Default.projectName, MessageBoxButtons.OK, CustomMessageBoxIcon.Error, Color.Red);
            }
        }

        private void btnOkuIni_Click(object sender, EventArgs e)
        {
            if (txtINIdosya.Text != "")
            {
                try
                {
                    if (File.Exists(txtINIdosya.Text))
                    {
                        INIKaydet ini = new INIKaydet(txtINIdosya.Text);

                        barcodeNum.Text = ini.Oku("barcodeNum", "Metin Kutusu");

                        step1Job.Text = ini.Oku("step1job", "Metin Kutusu");
                        step2Job.Text = ini.Oku("step2job", "Metin Kutusu");
                        step3Job.Text = ini.Oku("step3job", "Metin Kutusu");
                        step4Job.Text = ini.Oku("step4job", "Metin Kutusu");
                        step5Job.Text = ini.Oku("step5job", "Metin Kutusu");
                        step6Job.Text = ini.Oku("step6job", "Metin Kutusu");
                        step7Job.Text = ini.Oku("step7job", "Metin Kutusu");
                        step8Job.Text = ini.Oku("step8job", "Metin Kutusu");
                        step9Job.Text = ini.Oku("step9job", "Metin Kutusu");
                        step10Job.Text = ini.Oku("step10job", "Metin Kutusu");

                        barcode1.Text = ini.Oku("barcode1", "Metin Kutusu");
                        barcode2.Text = ini.Oku("barcode2", "Metin Kutusu");
                        barcode3.Text = ini.Oku("barcode3", "Metin Kutusu");
                        barcode4.Text = ini.Oku("barcode4", "Metin Kutusu");
                        barcode5.Text = ini.Oku("barcode5", "Metin Kutusu");
                        barcode6.Text = ini.Oku("barcode6", "Metin Kutusu");
                        barcode7.Text = ini.Oku("barcode7", "Metin Kutusu");
                        barcode8.Text = ini.Oku("barcode8", "Metin Kutusu");
                        barcode9.Text = ini.Oku("barcode9", "Metin Kutusu");
                        barcode10.Text = ini.Oku("barcode10", "Metin Kutusu");

                        Sbarcode1.Text = ini.Oku("Sbarcode1", "Metin Kutusu");
                        Sbarcode2.Text = ini.Oku("Sbarcode2", "Metin Kutusu");
                        Sbarcode3.Text = ini.Oku("Sbarcode3", "Metin Kutusu");
                        Sbarcode4.Text = ini.Oku("Sbarcode4", "Metin Kutusu");
                        Sbarcode5.Text = ini.Oku("Sbarcode5", "Metin Kutusu");
                        Sbarcode6.Text = ini.Oku("Sbarcode6", "Metin Kutusu");
                        Sbarcode7.Text = ini.Oku("Sbarcode7", "Metin Kutusu");
                        Sbarcode8.Text = ini.Oku("Sbarcode8", "Metin Kutusu");
                        Sbarcode9.Text = ini.Oku("Sbarcode9", "Metin Kutusu");
                        Sbarcode10.Text = ini.Oku("Sbarcode10", "Metin Kutusu");

                        txtBatchDosya.Text = ini.Oku("batchDosya", "Metin Kutusu");
                        companyNo.Text = ini.Oku("companyNo", "Metin Kutusu");
                        SAPNo.Text = ini.Oku("SAPNo", "Metin Kutusu");
                        cardNo.Text = ini.Oku("cardNo", "Metin Kutusu");
                        gerberVer.Text = ini.Oku("gerberVer", "Metin Kutusu");
                        BOMVer.Text = ini.Oku("BOMVer", "Metin Kutusu");
                        ICTRev.Text = ini.Oku("ICTRev", "Metin Kutusu");
                        FCTRev.Text = ini.Oku("FCTRev", "Metin Kutusu");
                        softwareVer.Text = ini.Oku("softwareVer", "Metin Kutusu");
                        softwareRev.Text = ini.Oku("softwareRev", "Metin Kutusu");

                        if (ini.Oku("successChecked", "Metin Kutusu") == "True")
                            chBoxSuccess.Checked = true;
                        else if (ini.Oku("successChecked", "Metin Kutusu") == "False")
                            chBoxSuccess.Checked = false;

                        if (ini.Oku("error1Checked", "Metin Kutusu") == "True")
                            chBoxError1.Checked = true;
                        else if (ini.Oku("error1Checked", "Metin Kutusu") == "False")
                            chBoxError1.Checked = false;

                        if (ini.Oku("error2Checked", "Metin Kutusu") == "True")
                            chBoxError2.Checked = true;
                        else if (ini.Oku("error2Checked", "Metin Kutusu") == "False")
                            chBoxError2.Checked = false;

                        if (ini.Oku("error3Checked", "Metin Kutusu") == "True")
                            chBoxError3.Checked = true;
                        else if (ini.Oku("error3Checked", "Metin Kutusu") == "False")
                            chBoxError3.Checked = false;

                        txtSuccessBatch.Text = ini.Oku("successBatch", "Metin Kutusu");
                        txtError1Batch.Text = ini.Oku("error1Batch", "Metin Kutusu");
                        txtError2Batch.Text = ini.Oku("error2Batch", "Metin Kutusu");
                        txtError3Batch.Text = ini.Oku("error3Batch", "Metin Kutusu");

                        kDutyCycleMin.Text = ini.Oku("kDutyCycleMin", "Metin Kutusu");
                        kFrekansMin.Text = ini.Oku("kFrekansMin", "Metin Kutusu");
                        kTP12Min.Text = ini.Oku("kTP12Min", "Metin Kutusu");
                        kTP13Min.Text = ini.Oku("kTP13Min", "Metin Kutusu");
                        kDutyCycleMax.Text = ini.Oku("kDutyCycleMax", "Metin Kutusu");
                        kFrekansMax.Text = ini.Oku("kFrekansMax", "Metin Kutusu");
                        kTP12Max.Text = ini.Oku("kTP12Max", "Metin Kutusu");
                        kTP13Max.Text = ini.Oku("kTP13Max", "Metin Kutusu");

                        iDutyCycleMin.Text = ini.Oku("iDutyCycleMin", "Metin Kutusu");
                        iFrekansMin.Text = ini.Oku("iFrekansMin", "Metin Kutusu");
                        iTP12Min.Text = ini.Oku("iTP12Min", "Metin Kutusu");
                        iTP13Min.Text = ini.Oku("iTP13Min", "Metin Kutusu");
                        iDutyCycleMax.Text = ini.Oku("iDutyCycleMax", "Metin Kutusu");
                        iFrekansMax.Text = ini.Oku("iFrekansMax", "Metin Kutusu");
                        iTP12Max.Text = ini.Oku("iTP12Max", "Metin Kutusu");
                        iTP13Max.Text = ini.Oku("iTP13Max", "Metin Kutusu");

                        CustomMessageBox.ShowMessage("Bütün Ayarlar Dosyadan Başarıyla Okundu.", Ayarlar.Default.projectName, MessageBoxButtons.OK, CustomMessageBoxIcon.Information, Color.Yellow);
                    }
                }
                catch (Exception hata)
                {
                    CustomMessageBox.ShowMessage("ini Dosyası Hasarlı" + hata.Message, Ayarlar.Default.projectName, MessageBoxButtons.OK, CustomMessageBoxIcon.Error, Color.Red);
                }
            }
            else
            {
                CustomMessageBox.ShowMessage("Dosya Yolu Boş Kalamaz", Ayarlar.Default.projectName, MessageBoxButtons.OK, CustomMessageBoxIcon.Error, Color.Red);
            }
        }

        private void btnIDsec_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "|*.ini";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;
            this.txtINIdosya.Text = openFileDialog.FileName;
        }

        private void btnLogsec_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Database Dosyaları|*.bat|Tüm Dosyalar|*.*";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;
            this.txtBatchDosya.Text = openFileDialog.FileName;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && this.components != null)
                this.components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProgAyarForm));
            this.btnKaydet = new System.Windows.Forms.Button();
            this.barcodeNum = new System.Windows.Forms.ComboBox();
            this.barcode1 = new System.Windows.Forms.TextBox();
            this.barcode2 = new System.Windows.Forms.TextBox();
            this.barcode4 = new System.Windows.Forms.TextBox();
            this.barcode3 = new System.Windows.Forms.TextBox();
            this.barcode8 = new System.Windows.Forms.TextBox();
            this.barcode7 = new System.Windows.Forms.TextBox();
            this.barcode6 = new System.Windows.Forms.TextBox();
            this.barcode5 = new System.Windows.Forms.TextBox();
            this.barcode10 = new System.Windows.Forms.TextBox();
            this.barcode9 = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.txtBatchDosya = new System.Windows.Forms.TextBox();
            this.btnLogsec = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.companyNo = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            this.SAPNo = new System.Windows.Forms.TextBox();
            this.label28 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btnINIsec = new System.Windows.Forms.Button();
            this.txtINIdosya = new System.Windows.Forms.TextBox();
            this.label220 = new System.Windows.Forms.Label();
            this.btnOkuIni = new System.Windows.Forms.Button();
            this.btnKaydetIni = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.softwareRev = new System.Windows.Forms.TextBox();
            this.label31 = new System.Windows.Forms.Label();
            this.label29 = new System.Windows.Forms.Label();
            this.FCTRev = new System.Windows.Forms.TextBox();
            this.softwareVer = new System.Windows.Forms.TextBox();
            this.label30 = new System.Windows.Forms.Label();
            this.label26 = new System.Windows.Forms.Label();
            this.BOMVer = new System.Windows.Forms.TextBox();
            this.ICTRev = new System.Windows.Forms.TextBox();
            this.label27 = new System.Windows.Forms.Label();
            this.label24 = new System.Windows.Forms.Label();
            this.cardNo = new System.Windows.Forms.TextBox();
            this.gerberVer = new System.Windows.Forms.TextBox();
            this.label25 = new System.Windows.Forms.Label();
            this.step10Job = new System.Windows.Forms.ComboBox();
            this.label68 = new System.Windows.Forms.Label();
            this.step9Job = new System.Windows.Forms.ComboBox();
            this.label69 = new System.Windows.Forms.Label();
            this.step8Job = new System.Windows.Forms.ComboBox();
            this.label70 = new System.Windows.Forms.Label();
            this.step7Job = new System.Windows.Forms.ComboBox();
            this.label71 = new System.Windows.Forms.Label();
            this.step6Job = new System.Windows.Forms.ComboBox();
            this.label72 = new System.Windows.Forms.Label();
            this.step5Job = new System.Windows.Forms.ComboBox();
            this.label37 = new System.Windows.Forms.Label();
            this.step4Job = new System.Windows.Forms.ComboBox();
            this.label38 = new System.Windows.Forms.Label();
            this.step3Job = new System.Windows.Forms.ComboBox();
            this.label39 = new System.Windows.Forms.Label();
            this.step2Job = new System.Windows.Forms.ComboBox();
            this.label40 = new System.Windows.Forms.Label();
            this.step1Job = new System.Windows.Forms.ComboBox();
            this.label41 = new System.Windows.Forms.Label();
            this.toolTip2 = new System.Windows.Forms.ToolTip(this.components);
            this.label52 = new System.Windows.Forms.Label();
            this.label53 = new System.Windows.Forms.Label();
            this.label54 = new System.Windows.Forms.Label();
            this.label55 = new System.Windows.Forms.Label();
            this.label56 = new System.Windows.Forms.Label();
            this.label57 = new System.Windows.Forms.Label();
            this.label58 = new System.Windows.Forms.Label();
            this.label59 = new System.Windows.Forms.Label();
            this.label60 = new System.Windows.Forms.Label();
            this.label61 = new System.Windows.Forms.Label();
            this.Sbarcode10 = new System.Windows.Forms.TextBox();
            this.Sbarcode9 = new System.Windows.Forms.TextBox();
            this.Sbarcode8 = new System.Windows.Forms.TextBox();
            this.Sbarcode7 = new System.Windows.Forms.TextBox();
            this.Sbarcode6 = new System.Windows.Forms.TextBox();
            this.Sbarcode5 = new System.Windows.Forms.TextBox();
            this.Sbarcode4 = new System.Windows.Forms.TextBox();
            this.Sbarcode3 = new System.Windows.Forms.TextBox();
            this.Sbarcode2 = new System.Windows.Forms.TextBox();
            this.Sbarcode1 = new System.Windows.Forms.TextBox();
            this.toolTip3 = new System.Windows.Forms.ToolTip(this.components);
            this.toolTip4 = new System.Windows.Forms.ToolTip(this.components);
            this.chBoxSuccess = new System.Windows.Forms.CheckBox();
            this.chBoxError1 = new System.Windows.Forms.CheckBox();
            this.chBoxError2 = new System.Windows.Forms.CheckBox();
            this.chBoxError3 = new System.Windows.Forms.CheckBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.txtError3Batch = new System.Windows.Forms.TextBox();
            this.txtError2Batch = new System.Windows.Forms.TextBox();
            this.txtError1Batch = new System.Windows.Forms.TextBox();
            this.txtSuccessBatch = new System.Windows.Forms.TextBox();
            this.infoPicture4 = new System.Windows.Forms.PictureBox();
            this.infoPicture3 = new System.Windows.Forms.PictureBox();
            this.infoPicture2 = new System.Windows.Forms.PictureBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.label65 = new System.Windows.Forms.Label();
            this.label66 = new System.Windows.Forms.Label();
            this.label67 = new System.Windows.Forms.Label();
            this.label73 = new System.Windows.Forms.Label();
            this.label51 = new System.Windows.Forms.Label();
            this.iTP13Min = new System.Windows.Forms.TextBox();
            this.label62 = new System.Windows.Forms.Label();
            this.label32 = new System.Windows.Forms.Label();
            this.label63 = new System.Windows.Forms.Label();
            this.iTP13Max = new System.Windows.Forms.TextBox();
            this.label64 = new System.Windows.Forms.Label();
            this.label35 = new System.Windows.Forms.Label();
            this.iTP12Min = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.iTP12Max = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.iFrekansMin = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.iDutyCycleMin = new System.Windows.Forms.TextBox();
            this.iFrekansMax = new System.Windows.Forms.TextBox();
            this.label22 = new System.Windows.Forms.Label();
            this.label33 = new System.Windows.Forms.Label();
            this.label34 = new System.Windows.Forms.Label();
            this.iDutyCycleMax = new System.Windows.Forms.TextBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.label49 = new System.Windows.Forms.Label();
            this.label50 = new System.Windows.Forms.Label();
            this.label48 = new System.Windows.Forms.Label();
            this.label47 = new System.Windows.Forms.Label();
            this.label46 = new System.Windows.Forms.Label();
            this.label45 = new System.Windows.Forms.Label();
            this.label44 = new System.Windows.Forms.Label();
            this.label43 = new System.Windows.Forms.Label();
            this.kTP13Min = new System.Windows.Forms.TextBox();
            this.label36 = new System.Windows.Forms.Label();
            this.kTP13Max = new System.Windows.Forms.TextBox();
            this.label42 = new System.Windows.Forms.Label();
            this.kTP12Min = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.kTP12Max = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.kFrekansMin = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.kDutyCycleMin = new System.Windows.Forms.TextBox();
            this.kFrekansMax = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.kDutyCycleMax = new System.Windows.Forms.TextBox();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.infoPicture4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.infoPicture3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.infoPicture2)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnKaydet
            // 
            this.btnKaydet.BackColor = System.Drawing.Color.Aqua;
            this.btnKaydet.Font = new System.Drawing.Font("Calibri", 42F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnKaydet.ForeColor = System.Drawing.Color.Black;
            this.btnKaydet.Location = new System.Drawing.Point(634, 617);
            this.btnKaydet.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnKaydet.Name = "btnKaydet";
            this.btnKaydet.Size = new System.Drawing.Size(379, 157);
            this.btnKaydet.TabIndex = 6;
            this.btnKaydet.Text = "Ayarları Kaydet";
            this.btnKaydet.UseVisualStyleBackColor = false;
            this.btnKaydet.Click += new System.EventHandler(this.btnKaydet_Click);
            // 
            // barcodeNum
            // 
            this.barcodeNum.FormattingEnabled = true;
            this.barcodeNum.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
            "13",
            "14",
            "16",
            "17",
            "18",
            "19",
            "20"});
            this.barcodeNum.Location = new System.Drawing.Point(249, 22);
            this.barcodeNum.Name = "barcodeNum";
            this.barcodeNum.Size = new System.Drawing.Size(115, 23);
            this.barcodeNum.TabIndex = 12;
            this.barcodeNum.SelectedIndexChanged += new System.EventHandler(this.barcodeNum_SelectedIndexChanged);
            // 
            // barcode1
            // 
            this.barcode1.Location = new System.Drawing.Point(249, 81);
            this.barcode1.Name = "barcode1";
            this.barcode1.Size = new System.Drawing.Size(115, 24);
            this.barcode1.TabIndex = 13;
            // 
            // barcode2
            // 
            this.barcode2.Location = new System.Drawing.Point(249, 116);
            this.barcode2.Name = "barcode2";
            this.barcode2.Size = new System.Drawing.Size(115, 24);
            this.barcode2.TabIndex = 14;
            // 
            // barcode4
            // 
            this.barcode4.Location = new System.Drawing.Point(249, 186);
            this.barcode4.Name = "barcode4";
            this.barcode4.Size = new System.Drawing.Size(115, 24);
            this.barcode4.TabIndex = 16;
            // 
            // barcode3
            // 
            this.barcode3.Location = new System.Drawing.Point(249, 151);
            this.barcode3.Name = "barcode3";
            this.barcode3.Size = new System.Drawing.Size(115, 24);
            this.barcode3.TabIndex = 15;
            // 
            // barcode8
            // 
            this.barcode8.Location = new System.Drawing.Point(249, 326);
            this.barcode8.Name = "barcode8";
            this.barcode8.Size = new System.Drawing.Size(115, 24);
            this.barcode8.TabIndex = 20;
            // 
            // barcode7
            // 
            this.barcode7.Location = new System.Drawing.Point(249, 291);
            this.barcode7.Name = "barcode7";
            this.barcode7.Size = new System.Drawing.Size(115, 24);
            this.barcode7.TabIndex = 19;
            // 
            // barcode6
            // 
            this.barcode6.Location = new System.Drawing.Point(249, 256);
            this.barcode6.Name = "barcode6";
            this.barcode6.Size = new System.Drawing.Size(115, 24);
            this.barcode6.TabIndex = 18;
            // 
            // barcode5
            // 
            this.barcode5.Location = new System.Drawing.Point(249, 221);
            this.barcode5.Name = "barcode5";
            this.barcode5.Size = new System.Drawing.Size(115, 24);
            this.barcode5.TabIndex = 17;
            // 
            // barcode10
            // 
            this.barcode10.Location = new System.Drawing.Point(249, 396);
            this.barcode10.Name = "barcode10";
            this.barcode10.Size = new System.Drawing.Size(115, 24);
            this.barcode10.TabIndex = 22;
            // 
            // barcode9
            // 
            this.barcode9.Location = new System.Drawing.Point(249, 361);
            this.barcode9.Name = "barcode9";
            this.barcode9.Size = new System.Drawing.Size(115, 24);
            this.barcode9.TabIndex = 21;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(6, 25);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(108, 17);
            this.label12.TabIndex = 0;
            this.label12.Text = "Batch Dosya Yolu:";
            // 
            // txtBatchDosya
            // 
            this.txtBatchDosya.Location = new System.Drawing.Point(126, 22);
            this.txtBatchDosya.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtBatchDosya.Name = "txtBatchDosya";
            this.txtBatchDosya.Size = new System.Drawing.Size(172, 24);
            this.txtBatchDosya.TabIndex = 1;
            // 
            // btnLogsec
            // 
            this.btnLogsec.BackColor = System.Drawing.Color.Aqua;
            this.btnLogsec.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnLogsec.Location = new System.Drawing.Point(306, 22);
            this.btnLogsec.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnLogsec.Name = "btnLogsec";
            this.btnLogsec.Size = new System.Drawing.Size(65, 24);
            this.btnLogsec.TabIndex = 2;
            this.btnLogsec.Text = "Seç";
            this.btnLogsec.UseVisualStyleBackColor = false;
            this.btnLogsec.Click += new System.EventHandler(this.btnLogsec_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(168, 83);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 17);
            this.label1.TabIndex = 26;
            this.label1.Text = "Barcode1:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(168, 118);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 17);
            this.label2.TabIndex = 27;
            this.label2.Text = "Barcode2:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(168, 153);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(66, 17);
            this.label3.TabIndex = 28;
            this.label3.Text = "Barcode3:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(168, 258);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(66, 17);
            this.label4.TabIndex = 31;
            this.label4.Text = "Barcode6:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(168, 223);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 17);
            this.label5.TabIndex = 30;
            this.label5.Text = "Barcode5:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(168, 188);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(66, 17);
            this.label6.TabIndex = 29;
            this.label6.Text = "Barcode4:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(168, 363);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(66, 17);
            this.label7.TabIndex = 34;
            this.label7.Text = "Barcode9:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(168, 328);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(66, 17);
            this.label8.TabIndex = 33;
            this.label8.Text = "Barcode8:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(168, 293);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(66, 17);
            this.label9.TabIndex = 32;
            this.label9.Text = "Barcode7:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(168, 398);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(73, 17);
            this.label10.TabIndex = 35;
            this.label10.Text = "Barcode10:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(151, 25);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(92, 17);
            this.label11.TabIndex = 36;
            this.label11.Text = "Barcode Sayısı:";
            // 
            // companyNo
            // 
            this.companyNo.Location = new System.Drawing.Point(125, 57);
            this.companyNo.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.companyNo.Name = "companyNo";
            this.companyNo.Size = new System.Drawing.Size(172, 24);
            this.companyNo.TabIndex = 58;
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(30, 60);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(83, 17);
            this.label23.TabIndex = 57;
            this.label23.Text = "Company No:";
            // 
            // SAPNo
            // 
            this.SAPNo.Location = new System.Drawing.Point(125, 95);
            this.SAPNo.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.SAPNo.Name = "SAPNo";
            this.SAPNo.Size = new System.Drawing.Size(172, 24);
            this.SAPNo.TabIndex = 60;
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.Location = new System.Drawing.Point(58, 98);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(52, 17);
            this.label28.TabIndex = 59;
            this.label28.Text = "SAP No:";
            // 
            // btnINIsec
            // 
            this.btnINIsec.BackColor = System.Drawing.Color.Aqua;
            this.btnINIsec.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnINIsec.Location = new System.Drawing.Point(306, 23);
            this.btnINIsec.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnINIsec.Name = "btnINIsec";
            this.btnINIsec.Size = new System.Drawing.Size(65, 24);
            this.btnINIsec.TabIndex = 587;
            this.btnINIsec.Text = "Seç";
            this.btnINIsec.UseVisualStyleBackColor = false;
            this.btnINIsec.Click += new System.EventHandler(this.btnIDsec_Click);
            // 
            // txtINIdosya
            // 
            this.txtINIdosya.Location = new System.Drawing.Point(131, 22);
            this.txtINIdosya.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtINIdosya.Name = "txtINIdosya";
            this.txtINIdosya.Size = new System.Drawing.Size(167, 24);
            this.txtINIdosya.TabIndex = 586;
            // 
            // label220
            // 
            this.label220.AutoSize = true;
            this.label220.Location = new System.Drawing.Point(9, 22);
            this.label220.Name = "label220";
            this.label220.Size = new System.Drawing.Size(116, 17);
            this.label220.TabIndex = 585;
            this.label220.Text = "Ayarlar Dosya Yolu:";
            // 
            // btnOkuIni
            // 
            this.btnOkuIni.BackColor = System.Drawing.Color.Aqua;
            this.btnOkuIni.Location = new System.Drawing.Point(217, 55);
            this.btnOkuIni.Name = "btnOkuIni";
            this.btnOkuIni.Size = new System.Drawing.Size(80, 30);
            this.btnOkuIni.TabIndex = 584;
            this.btnOkuIni.Text = "Oku";
            this.btnOkuIni.UseVisualStyleBackColor = false;
            this.btnOkuIni.Click += new System.EventHandler(this.btnOkuIni_Click);
            // 
            // btnKaydetIni
            // 
            this.btnKaydetIni.BackColor = System.Drawing.Color.Aqua;
            this.btnKaydetIni.Location = new System.Drawing.Point(131, 55);
            this.btnKaydetIni.Name = "btnKaydetIni";
            this.btnKaydetIni.Size = new System.Drawing.Size(80, 30);
            this.btnKaydetIni.TabIndex = 583;
            this.btnKaydetIni.Text = "Kaydet";
            this.btnKaydetIni.UseVisualStyleBackColor = false;
            this.btnKaydetIni.Click += new System.EventHandler(this.btnKaydetIni_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label220);
            this.groupBox2.Controls.Add(this.btnOkuIni);
            this.groupBox2.Controls.Add(this.btnINIsec);
            this.groupBox2.Controls.Add(this.btnKaydetIni);
            this.groupBox2.Controls.Add(this.txtINIdosya);
            this.groupBox2.Location = new System.Drawing.Point(634, 394);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(379, 92);
            this.groupBox2.TabIndex = 588;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Ini Dosyası Ayarları:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.softwareRev);
            this.groupBox3.Controls.Add(this.label31);
            this.groupBox3.Controls.Add(this.label29);
            this.groupBox3.Controls.Add(this.FCTRev);
            this.groupBox3.Controls.Add(this.softwareVer);
            this.groupBox3.Controls.Add(this.label30);
            this.groupBox3.Controls.Add(this.label26);
            this.groupBox3.Controls.Add(this.BOMVer);
            this.groupBox3.Controls.Add(this.ICTRev);
            this.groupBox3.Controls.Add(this.label27);
            this.groupBox3.Controls.Add(this.label24);
            this.groupBox3.Controls.Add(this.cardNo);
            this.groupBox3.Controls.Add(this.gerberVer);
            this.groupBox3.Controls.Add(this.label25);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.txtBatchDosya);
            this.groupBox3.Controls.Add(this.btnLogsec);
            this.groupBox3.Controls.Add(this.label23);
            this.groupBox3.Controls.Add(this.companyNo);
            this.groupBox3.Controls.Add(this.SAPNo);
            this.groupBox3.Controls.Add(this.label28);
            this.groupBox3.Location = new System.Drawing.Point(634, 3);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(379, 385);
            this.groupBox3.TabIndex = 589;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Proje Ayarları:";
            // 
            // softwareRev
            // 
            this.softwareRev.Location = new System.Drawing.Point(125, 350);
            this.softwareRev.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.softwareRev.Name = "softwareRev";
            this.softwareRev.Size = new System.Drawing.Size(172, 24);
            this.softwareRev.TabIndex = 74;
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Location = new System.Drawing.Point(14, 353);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(87, 17);
            this.label31.TabIndex = 73;
            this.label31.Text = "Software Rev:";
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Location = new System.Drawing.Point(46, 281);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(56, 17);
            this.label29.TabIndex = 69;
            this.label29.Text = "FCT Rev:";
            // 
            // FCTRev
            // 
            this.FCTRev.Location = new System.Drawing.Point(125, 278);
            this.FCTRev.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.FCTRev.Name = "FCTRev";
            this.FCTRev.Size = new System.Drawing.Size(172, 24);
            this.FCTRev.TabIndex = 70;
            // 
            // softwareVer
            // 
            this.softwareVer.Location = new System.Drawing.Point(125, 314);
            this.softwareVer.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.softwareVer.Name = "softwareVer";
            this.softwareVer.Size = new System.Drawing.Size(172, 24);
            this.softwareVer.TabIndex = 72;
            // 
            // label30
            // 
            this.label30.AutoSize = true;
            this.label30.Location = new System.Drawing.Point(17, 317);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(85, 17);
            this.label30.TabIndex = 71;
            this.label30.Text = "Software Ver:";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(43, 208);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(63, 17);
            this.label26.TabIndex = 65;
            this.label26.Text = "BOM Ver:";
            // 
            // BOMVer
            // 
            this.BOMVer.Location = new System.Drawing.Point(125, 205);
            this.BOMVer.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BOMVer.Name = "BOMVer";
            this.BOMVer.Size = new System.Drawing.Size(172, 24);
            this.BOMVer.TabIndex = 66;
            // 
            // ICTRev
            // 
            this.ICTRev.Location = new System.Drawing.Point(125, 241);
            this.ICTRev.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ICTRev.Name = "ICTRev";
            this.ICTRev.Size = new System.Drawing.Size(172, 24);
            this.ICTRev.TabIndex = 68;
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(51, 244);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(54, 17);
            this.label27.TabIndex = 67;
            this.label27.Text = "ICT Rev:";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(51, 136);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(57, 17);
            this.label24.TabIndex = 61;
            this.label24.Text = "Card No:";
            // 
            // cardNo
            // 
            this.cardNo.Location = new System.Drawing.Point(125, 133);
            this.cardNo.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.cardNo.Name = "cardNo";
            this.cardNo.Size = new System.Drawing.Size(172, 24);
            this.cardNo.TabIndex = 62;
            // 
            // gerberVer
            // 
            this.gerberVer.Location = new System.Drawing.Point(125, 169);
            this.gerberVer.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.gerberVer.Name = "gerberVer";
            this.gerberVer.Size = new System.Drawing.Size(172, 24);
            this.gerberVer.TabIndex = 64;
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(35, 172);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(74, 17);
            this.label25.TabIndex = 63;
            this.label25.Text = "Gerber Ver:";
            // 
            // step10Job
            // 
            this.step10Job.FormattingEnabled = true;
            this.step10Job.Items.AddRange(new object[] {
            "1",
            "2"});
            this.step10Job.Location = new System.Drawing.Point(74, 396);
            this.step10Job.Name = "step10Job";
            this.step10Job.Size = new System.Drawing.Size(58, 23);
            this.step10Job.TabIndex = 609;
            this.step10Job.SelectedIndexChanged += new System.EventHandler(this.step10Job_SelectedIndexChanged);
            // 
            // label68
            // 
            this.label68.AutoSize = true;
            this.label68.Location = new System.Drawing.Point(13, 398);
            this.label68.Name = "label68";
            this.label68.Size = new System.Drawing.Size(51, 17);
            this.label68.TabIndex = 608;
            this.label68.Text = "Step10:";
            // 
            // step9Job
            // 
            this.step9Job.FormattingEnabled = true;
            this.step9Job.Items.AddRange(new object[] {
            "1",
            "2"});
            this.step9Job.Location = new System.Drawing.Point(74, 361);
            this.step9Job.Name = "step9Job";
            this.step9Job.Size = new System.Drawing.Size(58, 23);
            this.step9Job.TabIndex = 607;
            this.step9Job.SelectedIndexChanged += new System.EventHandler(this.step9Job_SelectedIndexChanged);
            // 
            // label69
            // 
            this.label69.AutoSize = true;
            this.label69.Location = new System.Drawing.Point(13, 363);
            this.label69.Name = "label69";
            this.label69.Size = new System.Drawing.Size(44, 17);
            this.label69.TabIndex = 606;
            this.label69.Text = "Step9:";
            // 
            // step8Job
            // 
            this.step8Job.FormattingEnabled = true;
            this.step8Job.Items.AddRange(new object[] {
            "1",
            "2"});
            this.step8Job.Location = new System.Drawing.Point(74, 326);
            this.step8Job.Name = "step8Job";
            this.step8Job.Size = new System.Drawing.Size(58, 23);
            this.step8Job.TabIndex = 605;
            this.step8Job.SelectedIndexChanged += new System.EventHandler(this.step8Job_SelectedIndexChanged);
            // 
            // label70
            // 
            this.label70.AutoSize = true;
            this.label70.Location = new System.Drawing.Point(13, 328);
            this.label70.Name = "label70";
            this.label70.Size = new System.Drawing.Size(44, 17);
            this.label70.TabIndex = 604;
            this.label70.Text = "Step8:";
            // 
            // step7Job
            // 
            this.step7Job.FormattingEnabled = true;
            this.step7Job.Items.AddRange(new object[] {
            "1",
            "2"});
            this.step7Job.Location = new System.Drawing.Point(74, 291);
            this.step7Job.Name = "step7Job";
            this.step7Job.Size = new System.Drawing.Size(58, 23);
            this.step7Job.TabIndex = 603;
            this.step7Job.SelectedIndexChanged += new System.EventHandler(this.step7Job_SelectedIndexChanged);
            // 
            // label71
            // 
            this.label71.AutoSize = true;
            this.label71.Location = new System.Drawing.Point(13, 293);
            this.label71.Name = "label71";
            this.label71.Size = new System.Drawing.Size(44, 17);
            this.label71.TabIndex = 602;
            this.label71.Text = "Step7:";
            // 
            // step6Job
            // 
            this.step6Job.FormattingEnabled = true;
            this.step6Job.Items.AddRange(new object[] {
            "1",
            "2"});
            this.step6Job.Location = new System.Drawing.Point(74, 256);
            this.step6Job.Name = "step6Job";
            this.step6Job.Size = new System.Drawing.Size(58, 23);
            this.step6Job.TabIndex = 601;
            this.step6Job.SelectedIndexChanged += new System.EventHandler(this.step6Job_SelectedIndexChanged);
            // 
            // label72
            // 
            this.label72.AutoSize = true;
            this.label72.Location = new System.Drawing.Point(13, 258);
            this.label72.Name = "label72";
            this.label72.Size = new System.Drawing.Size(44, 17);
            this.label72.TabIndex = 600;
            this.label72.Text = "Step6:";
            // 
            // step5Job
            // 
            this.step5Job.FormattingEnabled = true;
            this.step5Job.Items.AddRange(new object[] {
            "1",
            "2"});
            this.step5Job.Location = new System.Drawing.Point(74, 221);
            this.step5Job.Name = "step5Job";
            this.step5Job.Size = new System.Drawing.Size(58, 23);
            this.step5Job.TabIndex = 599;
            this.step5Job.SelectedIndexChanged += new System.EventHandler(this.step5Job_SelectedIndexChanged);
            // 
            // label37
            // 
            this.label37.AutoSize = true;
            this.label37.Location = new System.Drawing.Point(13, 223);
            this.label37.Name = "label37";
            this.label37.Size = new System.Drawing.Size(44, 17);
            this.label37.TabIndex = 598;
            this.label37.Text = "Step5:";
            // 
            // step4Job
            // 
            this.step4Job.FormattingEnabled = true;
            this.step4Job.Items.AddRange(new object[] {
            "1",
            "2"});
            this.step4Job.Location = new System.Drawing.Point(74, 186);
            this.step4Job.Name = "step4Job";
            this.step4Job.Size = new System.Drawing.Size(58, 23);
            this.step4Job.TabIndex = 597;
            this.step4Job.SelectedIndexChanged += new System.EventHandler(this.step4Job_SelectedIndexChanged);
            // 
            // label38
            // 
            this.label38.AutoSize = true;
            this.label38.Location = new System.Drawing.Point(13, 188);
            this.label38.Name = "label38";
            this.label38.Size = new System.Drawing.Size(44, 17);
            this.label38.TabIndex = 596;
            this.label38.Text = "Step4:";
            // 
            // step3Job
            // 
            this.step3Job.FormattingEnabled = true;
            this.step3Job.Items.AddRange(new object[] {
            "1",
            "2"});
            this.step3Job.Location = new System.Drawing.Point(74, 151);
            this.step3Job.Name = "step3Job";
            this.step3Job.Size = new System.Drawing.Size(58, 23);
            this.step3Job.TabIndex = 595;
            this.step3Job.SelectedIndexChanged += new System.EventHandler(this.step3Job_SelectedIndexChanged);
            // 
            // label39
            // 
            this.label39.AutoSize = true;
            this.label39.Location = new System.Drawing.Point(13, 153);
            this.label39.Name = "label39";
            this.label39.Size = new System.Drawing.Size(44, 17);
            this.label39.TabIndex = 594;
            this.label39.Text = "Step3:";
            // 
            // step2Job
            // 
            this.step2Job.FormattingEnabled = true;
            this.step2Job.Items.AddRange(new object[] {
            "1",
            "2"});
            this.step2Job.Location = new System.Drawing.Point(74, 116);
            this.step2Job.Name = "step2Job";
            this.step2Job.Size = new System.Drawing.Size(58, 23);
            this.step2Job.TabIndex = 593;
            this.step2Job.SelectedIndexChanged += new System.EventHandler(this.step2Job_SelectedIndexChanged);
            // 
            // label40
            // 
            this.label40.AutoSize = true;
            this.label40.Location = new System.Drawing.Point(13, 118);
            this.label40.Name = "label40";
            this.label40.Size = new System.Drawing.Size(44, 17);
            this.label40.TabIndex = 592;
            this.label40.Text = "Step2:";
            // 
            // step1Job
            // 
            this.step1Job.FormattingEnabled = true;
            this.step1Job.Items.AddRange(new object[] {
            "1",
            "2"});
            this.step1Job.Location = new System.Drawing.Point(74, 81);
            this.step1Job.Name = "step1Job";
            this.step1Job.Size = new System.Drawing.Size(58, 23);
            this.step1Job.TabIndex = 591;
            this.step1Job.SelectedIndexChanged += new System.EventHandler(this.step1Job_SelectedIndexChanged);
            // 
            // label41
            // 
            this.label41.AutoSize = true;
            this.label41.Location = new System.Drawing.Point(13, 83);
            this.label41.Name = "label41";
            this.label41.Size = new System.Drawing.Size(44, 17);
            this.label41.TabIndex = 590;
            this.label41.Text = "Step1:";
            // 
            // label52
            // 
            this.label52.AutoSize = true;
            this.label52.Location = new System.Drawing.Point(390, 398);
            this.label52.Name = "label52";
            this.label52.Size = new System.Drawing.Size(73, 17);
            this.label52.TabIndex = 650;
            this.label52.Text = "Barcode10:";
            // 
            // label53
            // 
            this.label53.AutoSize = true;
            this.label53.Location = new System.Drawing.Point(390, 363);
            this.label53.Name = "label53";
            this.label53.Size = new System.Drawing.Size(66, 17);
            this.label53.TabIndex = 649;
            this.label53.Text = "Barcode9:";
            // 
            // label54
            // 
            this.label54.AutoSize = true;
            this.label54.Location = new System.Drawing.Point(390, 328);
            this.label54.Name = "label54";
            this.label54.Size = new System.Drawing.Size(66, 17);
            this.label54.TabIndex = 648;
            this.label54.Text = "Barcode8:";
            // 
            // label55
            // 
            this.label55.AutoSize = true;
            this.label55.Location = new System.Drawing.Point(390, 293);
            this.label55.Name = "label55";
            this.label55.Size = new System.Drawing.Size(66, 17);
            this.label55.TabIndex = 647;
            this.label55.Text = "Barcode7:";
            // 
            // label56
            // 
            this.label56.AutoSize = true;
            this.label56.Location = new System.Drawing.Point(390, 258);
            this.label56.Name = "label56";
            this.label56.Size = new System.Drawing.Size(66, 17);
            this.label56.TabIndex = 646;
            this.label56.Text = "Barcode6:";
            // 
            // label57
            // 
            this.label57.AutoSize = true;
            this.label57.Location = new System.Drawing.Point(390, 223);
            this.label57.Name = "label57";
            this.label57.Size = new System.Drawing.Size(66, 17);
            this.label57.TabIndex = 645;
            this.label57.Text = "Barcode5:";
            // 
            // label58
            // 
            this.label58.AutoSize = true;
            this.label58.Location = new System.Drawing.Point(390, 188);
            this.label58.Name = "label58";
            this.label58.Size = new System.Drawing.Size(66, 17);
            this.label58.TabIndex = 644;
            this.label58.Text = "Barcode4:";
            // 
            // label59
            // 
            this.label59.AutoSize = true;
            this.label59.Location = new System.Drawing.Point(390, 153);
            this.label59.Name = "label59";
            this.label59.Size = new System.Drawing.Size(66, 17);
            this.label59.TabIndex = 643;
            this.label59.Text = "Barcode3:";
            // 
            // label60
            // 
            this.label60.AutoSize = true;
            this.label60.Location = new System.Drawing.Point(390, 118);
            this.label60.Name = "label60";
            this.label60.Size = new System.Drawing.Size(66, 17);
            this.label60.TabIndex = 642;
            this.label60.Text = "Barcode2:";
            // 
            // label61
            // 
            this.label61.AutoSize = true;
            this.label61.Location = new System.Drawing.Point(390, 83);
            this.label61.Name = "label61";
            this.label61.Size = new System.Drawing.Size(66, 17);
            this.label61.TabIndex = 641;
            this.label61.Text = "Barcode1:";
            // 
            // Sbarcode10
            // 
            this.Sbarcode10.Location = new System.Drawing.Point(471, 396);
            this.Sbarcode10.Name = "Sbarcode10";
            this.Sbarcode10.Size = new System.Drawing.Size(115, 24);
            this.Sbarcode10.TabIndex = 640;
            // 
            // Sbarcode9
            // 
            this.Sbarcode9.Location = new System.Drawing.Point(471, 361);
            this.Sbarcode9.Name = "Sbarcode9";
            this.Sbarcode9.Size = new System.Drawing.Size(115, 24);
            this.Sbarcode9.TabIndex = 639;
            // 
            // Sbarcode8
            // 
            this.Sbarcode8.Location = new System.Drawing.Point(471, 326);
            this.Sbarcode8.Name = "Sbarcode8";
            this.Sbarcode8.Size = new System.Drawing.Size(115, 24);
            this.Sbarcode8.TabIndex = 638;
            // 
            // Sbarcode7
            // 
            this.Sbarcode7.Location = new System.Drawing.Point(471, 291);
            this.Sbarcode7.Name = "Sbarcode7";
            this.Sbarcode7.Size = new System.Drawing.Size(115, 24);
            this.Sbarcode7.TabIndex = 637;
            // 
            // Sbarcode6
            // 
            this.Sbarcode6.Location = new System.Drawing.Point(471, 256);
            this.Sbarcode6.Name = "Sbarcode6";
            this.Sbarcode6.Size = new System.Drawing.Size(115, 24);
            this.Sbarcode6.TabIndex = 636;
            // 
            // Sbarcode5
            // 
            this.Sbarcode5.Location = new System.Drawing.Point(471, 221);
            this.Sbarcode5.Name = "Sbarcode5";
            this.Sbarcode5.Size = new System.Drawing.Size(115, 24);
            this.Sbarcode5.TabIndex = 635;
            // 
            // Sbarcode4
            // 
            this.Sbarcode4.Location = new System.Drawing.Point(471, 186);
            this.Sbarcode4.Name = "Sbarcode4";
            this.Sbarcode4.Size = new System.Drawing.Size(115, 24);
            this.Sbarcode4.TabIndex = 634;
            // 
            // Sbarcode3
            // 
            this.Sbarcode3.Location = new System.Drawing.Point(471, 151);
            this.Sbarcode3.Name = "Sbarcode3";
            this.Sbarcode3.Size = new System.Drawing.Size(115, 24);
            this.Sbarcode3.TabIndex = 633;
            // 
            // Sbarcode2
            // 
            this.Sbarcode2.Location = new System.Drawing.Point(471, 116);
            this.Sbarcode2.Name = "Sbarcode2";
            this.Sbarcode2.Size = new System.Drawing.Size(115, 24);
            this.Sbarcode2.TabIndex = 632;
            // 
            // Sbarcode1
            // 
            this.Sbarcode1.Location = new System.Drawing.Point(471, 81);
            this.Sbarcode1.Name = "Sbarcode1";
            this.Sbarcode1.Size = new System.Drawing.Size(115, 24);
            this.Sbarcode1.TabIndex = 631;
            // 
            // chBoxSuccess
            // 
            this.chBoxSuccess.AutoSize = true;
            this.chBoxSuccess.Location = new System.Drawing.Point(10, 18);
            this.chBoxSuccess.Name = "chBoxSuccess";
            this.chBoxSuccess.Size = new System.Drawing.Size(69, 21);
            this.chBoxSuccess.TabIndex = 673;
            this.chBoxSuccess.Text = "Success";
            this.chBoxSuccess.UseVisualStyleBackColor = true;
            this.chBoxSuccess.CheckedChanged += new System.EventHandler(this.chBoxSuccess_CheckedChanged);
            // 
            // chBoxError1
            // 
            this.chBoxError1.AutoSize = true;
            this.chBoxError1.Location = new System.Drawing.Point(10, 43);
            this.chBoxError1.Name = "chBoxError1";
            this.chBoxError1.Size = new System.Drawing.Size(63, 21);
            this.chBoxError1.TabIndex = 674;
            this.chBoxError1.Text = "Error1";
            this.chBoxError1.UseVisualStyleBackColor = true;
            this.chBoxError1.CheckedChanged += new System.EventHandler(this.chBoxError1_CheckedChanged);
            // 
            // chBoxError2
            // 
            this.chBoxError2.AutoSize = true;
            this.chBoxError2.Location = new System.Drawing.Point(10, 68);
            this.chBoxError2.Name = "chBoxError2";
            this.chBoxError2.Size = new System.Drawing.Size(63, 21);
            this.chBoxError2.TabIndex = 675;
            this.chBoxError2.Text = "Error2";
            this.chBoxError2.UseVisualStyleBackColor = true;
            this.chBoxError2.CheckedChanged += new System.EventHandler(this.chBoxError2_CheckedChanged);
            // 
            // chBoxError3
            // 
            this.chBoxError3.AutoSize = true;
            this.chBoxError3.Location = new System.Drawing.Point(10, 93);
            this.chBoxError3.Name = "chBoxError3";
            this.chBoxError3.Size = new System.Drawing.Size(63, 21);
            this.chBoxError3.TabIndex = 676;
            this.chBoxError3.Text = "Error3";
            this.chBoxError3.UseVisualStyleBackColor = true;
            this.chBoxError3.CheckedChanged += new System.EventHandler(this.chBoxError3_CheckedChanged);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.txtError3Batch);
            this.groupBox4.Controls.Add(this.txtError2Batch);
            this.groupBox4.Controls.Add(this.txtError1Batch);
            this.groupBox4.Controls.Add(this.txtSuccessBatch);
            this.groupBox4.Controls.Add(this.chBoxError3);
            this.groupBox4.Controls.Add(this.chBoxSuccess);
            this.groupBox4.Controls.Add(this.chBoxError2);
            this.groupBox4.Controls.Add(this.chBoxError1);
            this.groupBox4.Location = new System.Drawing.Point(634, 492);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(379, 120);
            this.groupBox4.TabIndex = 677;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Batch Error Name";
            // 
            // txtError3Batch
            // 
            this.txtError3Batch.Location = new System.Drawing.Point(85, 90);
            this.txtError3Batch.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtError3Batch.Name = "txtError3Batch";
            this.txtError3Batch.Size = new System.Drawing.Size(283, 24);
            this.txtError3Batch.TabIndex = 680;
            // 
            // txtError2Batch
            // 
            this.txtError2Batch.Location = new System.Drawing.Point(85, 65);
            this.txtError2Batch.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtError2Batch.Name = "txtError2Batch";
            this.txtError2Batch.Size = new System.Drawing.Size(283, 24);
            this.txtError2Batch.TabIndex = 679;
            // 
            // txtError1Batch
            // 
            this.txtError1Batch.Location = new System.Drawing.Point(85, 40);
            this.txtError1Batch.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtError1Batch.Name = "txtError1Batch";
            this.txtError1Batch.Size = new System.Drawing.Size(283, 24);
            this.txtError1Batch.TabIndex = 678;
            // 
            // txtSuccessBatch
            // 
            this.txtSuccessBatch.Location = new System.Drawing.Point(85, 15);
            this.txtSuccessBatch.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtSuccessBatch.Name = "txtSuccessBatch";
            this.txtSuccessBatch.Size = new System.Drawing.Size(283, 24);
            this.txtSuccessBatch.TabIndex = 677;
            // 
            // infoPicture4
            // 
            this.infoPicture4.Image = ((System.Drawing.Image)(resources.GetObject("infoPicture4.Image")));
            this.infoPicture4.Location = new System.Drawing.Point(516, 61);
            this.infoPicture4.Name = "infoPicture4";
            this.infoPicture4.Size = new System.Drawing.Size(20, 20);
            this.infoPicture4.TabIndex = 672;
            this.infoPicture4.TabStop = false;
            // 
            // infoPicture3
            // 
            this.infoPicture3.Image = ((System.Drawing.Image)(resources.GetObject("infoPicture3.Image")));
            this.infoPicture3.Location = new System.Drawing.Point(296, 61);
            this.infoPicture3.Name = "infoPicture3";
            this.infoPicture3.Size = new System.Drawing.Size(20, 20);
            this.infoPicture3.TabIndex = 671;
            this.infoPicture3.TabStop = false;
            // 
            // infoPicture2
            // 
            this.infoPicture2.Image = ((System.Drawing.Image)(resources.GetObject("infoPicture2.Image")));
            this.infoPicture2.Location = new System.Drawing.Point(92, 61);
            this.infoPicture2.Name = "infoPicture2";
            this.infoPicture2.Size = new System.Drawing.Size(20, 20);
            this.infoPicture2.TabIndex = 630;
            this.infoPicture2.TabStop = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox6);
            this.groupBox1.Controls.Add(this.groupBox5);
            this.groupBox1.Controls.Add(this.infoPicture3);
            this.groupBox1.Controls.Add(this.barcode1);
            this.groupBox1.Controls.Add(this.infoPicture4);
            this.groupBox1.Controls.Add(this.barcode2);
            this.groupBox1.Controls.Add(this.barcode3);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.barcodeNum);
            this.groupBox1.Controls.Add(this.barcode4);
            this.groupBox1.Controls.Add(this.barcode5);
            this.groupBox1.Controls.Add(this.barcode6);
            this.groupBox1.Controls.Add(this.barcode7);
            this.groupBox1.Controls.Add(this.barcode8);
            this.groupBox1.Controls.Add(this.barcode9);
            this.groupBox1.Controls.Add(this.barcode10);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label52);
            this.groupBox1.Controls.Add(this.label53);
            this.groupBox1.Controls.Add(this.label54);
            this.groupBox1.Controls.Add(this.label55);
            this.groupBox1.Controls.Add(this.label56);
            this.groupBox1.Controls.Add(this.label57);
            this.groupBox1.Controls.Add(this.label58);
            this.groupBox1.Controls.Add(this.label59);
            this.groupBox1.Controls.Add(this.label60);
            this.groupBox1.Controls.Add(this.label61);
            this.groupBox1.Controls.Add(this.Sbarcode10);
            this.groupBox1.Controls.Add(this.Sbarcode9);
            this.groupBox1.Controls.Add(this.Sbarcode8);
            this.groupBox1.Controls.Add(this.Sbarcode7);
            this.groupBox1.Controls.Add(this.Sbarcode6);
            this.groupBox1.Controls.Add(this.Sbarcode5);
            this.groupBox1.Controls.Add(this.Sbarcode4);
            this.groupBox1.Controls.Add(this.Sbarcode3);
            this.groupBox1.Controls.Add(this.label41);
            this.groupBox1.Controls.Add(this.Sbarcode2);
            this.groupBox1.Controls.Add(this.step1Job);
            this.groupBox1.Controls.Add(this.Sbarcode1);
            this.groupBox1.Controls.Add(this.label40);
            this.groupBox1.Controls.Add(this.infoPicture2);
            this.groupBox1.Controls.Add(this.step2Job);
            this.groupBox1.Controls.Add(this.label39);
            this.groupBox1.Controls.Add(this.step3Job);
            this.groupBox1.Controls.Add(this.label38);
            this.groupBox1.Controls.Add(this.step4Job);
            this.groupBox1.Controls.Add(this.label37);
            this.groupBox1.Controls.Add(this.step5Job);
            this.groupBox1.Controls.Add(this.label72);
            this.groupBox1.Controls.Add(this.step6Job);
            this.groupBox1.Controls.Add(this.label71);
            this.groupBox1.Controls.Add(this.step7Job);
            this.groupBox1.Controls.Add(this.label70);
            this.groupBox1.Controls.Add(this.step8Job);
            this.groupBox1.Controls.Add(this.label69);
            this.groupBox1.Controls.Add(this.step9Job);
            this.groupBox1.Controls.Add(this.label68);
            this.groupBox1.Controls.Add(this.step10Job);
            this.groupBox1.Location = new System.Drawing.Point(12, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(596, 771);
            this.groupBox1.TabIndex = 678;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Barkod Ayarları";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.label65);
            this.groupBox6.Controls.Add(this.label66);
            this.groupBox6.Controls.Add(this.label67);
            this.groupBox6.Controls.Add(this.label73);
            this.groupBox6.Controls.Add(this.label51);
            this.groupBox6.Controls.Add(this.iTP13Min);
            this.groupBox6.Controls.Add(this.label62);
            this.groupBox6.Controls.Add(this.label32);
            this.groupBox6.Controls.Add(this.label63);
            this.groupBox6.Controls.Add(this.iTP13Max);
            this.groupBox6.Controls.Add(this.label64);
            this.groupBox6.Controls.Add(this.label35);
            this.groupBox6.Controls.Add(this.iTP12Min);
            this.groupBox6.Controls.Add(this.label19);
            this.groupBox6.Controls.Add(this.iTP12Max);
            this.groupBox6.Controls.Add(this.label20);
            this.groupBox6.Controls.Add(this.iFrekansMin);
            this.groupBox6.Controls.Add(this.label21);
            this.groupBox6.Controls.Add(this.iDutyCycleMin);
            this.groupBox6.Controls.Add(this.iFrekansMax);
            this.groupBox6.Controls.Add(this.label22);
            this.groupBox6.Controls.Add(this.label33);
            this.groupBox6.Controls.Add(this.label34);
            this.groupBox6.Controls.Add(this.iDutyCycleMax);
            this.groupBox6.Location = new System.Drawing.Point(16, 596);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(570, 163);
            this.groupBox6.TabIndex = 698;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "İNCE (313000)";
            // 
            // label65
            // 
            this.label65.AutoSize = true;
            this.label65.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label65.ForeColor = System.Drawing.Color.Red;
            this.label65.Location = new System.Drawing.Point(545, 131);
            this.label65.Name = "label65";
            this.label65.Size = new System.Drawing.Size(18, 19);
            this.label65.TabIndex = 704;
            this.label65.Text = "V";
            // 
            // label66
            // 
            this.label66.AutoSize = true;
            this.label66.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label66.ForeColor = System.Drawing.Color.Red;
            this.label66.Location = new System.Drawing.Point(545, 93);
            this.label66.Name = "label66";
            this.label66.Size = new System.Drawing.Size(18, 19);
            this.label66.TabIndex = 703;
            this.label66.Text = "V";
            // 
            // label67
            // 
            this.label67.AutoSize = true;
            this.label67.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label67.ForeColor = System.Drawing.Color.Red;
            this.label67.Location = new System.Drawing.Point(542, 57);
            this.label67.Name = "label67";
            this.label67.Size = new System.Drawing.Size(23, 19);
            this.label67.TabIndex = 702;
            this.label67.Text = "hz";
            // 
            // label73
            // 
            this.label73.AutoSize = true;
            this.label73.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label73.ForeColor = System.Drawing.Color.Red;
            this.label73.Location = new System.Drawing.Point(542, 20);
            this.label73.Name = "label73";
            this.label73.Size = new System.Drawing.Size(20, 19);
            this.label73.TabIndex = 701;
            this.label73.Text = "%";
            // 
            // label51
            // 
            this.label51.AutoSize = true;
            this.label51.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label51.ForeColor = System.Drawing.Color.Red;
            this.label51.Location = new System.Drawing.Point(258, 131);
            this.label51.Name = "label51";
            this.label51.Size = new System.Drawing.Size(18, 19);
            this.label51.TabIndex = 700;
            this.label51.Text = "V";
            // 
            // iTP13Min
            // 
            this.iTP13Min.Location = new System.Drawing.Point(97, 128);
            this.iTP13Min.Name = "iTP13Min";
            this.iTP13Min.Size = new System.Drawing.Size(152, 24);
            this.iTP13Min.TabIndex = 685;
            // 
            // label62
            // 
            this.label62.AutoSize = true;
            this.label62.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label62.ForeColor = System.Drawing.Color.Red;
            this.label62.Location = new System.Drawing.Point(258, 94);
            this.label62.Name = "label62";
            this.label62.Size = new System.Drawing.Size(18, 19);
            this.label62.TabIndex = 699;
            this.label62.Text = "V";
            // 
            // label32
            // 
            this.label32.AutoSize = true;
            this.label32.Location = new System.Drawing.Point(287, 131);
            this.label32.Name = "label32";
            this.label32.Size = new System.Drawing.Size(68, 17);
            this.label32.TabIndex = 688;
            this.label32.Text = "TP13 Max:";
            // 
            // label63
            // 
            this.label63.AutoSize = true;
            this.label63.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label63.ForeColor = System.Drawing.Color.Red;
            this.label63.Location = new System.Drawing.Point(255, 58);
            this.label63.Name = "label63";
            this.label63.Size = new System.Drawing.Size(23, 19);
            this.label63.TabIndex = 698;
            this.label63.Text = "hz";
            // 
            // iTP13Max
            // 
            this.iTP13Max.Location = new System.Drawing.Point(386, 128);
            this.iTP13Max.Name = "iTP13Max";
            this.iTP13Max.Size = new System.Drawing.Size(152, 24);
            this.iTP13Max.TabIndex = 687;
            // 
            // label64
            // 
            this.label64.AutoSize = true;
            this.label64.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label64.ForeColor = System.Drawing.Color.Red;
            this.label64.Location = new System.Drawing.Point(255, 21);
            this.label64.Name = "label64";
            this.label64.Size = new System.Drawing.Size(20, 19);
            this.label64.TabIndex = 697;
            this.label64.Text = "%";
            // 
            // label35
            // 
            this.label35.AutoSize = true;
            this.label35.Location = new System.Drawing.Point(5, 131);
            this.label35.Name = "label35";
            this.label35.Size = new System.Drawing.Size(65, 17);
            this.label35.TabIndex = 686;
            this.label35.Text = "TP13 Min:";
            // 
            // iTP12Min
            // 
            this.iTP12Min.Location = new System.Drawing.Point(97, 90);
            this.iTP12Min.Name = "iTP12Min";
            this.iTP12Min.Size = new System.Drawing.Size(152, 24);
            this.iTP12Min.TabIndex = 681;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(287, 93);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(68, 17);
            this.label19.TabIndex = 684;
            this.label19.Text = "TP12 Max:";
            // 
            // iTP12Max
            // 
            this.iTP12Max.Location = new System.Drawing.Point(386, 90);
            this.iTP12Max.Name = "iTP12Max";
            this.iTP12Max.Size = new System.Drawing.Size(152, 24);
            this.iTP12Max.TabIndex = 683;
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(5, 93);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(65, 17);
            this.label20.TabIndex = 682;
            this.label20.Text = "TP12 Min:";
            // 
            // iFrekansMin
            // 
            this.iFrekansMin.Location = new System.Drawing.Point(97, 54);
            this.iFrekansMin.Name = "iFrekansMin";
            this.iFrekansMin.Size = new System.Drawing.Size(152, 24);
            this.iFrekansMin.TabIndex = 675;
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(287, 57);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(83, 17);
            this.label21.TabIndex = 680;
            this.label21.Text = "Frekans Max:";
            // 
            // iDutyCycleMin
            // 
            this.iDutyCycleMin.Location = new System.Drawing.Point(97, 19);
            this.iDutyCycleMin.Name = "iDutyCycleMin";
            this.iDutyCycleMin.Size = new System.Drawing.Size(152, 24);
            this.iDutyCycleMin.TabIndex = 673;
            // 
            // iFrekansMax
            // 
            this.iFrekansMax.Location = new System.Drawing.Point(386, 54);
            this.iFrekansMax.Name = "iFrekansMax";
            this.iFrekansMax.Size = new System.Drawing.Size(152, 24);
            this.iFrekansMax.TabIndex = 679;
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(5, 22);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(93, 17);
            this.label22.TabIndex = 674;
            this.label22.Text = "DutyCycle Min:";
            // 
            // label33
            // 
            this.label33.AutoSize = true;
            this.label33.Location = new System.Drawing.Point(287, 22);
            this.label33.Name = "label33";
            this.label33.Size = new System.Drawing.Size(96, 17);
            this.label33.TabIndex = 678;
            this.label33.Text = "DutyCycle Max:";
            // 
            // label34
            // 
            this.label34.AutoSize = true;
            this.label34.Location = new System.Drawing.Point(5, 57);
            this.label34.Name = "label34";
            this.label34.Size = new System.Drawing.Size(80, 17);
            this.label34.TabIndex = 676;
            this.label34.Text = "Frekans Min:";
            // 
            // iDutyCycleMax
            // 
            this.iDutyCycleMax.Location = new System.Drawing.Point(386, 19);
            this.iDutyCycleMax.Name = "iDutyCycleMax";
            this.iDutyCycleMax.Size = new System.Drawing.Size(152, 24);
            this.iDutyCycleMax.TabIndex = 677;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.label49);
            this.groupBox5.Controls.Add(this.label50);
            this.groupBox5.Controls.Add(this.label48);
            this.groupBox5.Controls.Add(this.label47);
            this.groupBox5.Controls.Add(this.label46);
            this.groupBox5.Controls.Add(this.label45);
            this.groupBox5.Controls.Add(this.label44);
            this.groupBox5.Controls.Add(this.label43);
            this.groupBox5.Controls.Add(this.kTP13Min);
            this.groupBox5.Controls.Add(this.label36);
            this.groupBox5.Controls.Add(this.kTP13Max);
            this.groupBox5.Controls.Add(this.label42);
            this.groupBox5.Controls.Add(this.kTP12Min);
            this.groupBox5.Controls.Add(this.label17);
            this.groupBox5.Controls.Add(this.kTP12Max);
            this.groupBox5.Controls.Add(this.label18);
            this.groupBox5.Controls.Add(this.kFrekansMin);
            this.groupBox5.Controls.Add(this.label15);
            this.groupBox5.Controls.Add(this.kDutyCycleMin);
            this.groupBox5.Controls.Add(this.kFrekansMax);
            this.groupBox5.Controls.Add(this.label13);
            this.groupBox5.Controls.Add(this.label16);
            this.groupBox5.Controls.Add(this.label14);
            this.groupBox5.Controls.Add(this.kDutyCycleMax);
            this.groupBox5.Location = new System.Drawing.Point(16, 439);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(570, 151);
            this.groupBox5.TabIndex = 697;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "KALIN (312700)";
            // 
            // label49
            // 
            this.label49.AutoSize = true;
            this.label49.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label49.ForeColor = System.Drawing.Color.Red;
            this.label49.Location = new System.Drawing.Point(544, 129);
            this.label49.Name = "label49";
            this.label49.Size = new System.Drawing.Size(18, 19);
            this.label49.TabIndex = 696;
            this.label49.Text = "V";
            // 
            // label50
            // 
            this.label50.AutoSize = true;
            this.label50.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label50.ForeColor = System.Drawing.Color.Red;
            this.label50.Location = new System.Drawing.Point(544, 95);
            this.label50.Name = "label50";
            this.label50.Size = new System.Drawing.Size(18, 19);
            this.label50.TabIndex = 695;
            this.label50.Text = "V";
            // 
            // label48
            // 
            this.label48.AutoSize = true;
            this.label48.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label48.ForeColor = System.Drawing.Color.Red;
            this.label48.Location = new System.Drawing.Point(258, 127);
            this.label48.Name = "label48";
            this.label48.Size = new System.Drawing.Size(18, 19);
            this.label48.TabIndex = 694;
            this.label48.Text = "V";
            // 
            // label47
            // 
            this.label47.AutoSize = true;
            this.label47.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label47.ForeColor = System.Drawing.Color.Red;
            this.label47.Location = new System.Drawing.Point(258, 93);
            this.label47.Name = "label47";
            this.label47.Size = new System.Drawing.Size(18, 19);
            this.label47.TabIndex = 693;
            this.label47.Text = "V";
            // 
            // label46
            // 
            this.label46.AutoSize = true;
            this.label46.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label46.ForeColor = System.Drawing.Color.Red;
            this.label46.Location = new System.Drawing.Point(544, 54);
            this.label46.Name = "label46";
            this.label46.Size = new System.Drawing.Size(23, 19);
            this.label46.TabIndex = 692;
            this.label46.Text = "hz";
            // 
            // label45
            // 
            this.label45.AutoSize = true;
            this.label45.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label45.ForeColor = System.Drawing.Color.Red;
            this.label45.Location = new System.Drawing.Point(255, 57);
            this.label45.Name = "label45";
            this.label45.Size = new System.Drawing.Size(23, 19);
            this.label45.TabIndex = 691;
            this.label45.Text = "hz";
            // 
            // label44
            // 
            this.label44.AutoSize = true;
            this.label44.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label44.ForeColor = System.Drawing.Color.Red;
            this.label44.Location = new System.Drawing.Point(544, 22);
            this.label44.Name = "label44";
            this.label44.Size = new System.Drawing.Size(20, 19);
            this.label44.TabIndex = 690;
            this.label44.Text = "%";
            // 
            // label43
            // 
            this.label43.AutoSize = true;
            this.label43.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label43.ForeColor = System.Drawing.Color.Red;
            this.label43.Location = new System.Drawing.Point(255, 20);
            this.label43.Name = "label43";
            this.label43.Size = new System.Drawing.Size(20, 19);
            this.label43.TabIndex = 689;
            this.label43.Text = "%";
            // 
            // kTP13Min
            // 
            this.kTP13Min.Location = new System.Drawing.Point(97, 127);
            this.kTP13Min.Name = "kTP13Min";
            this.kTP13Min.Size = new System.Drawing.Size(152, 24);
            this.kTP13Min.TabIndex = 685;
            // 
            // label36
            // 
            this.label36.AutoSize = true;
            this.label36.Location = new System.Drawing.Point(287, 130);
            this.label36.Name = "label36";
            this.label36.Size = new System.Drawing.Size(68, 17);
            this.label36.TabIndex = 688;
            this.label36.Text = "TP13 Max:";
            // 
            // kTP13Max
            // 
            this.kTP13Max.Location = new System.Drawing.Point(386, 127);
            this.kTP13Max.Name = "kTP13Max";
            this.kTP13Max.Size = new System.Drawing.Size(152, 24);
            this.kTP13Max.TabIndex = 687;
            // 
            // label42
            // 
            this.label42.AutoSize = true;
            this.label42.Location = new System.Drawing.Point(5, 130);
            this.label42.Name = "label42";
            this.label42.Size = new System.Drawing.Size(65, 17);
            this.label42.TabIndex = 686;
            this.label42.Text = "TP13 Min:";
            // 
            // kTP12Min
            // 
            this.kTP12Min.Location = new System.Drawing.Point(97, 90);
            this.kTP12Min.Name = "kTP12Min";
            this.kTP12Min.Size = new System.Drawing.Size(152, 24);
            this.kTP12Min.TabIndex = 681;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(287, 93);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(68, 17);
            this.label17.TabIndex = 684;
            this.label17.Text = "TP12 Max:";
            // 
            // kTP12Max
            // 
            this.kTP12Max.Location = new System.Drawing.Point(386, 90);
            this.kTP12Max.Name = "kTP12Max";
            this.kTP12Max.Size = new System.Drawing.Size(152, 24);
            this.kTP12Max.TabIndex = 683;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(5, 93);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(65, 17);
            this.label18.TabIndex = 682;
            this.label18.Text = "TP12 Min:";
            // 
            // kFrekansMin
            // 
            this.kFrekansMin.Location = new System.Drawing.Point(97, 54);
            this.kFrekansMin.Name = "kFrekansMin";
            this.kFrekansMin.Size = new System.Drawing.Size(152, 24);
            this.kFrekansMin.TabIndex = 675;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(287, 57);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(83, 17);
            this.label15.TabIndex = 680;
            this.label15.Text = "Frekans Max:";
            // 
            // kDutyCycleMin
            // 
            this.kDutyCycleMin.Location = new System.Drawing.Point(97, 19);
            this.kDutyCycleMin.Name = "kDutyCycleMin";
            this.kDutyCycleMin.Size = new System.Drawing.Size(152, 24);
            this.kDutyCycleMin.TabIndex = 673;
            // 
            // kFrekansMax
            // 
            this.kFrekansMax.Location = new System.Drawing.Point(386, 54);
            this.kFrekansMax.Name = "kFrekansMax";
            this.kFrekansMax.Size = new System.Drawing.Size(152, 24);
            this.kFrekansMax.TabIndex = 679;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(5, 22);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(93, 17);
            this.label13.TabIndex = 674;
            this.label13.Text = "DutyCycle Min:";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(287, 22);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(96, 17);
            this.label16.TabIndex = 678;
            this.label16.Text = "DutyCycle Max:";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(5, 57);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(80, 17);
            this.label14.TabIndex = 676;
            this.label14.Text = "Frekans Min:";
            // 
            // kDutyCycleMax
            // 
            this.kDutyCycleMax.Location = new System.Drawing.Point(386, 19);
            this.kDutyCycleMax.Name = "kDutyCycleMax";
            this.kDutyCycleMax.Size = new System.Drawing.Size(152, 24);
            this.kDutyCycleMax.TabIndex = 677;
            // 
            // ProgAyarForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.ClientSize = new System.Drawing.Size(1023, 778);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnKaydet);
            this.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "ProgAyarForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ayarlar";
            this.Load += new System.EventHandler(this.AyarForm_Load);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.infoPicture4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.infoPicture3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.infoPicture2)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);

        }
    }
}
