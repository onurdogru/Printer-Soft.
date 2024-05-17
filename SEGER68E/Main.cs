using EasyModbus;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text.RegularExpressions;
using SEGER68E.Printer;
using System.Printing;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Data.SqlClient;

namespace SEGER68E
{
    public partial class Main : Form
    {
        private Thread saniyeThread = null;
        public AyarForm AyarFrm;
        public Sifre SifreFrm;
        public ProgAyarForm ProgAyarFrm;
        public ProcessForm ProcessFrm;
        public ProgramlamaForm ProgramlamaFrm;

        private IntPtr ShellHwnd;
        private DateTime lastDateTime = DateTime.Now;
        private ModbusClient modbusClientPLC = null;

        const int M0 = 2048;
        const int M1 = 2049;
        const int M2 = 2050;
        const int M3 = 2051;
        const int M4 = 2052;
        const int M5 = 2053;
        const int M6 = 2054;
        const int M7 = 2055;
        const int M8 = 2056;
        const int M9 = 2057;
        const int M10 = 2058;

        const int X0 = 1024;
        const int X1 = 1025;
        const int X2 = 1026;
        const int X3 = 1027;
        const int X4 = 1028;
        const int X5 = 1029;
        const int X6 = 1030;
        const int X7 = 1031;
        const int X10 = 1032;
        const int X11 = 1033;
        const int X21 = 1041;
        const int X23 = 1043;
        const int X25 = 1045;
        const int X27 = 1047;
        const int X31 = 1049;
        const int X33 = 1051;
        const int X20 = 1040;
        const int X22 = 1042;
        const int X24 = 1044;
        const int X26 = 1046;
        const int X30 = 1048;
        const int X32 = 1050;

        const int D4 = 4100;
        const int D6 = 4102;
        const int D8 = 4104;
        const int D10 = 4106;
        const int D12 = 4108;
        const int D14 = 4110;
        const int D16 = 4112;
        const int D18 = 4114;
        const int D20 = 4116;
        const int D30 = 4126;
        const int D40 = 4136;
        const int D400 = 4496;
        const int D405 = 4501;
        const int D410 = 4506;
        const int D415 = 4511;
        const int D440 = 4536;
        const int D445 = 4541;
        const int D450 = 4546;
        const int D455 = 4551;
        const int D460 = 4556;
        const int D465 = 4561;

        const int Y1 = 1281;
        const int Y2 = 1282;
        const int Y3 = 1283;
        const int Y4 = 1284;
        const int Y5 = 1285;
        const int Y6 = 1286;
        const int Y7 = 1287;
        const int Y10 = 1288;
        const int Y11 = 1289;
        const int Y12 = 1290;
        const int Y13 = 1291;
        const int Y20 = 1296;
        const int Y21 = 1297;
        const int Y22 = 1298;
        const int Y23 = 1299;
        const int Y24 = 1300;
        const int Y25 = 1301;
        const int Y40 = 1312;
        const int Y41 = 1313;
        const int Y42 = 1314;
        const int Y43 = 1315;
        const int Y44 = 1316;
        const int Y45 = 1317;
        const int Y46 = 1318;
        const int Y47 = 1319;
        const int Y50 = 1320;
        const int Y51 = 1321;
        const int Y52 = 1322;
        const int Y53 = 1323;
        const int Y35 = 1309;
        const int Y36 = 1310;
        const int Y37 = 1311;
        const int FCT_CARD_NUMBER = 2;
        const int FCT_STEP_MAX = 10;

        //Sıfırlanmamalı
        int totalCard = 0;
        int errorCard = 0;
        public string customMessageBoxTitle = "";
        string logDosyaPath = "";

        //Sıfırlanmalı
        int step = 0;
        int stepState = 0;
        int adminTimerCounter = 0;
        int timeoutTimerCounter = 0;
        int saniyeTimerCounter = 0;
        int fctSaniye = 0;

        //Sıfırlanmalı
        public int yetki = 0;
        public int barcodeCounter = 0;
        string[] filePathTxt = new string[FCT_CARD_NUMBER + 1];
        public string[] barcode50 = new string[FCT_CARD_NUMBER + 1];
        bool[] cardResult = new bool[FCT_CARD_NUMBER + 1];
        public string[] sap_no = new string[FCT_CARD_NUMBER + 1];

        //Ölçümler
        double firstCardDutyCycle = 0;
        double firstCardFrekans = 0;
        double firstCardTP12 = 0;
        double firstCardTP13 = 0;
        double secondCardDutyCycle = 0;
        double secondCardFrekans = 0;
        double secondCardTP12 = 0;
        double secondCardTP13 = 0;

        //Sabitler
        int dutyCycleMin = 0;
        int dutyCycleMax = 0;
        int frekansMin = 0;
        int frekansMax = 0;
        int tp12Min = 0;
        int tp12Max = 0;
        int tp13Min = 0;
        int tp13Max = 0;
        string activeSAP = "";
        int period = 0;

        //Sıfırlanmalı
        SqlConnection SQLConnection;
        bool sqlConnection = false;
        public string[] urun_id = new string[FCT_CARD_NUMBER + 1];
        public string urun_id_carrier = "";
        string urun_barkod = "";
        string son_istasyon_id = "";
        string giris_zamani = "";
        string son_istasyon_zamani = "";
        string urun_durum_no = "";
        string ariza_kodu = "";
        string tamir_edildi = "";
        string son_islem_tamamlandi = "";
        string firma_no = "";
        string urun_kodu = "";
        string panacim_kodu = "";
        string parti_no = "";
        string alan_5 = "";
        string alan_6 = "";
        string alan_7 = "";
        string pcb_barkod = "";

        const string POTA_STATION = "1";
        const string PAKETLEME_STATION = "5";
        const string ICT_STATION_ISTANBUL = "15";
        const string ICT_STATION_BOLU_1 = "19";
        const string ICT_STATION_BOLU_2 = "22";
        const string ALPPLAS_STATION_SEGER_68E = "35";      //YENİ

        const string URUN_DURUM_HURDA = "2";
        const string URUN_DURUM_BEKLETILIYOR = "3";
        const string URUN_DURUM_TAMIR_EDILECEK = "4";
        const string URUN_DURUM_PROCESS = "5";
        const string URUN_DURUM_TAMIR_EDILDI = "6";
        const string URUN_DURUM_HAZIR = "7";
        const string URUN_DURUM_SEVK_EDILECEK = "8";

        const string ARIZA_YOK = "0";
        const string CHECKSUM_HATA = "1";
        const string READ_SOFTWARE = "23";
        const string DUO_TESTI = "32";

        public bool traceabilityStatus = false;

        public Main()
        {
            this.AyarFrm = new AyarForm();
            this.AyarFrm.MainFrm = this;
            this.SifreFrm = new Sifre();
            this.SifreFrm.MainFrm = this;
            this.ProgAyarFrm = new ProgAyarForm();
            this.ProgAyarFrm.MainFrm = this;
            this.ProcessFrm = new ProcessForm();
            this.ProcessFrm.MainFrm = this;
            this.ProgramlamaFrm = new ProgramlamaForm();
            this.ProgramlamaFrm.MainFrm = this;
            InitializeComponent();
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

        [DllImport("user32.dll")]
        public static extern byte ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string ClassName, string WindowName);

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (traceabilityStatus)
            {
                if (sqlConnection)
                {
                    sqlConnection = false;
                    SQLConnection.Close();
                }
            }
            if (saniyeThread != null)
            {
                saniyeThread.Abort();
            }
            modbusClientPLC.Disconnect();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            this.ShellHwnd = Main.FindWindow("Shell TrayWnd", (string)null);
            IntPtr shellHwnd = this.ShellHwnd;
            int num1 = (int)Main.ShowWindow(this.ShellHwnd, 0);
            traceabilityStatus = Ayarlar.Default.chBoxIzlenebilirlik;

            sqlCommonConnection();
            if (sqlConnection)
            {
                settingGetInit();
                FCT_Clear();
                this.ProgramlamaFrm.programlamaInit();
                this.yetkidegistir();
                saniyeThread = new Thread(saniyeThreadFunc);
                saniyeThread.Start();
            }
            //deneme();
            //timerInit.Start();
        }

        private void settingGetInit()
        {
            this.customMessageBoxTitle = Ayarlar.Default.projectName;
            this.projectNameTxt.Text = customMessageBoxTitle;
            this.Text = customMessageBoxTitle;
            this.cardPicture.ImageLocation = Ayarlar.Default.PNGdosyayolu;
            this.logDosyaPath = Ayarlar.Default.txtLogDosya;

            modbusClientPLC = new ModbusClient(Ayarlar.Default.SerialPort2Com);
            modbusClientPLC.UnitIdentifier = 1; //Not necessary since default slaveID = 1;
            modbusClientPLC.Baudrate = Ayarlar.Default.SerialPort2Baud;   // Not necessary since default baudrate = 9600
            modbusClientPLC.Parity = Ayarlar.Default.SerialPort2Parity;
            modbusClientPLC.StopBits = Ayarlar.Default.SerialPort2stopBit;
            modbusClientPLC.ConnectionTimeout = 2000;

            this.timerAdmin.Interval = Ayarlar.Default.timerAdmin;
            this.serialRxTimeout.Interval = Ayarlar.Default.serialRxTimeout;

            if (Ayarlar.Default.chBoxSerial2)  //PLC
            {
                try
                {
                    modbusClientPLC.Connect();
                    lblStatusCom2.Text = "ON";
                    lblStatusCom2.BackColor = Color.Green;
                }
                catch (Exception ex)
                {
                    int num2 = (int)MessageBox.Show("PLC Port Hatası: " + ex.ToString());
                    lblStatusCom2.Text = "OFF";
                    lblStatusCom2.BackColor = Color.Red;
                }
            }
        }

        private void deneme()
        {/*
            //DENEME ???
            string barcode1 = "PSEGER312700D-31-05-22-S00042";
            string barcode2 = "PSEGER313000D-31-05-22-S00043";
            barcodeCounter++;
            ProgramlamaFrm.BarcodeControl(barcode1);
            barcodeCounter++;
            ProgramlamaFrm.BarcodeControl(barcode2);
            barcode50[1] = barcode1;
            barcode50[2] = barcode2;
            //DENEME*/
        }
         
        /****************************************** SQL *************************************************/
        public void sqlCommonConnection()
        {
            if (traceabilityStatus)
            {
                if (sqlConnection == false)
                {
                    try
                    {
                        string connetionString = @"Data Source=192.168.0.8\MEYER;Initial Catalog=Alpplas_Uretim_Takip;User ID=Alpplas_user;Password=Alp-User-21*";
                        SQLConnection = new SqlConnection(connetionString);
                        SQLConnection.Open();
                        ConsoleAppendLine("SQL Baglantısı Açıldı", Color.Green);
                        sqlConnection = true;
                        lblStatusSQL.Text = "ON";
                        lblStatusSQL.BackColor = Color.Green;
                    }
                    catch (Exception ex)
                    {
                        sqlConnection = false;
                        lblStatusSQL.Text = "OFF";
                        lblStatusSQL.BackColor = Color.Red;
                        ConsoleAppendLine("sqlCommonConnection Error: " + ex.Message, Color.Red);
                    }
                }
            }
            else
            {
                lblStatusSQL.Text = "OFF";
                lblStatusSQL.BackColor = Color.Red;
                sqlConnection = true;
            }
        }

        public void sqlWriteError()
        {
            sqlConnection = false;
            lblStatusSQL.Text = "OFF";
            lblStatusSQL.BackColor = Color.Red;
            ConsoleAppendLine("sqlWriteError()", Color.Red);
        }

        public bool urunlerRead(string fullproductCode)    //YENİ
        {
            if (traceabilityStatus)
            {
                sqlCommonConnection();
                if (sqlConnection)
                {
                    try
                    {
                        string sql1 = "SELECT URUN_ID, URUN_BARKOD, SON_ISTASYON_ID, GIRIS_ZAMANI, SON_ISTASYON_ZAMANI, URUN_DURUM_NO, ARIZA_KODU, TAMIR_EDILDI, SON_ISLEM_TAMAMLANDI, FIRMA_NO, URUN_KODU, PANACIM_KODU, PARTI_NO, ALAN_5, ALAN_6, ALAN_7, PCB_BARKOD FROM URUNLER WHERE URUN_BARKOD='" + fullproductCode + "'";
                        SqlCommand command1 = new SqlCommand(sql1, SQLConnection);
                        SqlDataReader dataReader1 = command1.ExecuteReader(CommandBehavior.CloseConnection);

                        bool findState = false;
                        dataReader1.Read();
                        findState = dataReader1.HasRows;
                        if (findState)
                        {
                            urun_id_carrier = Convert.ToString(dataReader1.GetValue(0));
                            urun_barkod = Convert.ToString(dataReader1.GetValue(1));
                            son_istasyon_id = Convert.ToString(dataReader1.GetValue(2));
                            giris_zamani = Convert.ToString(dataReader1.GetValue(3));
                            son_istasyon_zamani = Convert.ToString(dataReader1.GetValue(4));
                            urun_durum_no = Convert.ToString(dataReader1.GetValue(5));
                            ariza_kodu = Convert.ToString(dataReader1.GetValue(6));
                            tamir_edildi = Convert.ToString(dataReader1.GetValue(7));
                            son_islem_tamamlandi = Convert.ToString(dataReader1.GetValue(8));
                            firma_no = Convert.ToString(dataReader1.GetValue(9));
                            urun_kodu = Convert.ToString(dataReader1.GetValue(10));
                            panacim_kodu = Convert.ToString(dataReader1.GetValue(11));
                            parti_no = Convert.ToString(dataReader1.GetValue(12));
                            alan_5 = Convert.ToString(dataReader1.GetValue(13));
                            alan_6 = Convert.ToString(dataReader1.GetValue(14));
                            alan_7 = Convert.ToString(dataReader1.GetValue(15));
                            pcb_barkod = Convert.ToString(dataReader1.GetValue(16));
                            ConsoleAppendLine("Ürün Id: " + urun_id_carrier, Color.Black);
                            ConsoleAppendLine("Son İstasyon Id: " + son_istasyon_id, Color.Black);
                            ConsoleAppendLine("İlk Giriş Zamanı: " + giris_zamani, Color.Black);
                            ConsoleAppendLine("Son İstasyon Zamanı: " + son_istasyon_zamani, Color.Black);
                            ConsoleAppendLine("Ürün Durum No: " + urun_durum_no, Color.Black);
                            ConsoleAppendLine("Arıza Kodu: " + ariza_kodu, Color.Black);
                            ConsoleAppendLine("Tamir Edildi: " + tamir_edildi, Color.Black);
                            ConsoleAppendLine("Son İşlem Tamamlandı: " + son_islem_tamamlandi, Color.Black);
                            ConsoleNewLine();
                            urunDurum();
                            sonIstasyonDurum();
                            arizaDurum();

                            dataReader1.Close();
                            if (sqlConnection)
                            {
                                sqlConnection = false;
                                SQLConnection.Close();
                            }
                        }
                        else
                        {
                            dataReader1.Close();
                            if (sqlConnection)
                            {
                                sqlConnection = false;
                                SQLConnection.Close();
                            }
                            ConsoleNewLine();
                            ConsoleNewLine();
                            ConsoleAppendLine("YANLIŞ BARKOD YA DA ÜRÜN SİSTEM'DE KAYITLI DEĞİL!", Color.Red);
                            return false;
                        }

                        ConsoleNewLine();
                        ConsoleNewLine();
                        if (son_istasyon_id == POTA_STATION && urun_durum_no == URUN_DURUM_HAZIR && son_islem_tamamlandi == "True")
                        {
                            ConsoleAppendLine("ÜRÜN POTADAN'DAN GEÇMİŞ ICT'YE GİRMELİ", Color.Green);
                            return false;
                        }
                        else if (son_istasyon_id == POTA_STATION && urun_durum_no == URUN_DURUM_TAMIR_EDILDI && son_islem_tamamlandi == "True" && tamir_edildi == "True")
                        {
                            ConsoleAppendLine("ÜRÜN TAMİR'DEN GEÇMİŞ FCT'YE GİREBİLİR", Color.Green);
                            return true;
                        }
                        else if ((son_istasyon_id == ICT_STATION_BOLU_1 || son_istasyon_id == ICT_STATION_BOLU_2 || son_istasyon_id == ICT_STATION_ISTANBUL) && urun_durum_no == URUN_DURUM_HAZIR && son_islem_tamamlandi == "True")
                        {
                            ConsoleAppendLine("ÜRÜN ICT'DEN GEÇMİŞ FCT'YE GİREBİLİR", Color.Green);
                            return true;
                        }
                        else if ((son_istasyon_id == ICT_STATION_BOLU_1 || son_istasyon_id == ICT_STATION_BOLU_2 || son_istasyon_id == ICT_STATION_ISTANBUL) && (urun_durum_no == URUN_DURUM_PROCESS && urun_durum_no == URUN_DURUM_TAMIR_EDILECEK) && son_islem_tamamlandi == "False")
                        {
                            ConsoleAppendLine("ÜRÜN ICT'DEN KALMIŞ FCT'YE GİREMEZ", Color.Red);
                            return false;
                        }
                        else if ((son_istasyon_id == ALPPLAS_STATION_SEGER_68E) && urun_durum_no == URUN_DURUM_TAMIR_EDILECEK && son_islem_tamamlandi == "True")
                        {
                            ConsoleAppendLine("KART TAMİRE GİRMELİ YA DA TEKRAR FCT'YE GİREBİLİR", Color.Green);
                            return true;
                        }
                        else if ((son_istasyon_id == ALPPLAS_STATION_SEGER_68E) && urun_durum_no == URUN_DURUM_HAZIR && son_islem_tamamlandi == "True")
                        {
                            ConsoleAppendLine("ÜRÜN FCT'DEN DAHA ÖNCE GEÇTİ FCT'YE GİREBİLİR", Color.Green);
                            return true;
                        }
                        else if (son_istasyon_id == PAKETLEME_STATION)
                        {
                            ConsoleAppendLine("KART PAKETLEMEDEN GEÇMİŞ FCT-ICT'YE SOKMAYIN", Color.Orange);
                            return false;
                        }
                        else
                        {
                            ConsoleAppendLine("KART BİR ÖNCEKİ İSTASYONA GİRMELİ", Color.Red);
                            return false;
                        }
                    }
                    catch (Exception ex)
                    {
                        sqlWriteError();  //READ
                        ConsoleAppendLine("urunlerRead Error: " + ex.Message, Color.Red);
                        return false;
                    }
                }
                else
                {
                    ConsoleAppendLine("SQL BAĞLANTI KAPALI", Color.Red);
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        private void sonIstasyonDurum()
        {
            if (son_istasyon_id == POTA_STATION)
            {
                if (urun_durum_no == URUN_DURUM_HAZIR)
                {
                    ConsoleAppendLine("SON GİRDİĞİ İSTASYON: POTA", Color.Green);
                }
                else if (urun_durum_no == URUN_DURUM_TAMIR_EDILDI)
                {
                    ConsoleAppendLine("SON GİRDİĞİ İSTASYON: TAMİR", Color.Green);
                }
            }
            else if (son_istasyon_id == PAKETLEME_STATION)
            {
                ConsoleAppendLine("SON GİRDİĞİ İSTASYON: PAKETLEME", Color.Green);
            }
            else if (son_istasyon_id == ICT_STATION_BOLU_1)
            {
                ConsoleAppendLine("SON GİRDİĞİ İSTASYON: ICT-1", Color.Green);
            }
            else if (son_istasyon_id == ICT_STATION_BOLU_2)
            {
                ConsoleAppendLine("SON GİRDİĞİ İSTASYON: ICT-2", Color.Green);
            }
            else if (son_istasyon_id == ICT_STATION_ISTANBUL)
            {
                ConsoleAppendLine("SON GİRDİĞİ İSTASYON: ICT-İSTANBUL", Color.Green);
            }
            else if (son_istasyon_id == ALPPLAS_STATION_SEGER_68E)
            {
                ConsoleAppendLine("SON GİRDİĞİ İSTASYON: ALPPLAS_STATION_SEGER_68E FCT", Color.Green);
            }
        }

        private void urunDurum()
        {
            if (son_istasyon_id == PAKETLEME_STATION)
            {
                ConsoleAppendLine("ÜRÜN DURUM: ÜRÜN SEVKİYATA HAZIR", Color.Green);
            }
            else
            {
                if (urun_durum_no == URUN_DURUM_HURDA)
                {
                    ConsoleAppendLine("ÜRÜN DURUM: ÜRÜN HURDA", Color.Red);
                }
                else if (urun_durum_no == URUN_DURUM_BEKLETILIYOR)
                {
                    ConsoleAppendLine("ÜRÜN DURUM: ÜRÜN BEKLETİLİYOR", Color.Red);
                }
                else if (urun_durum_no == URUN_DURUM_TAMIR_EDILECEK)
                {
                    ConsoleAppendLine("ÜRÜN DURUM: ÜRÜN TAMİR EDİLECEK", Color.Red);
                }
                else if (urun_durum_no == URUN_DURUM_PROCESS)
                {
                    ConsoleAppendLine("ÜRÜN DURUM: ÜRÜN TEST EDİLİYOR VEYA İŞLEM ALTINDA", Color.Green);
                }
                else if (urun_durum_no == URUN_DURUM_TAMIR_EDILDI)
                {
                    ConsoleAppendLine("ÜRÜN DURUM: ÜRÜN TAMİR EDİLDİ", Color.Green);
                }
                else if (urun_durum_no == URUN_DURUM_HAZIR)
                {
                    ConsoleAppendLine("ÜRÜN DURUM: ÜRÜN BİR SONRAKİ TESTE HAZIR", Color.Green);
                }
                else if (urun_durum_no == URUN_DURUM_SEVK_EDILECEK)
                {
                    ConsoleAppendLine("ÜRÜN DURUM: ÜRÜN SEVKİYATA HAZIR", Color.Green);
                }
            }
        }

        private void arizaDurum()
        {
            if (ariza_kodu == ARIZA_YOK)
            {
                ConsoleAppendLine("ARIZA DURUM: ARIZA_YOK", Color.Green);
            }
            else if (ariza_kodu == CHECKSUM_HATA)
            {
                ConsoleAppendLine("ARIZA DURUM: CHECKSUM_HATA", Color.Red);
            }
            else if (ariza_kodu == READ_SOFTWARE)
            {
                ConsoleAppendLine("ARIZA DURUM: YAZILIM_TESTİ_HATA", Color.Red);
            }
            else if (ariza_kodu == DUO_TESTI)
            {
                ConsoleAppendLine("ARIZA DURUM: HABERLEŞME_HATA", Color.Red);
            }
        }

        public string barcodeUnıqTestIdRead(string urun_id_static)       //YENİ
        {
            string urun_test_id = "";
            sqlCommonConnection();
            if (sqlConnection)
            {
                try
                {
                    string sql5 = "SELECT URUN_TEST_ID, URUN_ID, MAKINA_NO, TEST_BASLANGIC_ZAMANI, TEST_BITIS_ZAMANI FROM URUN_TESTLER WHERE URUN_ID=" + Convert.ToInt32(urun_id_static);
                    SqlCommand command5 = new SqlCommand(sql5, SQLConnection);
                    SqlDataReader dataReader5 = command5.ExecuteReader(CommandBehavior.CloseConnection);

                    bool findState = false;
                    dataReader5.Read();
                    findState = dataReader5.HasRows;
                    if (findState)
                    {
                        while (dataReader5.Read())
                        {
                            urun_test_id = Convert.ToString(dataReader5["URUN_TEST_ID"].ToString());
                        }
                        dataReader5.Close();
                        if (sqlConnection)
                        {
                            sqlConnection = false;
                            SQLConnection.Close();
                        }
                        return urun_test_id;
                    }
                    else
                    {
                        dataReader5.Close();
                        if (sqlConnection)
                        {
                            sqlConnection = false;
                            SQLConnection.Close();
                        }
                        ConsoleNewLine();
                        ConsoleNewLine();
                        ConsoleAppendLine("YANLIŞ BARKOD YA DA ÜRÜN SİSTEM'DE KAYITLI DEĞİL !", Color.Red);
                        ConsoleAppendLine("LÜTFEN KARTI POTADAN GEÇİRİN !", Color.Red);
                        return urun_test_id;
                    }
                }
                catch (Exception ex)
                {
                    sqlWriteError();  //ID READ
                    ConsoleAppendLine("barcodeUnıqTestIdRead Error: " + ex.Message, Color.Red);
                    return urun_test_id;
                }
            }
            else
            {
                ConsoleAppendLine("SQL BAĞLANTI KAPALI", Color.Red);
                return urun_test_id;
            }
        }

        private bool urunTestlerInsert(int i)     //YENİ
        {
            if (traceabilityStatus)
            {
                sqlCommonConnection();
                if (sqlConnection)
                {
                    try
                    {
                        DateTime dt = DateTime.Now;
                        string nowYear = Convert.ToString(dt.Year);
                        string nowMonth = Convert.ToString(dt.Month);
                        string nowDay = Convert.ToString(dt.Day);
                        string nowHour = Convert.ToString(dt.Hour);
                        string nowMinute = Convert.ToString(dt.Minute);
                        string nowSecond = Convert.ToString(dt.Second);
                        string mnowSecond = Convert.ToString(dt.Millisecond);
                        string firstTime = nowYear + "-" + nowMonth + "-" + nowDay + " " + nowHour + ":" + nowMinute + ":" + nowSecond + "." + mnowSecond;
                        string sql2 = "INSERT INTO URUN_TESTLER (URUN_ID,MAKINA_NO,TEST_BASLANGIC_ZAMANI,TEST_BITIS_ZAMANI) VALUES('"
                            + urun_id[i] + "'," + "'" + ALPPLAS_STATION_SEGER_68E + "'," + "'" + firstTime + "'," + "NULL" + ")";

                        SqlCommand command2 = new SqlCommand(sql2, SQLConnection);
                        SqlDataReader dataReader2 = command2.ExecuteReader();
                        while (dataReader2.Read())
                        {
                            if (command2.ExecuteNonQuery() == 1)
                            {
                                ConsoleAppendLine("SQL Success 1", Color.Green);
                            }
                            else
                            {
                                ConsoleAppendLine("SQL Success 2", Color.Green);
                            }
                        }
                        ConsoleAppendLine("SQL Firt Insert", Color.Green);
                        dataReader2.Close();
                        if (sqlConnection)
                        {
                            sqlConnection = false;
                            SQLConnection.Close();
                        }
                        return true;
                    }
                    catch (Exception ex)
                    {
                        sqlWriteError();  //INSERT
                        ConsoleAppendLine("urunTestlerInsert Error: " + ex.Message, Color.Red);
                        return false;
                    }
                }
                else
                {
                    ConsoleAppendLine("SQL BAĞLANTI KAPALI", Color.Red);
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        private bool urunlerUpdate(int i)   //YENİ
        {
            if (traceabilityStatus)
            {
                sqlCommonConnection();
                if (sqlConnection)
                {
                    try
                    {
                        DateTime dt = DateTime.Now;
                        string nowYear = Convert.ToString(dt.Year);
                        string nowMonth = Convert.ToString(dt.Month);
                        string nowDay = Convert.ToString(dt.Day);
                        string nowHour = Convert.ToString(dt.Hour);
                        string nowMinute = Convert.ToString(dt.Minute);
                        string nowSecond = Convert.ToString(dt.Second);
                        string mnowSecond = Convert.ToString(dt.Millisecond);
                        string lastTime = nowYear + "-" + nowMonth + "-" + nowDay + " " + nowHour + ":" + nowMinute + ":" + nowSecond + "." + mnowSecond;
                        //  string lastTime = "2021-05-03 14:41:10.587";
                        string sql3 = "UPDATE URUNLER SET SON_ISTASYON_ID='" + ALPPLAS_STATION_SEGER_68E + "'" + ",URUN_DURUM_NO='" + urun_durum_no + "'"
                            + ",SON_ISLEM_TAMAMLANDI='" + "1" + "'" + ",ARIZA_KODU='" + ariza_kodu + "'"
                            + ",SON_ISTASYON_ZAMANI='" + lastTime + "'" + "WHERE URUN_ID='" + urun_id[i] + "'";
                        SqlCommand command3 = new SqlCommand(sql3, SQLConnection);
                        SqlDataReader dataReader3 = command3.ExecuteReader();
                        while (dataReader3.Read())
                        {
                            if (command3.ExecuteNonQuery() == 1)
                            {
                                ConsoleAppendLine("SQL Success 1", Color.Green);
                            }
                            else
                            {
                                ConsoleAppendLine("SQL Success 2", Color.Green);
                            }
                        }
                        ConsoleAppendLine("SQL Last Update", Color.Green);
                        dataReader3.Close();
                        if (sqlConnection)
                        {
                            sqlConnection = false;
                            SQLConnection.Close();
                        }
                        return true;
                    }
                    catch (Exception ex)
                    {
                        sqlWriteError();  //INSERT
                        ConsoleAppendLine("urunlerUpdate Error: " + ex.Message, Color.Red);
                        return false;
                    }
                }
                else
                {
                    ConsoleAppendLine("SQL BAĞLANTI KAPALI", Color.Red);
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        public bool fonksiyonTestInsert(string adimNo, string testTipi, string birim, string altLimit, string ustLimit, string olculen, string sonuc, string urun_id_static) //YENİ
        {
            if (traceabilityStatus)
            {
                string urun_test_id = "";
                urun_test_id = barcodeUnıqTestIdRead(urun_id_static);
                sqlCommonConnection();
                if (sqlConnection)
                {
                    try
                    {
                        DateTime dt = DateTime.Now;
                        string nowYear = Convert.ToString(dt.Year);
                        string nowMonth = Convert.ToString(dt.Month);
                        string nowDay = Convert.ToString(dt.Day);
                        string nowHour = Convert.ToString(dt.Hour);
                        string nowMinute = Convert.ToString(dt.Minute);
                        string nowSecond = Convert.ToString(dt.Second);
                        string mnowSecond = Convert.ToString(dt.Millisecond);
                        string firstTime = nowYear + "-" + nowMonth + "-" + nowDay + " " + nowHour + ":" + nowMinute + ":" + nowSecond + "." + mnowSecond;
                        string sql4 = "INSERT INTO FONKSIYON_TEST (URUN_TEST_ID,ADIM_NO,TEST_TIPI,BIRIM,ALT_LIMIT,UST_LIMIT,OLCULEN,SONUC,CREATE_DATE) VALUES('"
                          + urun_test_id + "'," + "'" + adimNo + "'," + "'" + testTipi + "'," + "'" + birim + "'," + "'" + altLimit + "'," + "'" + ustLimit + "'," + "'" + olculen + "'," + "'" + sonuc + "'," + "'" + firstTime + "'" + ")";
                        SqlCommand command4 = new SqlCommand(sql4, SQLConnection);
                        SqlDataReader dataReader4 = command4.ExecuteReader();
                        while (dataReader4.Read())
                        {
                            if (command4.ExecuteNonQuery() == 1)
                            {
                                ConsoleAppendLine("SQL Success 1", Color.Green);
                            }
                            else
                            {
                                ConsoleAppendLine("SQL Success 2", Color.Green);
                            }
                        }
                        ConsoleAppendLine("SQL fonksiyonTestInsert Insert", Color.Green);
                        dataReader4.Close();
                        if (sqlConnection)
                        {
                            sqlConnection = false;
                            SQLConnection.Close();
                        }
                        return true;
                    }
                    catch (Exception ex)
                    {
                        sqlWriteError();  //INSERT
                        ConsoleAppendLine("fonksiyonTestInsert Error: " + ex.Message, Color.Red);
                        return false;
                    }
                }
                else
                {
                    ConsoleAppendLine("SQL BAĞLANTI KAPALI", Color.Red);
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        public bool fctGenelInsert(int i, string sonuc)   //YENİ
        {
            if (traceabilityStatus)
            {
                string urun_test_id = "";
                urun_test_id = barcodeUnıqTestIdRead(urun_id[i]);
                sqlCommonConnection();
                if (sqlConnection)
                {
                    try
                    {
                        DateTime dt = DateTime.Now;
                        string nowYear = Convert.ToString(dt.Year);
                        string nowMonth = Convert.ToString(dt.Month);
                        string nowDay = Convert.ToString(dt.Day);
                        string nowHour = Convert.ToString(dt.Hour);
                        string nowMinute = Convert.ToString(dt.Minute);
                        string nowSecond = Convert.ToString(dt.Second);
                        string mnowSecond = Convert.ToString(dt.Millisecond);
                        string firstTime = nowYear + "-" + nowMonth + "-" + nowDay + " " + nowHour + ":" + nowMinute + ":" + nowSecond + "." + mnowSecond;
                        string sql6 = "INSERT INTO FCT_GENEL (URUN_ID,URUN_TEST_ID,SONUC,CREATE_DATE,SON_ISTASYON_ID) VALUES('"
                          + urun_id[i] + "'," + "'" + urun_test_id + "'," + "'" + sonuc + "'," + "'" + firstTime + "'," + "'" + ALPPLAS_STATION_SEGER_68E + "'" + ")";
                        SqlCommand command6 = new SqlCommand(sql6, SQLConnection);
                        SqlDataReader dataReader6 = command6.ExecuteReader();
                        while (dataReader6.Read())
                        {
                            if (command6.ExecuteNonQuery() == 1)
                            {
                                ConsoleAppendLine("SQL Success 1", Color.Green);
                            }
                            else
                            {
                                ConsoleAppendLine("SQL Success 2", Color.Green);
                            }
                        }
                        ConsoleAppendLine("SQL fctGenelInsert Insert", Color.Green);
                        dataReader6.Close();
                        if (sqlConnection)
                        {
                            sqlConnection = false;
                            SQLConnection.Close();
                        }
                        return true;
                    }
                    catch (Exception ex)
                    {
                        sqlWriteError();  //INSERT
                        ConsoleAppendLine("fctGenelInsert Error: " + ex.Message, Color.Red);
                        return false;
                    }
                }
                else
                {
                    ConsoleAppendLine("SQL BAĞLANTI KAPALI", Color.Red);
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        /****************************************** MODBUS *************************************************/
        private bool ModBusReadCoils(int address, int length)
        {
            if (modbusClientPLC.Connected)
            {
                try
                {
                    return modbusClientPLC.ReadCoils(address, length)[0];
                }
                catch
                {
                    ConsoleAppendLine("ModBus Read Coil Hatası." + address, Color.Red);
                    return false;
                }
            }
            else
            {
                ConsoleAppendLine("ModBus Kapalı Hatası." + address, Color.Red);
                return false;
            }
        }

        private void ModBusWriteSingleCoils(int address, bool state)
        {
            if (modbusClientPLC.Connected)
            {
                try
                {
                    modbusClientPLC.WriteSingleCoil(address, state);
                }
                catch
                {
                    ConsoleAppendLine("ModBus WriteSingle Coil Hatası." + address, Color.Red);
                }
            }
            else
            {
                ConsoleAppendLine("ModBus Kapalı Hatası." + address, Color.Red);
            }
        }

        private int ModBusReadHoldingRegisters(int address, int length)
        {
            if (modbusClientPLC.Connected)
            {
                try
                {
                    return modbusClientPLC.ReadHoldingRegisters(address, length)[0];
                }
                catch
                {
                    ConsoleAppendLine("ModBus ReadHoldingRegisters Coil Hatası." + address, Color.Red);
                    return 0;
                }
            }
            else
            {
                ConsoleAppendLine("ModBus Kapalı Hatası." + address, Color.Red);
                return 0;
            }
        }

        private bool ModBusReadDiscreteInputs(int address, int length)
        {
            if (modbusClientPLC.Connected)
            {
                try
                {
                    return modbusClientPLC.ReadDiscreteInputs(address, length)[0];
                }
                catch
                {
                    ConsoleAppendLine("ModBus ReadDiscreteInputs Hatası." + address, Color.Red);
                    return false;
                }
            }
            else
            {
                ConsoleAppendLine("ModBus Kapalı Hatası." + address, Color.Red);
                return false;
            }
        }

        /****************************************** INIT *************************************************/
        private void tbBarcodeCurrent_TextChanged(object sender, EventArgs e)
        {
            int maxLenght = 29;

            string barcode = tbBarcodeCurrent.Text;
            if (Convert.ToInt32(barcode.Length) >= maxLenght)
            {
                barcodeCounter++;
                if (ProgramlamaFrm.BarcodeControl(tbBarcodeCurrent.Text))
                {
                    barcode50[barcodeCounter] = tbBarcodeLast.Text;
                    btnFCTInit.BackColor = Color.Green;
                    btnFCTInit.Text = barcodeCounter + ".KART EKLENDİ!";
                    //CustomMessageBox.ShowMessage(barcodeCounter + ".Kart Eklendi!", customMessageBoxTitle, MessageBoxButtons.OK, CustomMessageBoxIcon.Information, Color.Green);
                }
                else
                {
                    barcodeCounter--;
                    btnFCTInit.BackColor = Color.Red;
                    btnFCTInit.Text = barcodeCounter + ".KART EKLENEMEDİ!";
                    //CustomMessageBox.ShowMessage(barcodeCounter + ".Kart Eklenemedi! Lütfen Başka Kart Ekleyiniz", customMessageBoxTitle, MessageBoxButtons.OK, CustomMessageBoxIcon.Information, Color.Red);
                }

                if (barcodeCounter == FCT_CARD_NUMBER)
                {
                    btnFCTInit.BackColor = Color.Green;
                    btnFCTInit.Text = "BUTONLARA BASARAK FCT TESTİNİ BAŞLAT";
                    timerInit.Start();
                }
            }
        }

        private void timerInit_Tick(object sender, EventArgs e)
        {
            if (ModBusReadCoils(M10, 1) && ModBusReadDiscreteInputs(X3, 1) && cardFailed == false)  //Güvenlik Biti ve Piston Aşağıda ise
            {
                timerInit.Stop();
                timerInit.Enabled = false;
                lblTimer1.BackColor = Color.Transparent;
                lblTimer1.Text = "OFF";
                timerEmergencyStop.Start();
                lblTimer2.BackColor = Color.Green;
                lblTimer2.Text = "ON";
                FCTInit();
            }
        }
       
        private void timerEmergencyStop_Tick_1(object sender, EventArgs e)
        {
            if (ModBusReadDiscreteInputs(X0, 1) && ModBusReadDiscreteInputs(X3, 1) == false)  //Acil Basıldı ve Piston Yukarıda ise
            {
                timerEmergencyStop.Stop();
                timerEmergencyStop.Enabled = false;
                lblTimer2.BackColor = Color.Transparent;
                lblTimer2.Text = "OFF";
                FCT_Fail();
            }
        }

        private void nextTimer_Tick(object sender, EventArgs e)
        {
            nextTimer.Stop();
            nextTimer.Enabled = false;
            stepState++;
            step++;
            ProcessFCT();
        }

        private void FCTInit()
        {
            if (modbusClientPLC.Connected)
            {
                saniyeState = true;
                Thread.Sleep(500);
                for (int i = 1; i <= FCT_CARD_NUMBER; i++)
                {
                    textCreate(i);
                    Thread.Sleep(200);
                    urunTestlerInsert(i);   //Teste Girdim
                    Thread.Sleep(200);
                }
                Thread.Sleep(500);
                btnFCTInit.BackColor = Color.Green;
                btnFCTInit.Text = "TEST BAŞLADI";
                nextTimer.Start();
            }
            else
            {
                CustomMessageBox.ShowMessage("Serial Bağlantısını Kontrol Ediniz!", customMessageBoxTitle, MessageBoxButtons.OK, CustomMessageBoxIcon.Information, Color.Red);
                ProcessFrm.ProcessFailed(1);
            }
        }

        /****************************************************FCT*******************************************************************/
        public void textCreate(int barcodeNum)
        {
            try
            {
                DateTime dt = DateTime.Now;
                string nowYear = Convert.ToString(dt.Year);
                string nowMonth = Convert.ToString(dt.Month);
                string nowDay = Convert.ToString(dt.Day);
                string nowHour = Convert.ToString(dt.Hour);
                string nowMinute = Convert.ToString(dt.Minute);
                string nowSecond = Convert.ToString(dt.Second);
                string name = barcode50[barcodeNum] + "-"+ nowYear + "-" + nowMonth + "-" + nowDay + "-" + nowHour + "-" + nowMinute + "-" + nowSecond;
                filePathTxt[barcodeNum] = logDosyaPath + "//" + name + ".txt"; //
                StreamWriter FileWrite = new StreamWriter(filePathTxt[barcodeNum]);
                FileWrite.Close();
            }
            catch (Exception ex)
            {
                ConsoleAppendLine("textCreate: " + ex.Message, Color.Red);
            }
        }

        public bool cardFailed = false;
        private void ProcessFCT()  //Sol 1.Kart Kalın(312700) - Sağ 2.Kart İnce(313000)  
        {
            if (stepState >= 3 && stepState <= 6)
            {
                activeSAP = sap_no[1];
            }
            else if (stepState >= 7 && stepState <= 10)
            {
                activeSAP = sap_no[2];
            }

            if (activeSAP == "312700")  //Kalın
            {
                dutyCycleMin = Convert.ToInt32(Prog_Ayarlar.Default.kDutyCycleMin);
                dutyCycleMax = Convert.ToInt32(Prog_Ayarlar.Default.kDutyCycleMax);
                frekansMin = Convert.ToInt32(Prog_Ayarlar.Default.kFrekansMin);
                frekansMax = Convert.ToInt32(Prog_Ayarlar.Default.kFrekansMax);
                tp12Min = Convert.ToInt32(Prog_Ayarlar.Default.kTP12Min);
                tp12Max = Convert.ToInt32(Prog_Ayarlar.Default.kTP12Max);
                tp13Min = Convert.ToInt32(Prog_Ayarlar.Default.kTP13Min);
                tp13Max = Convert.ToInt32(Prog_Ayarlar.Default.kTP13Max);
            }
            else if (activeSAP == "313000")  // İnce
            {
                dutyCycleMin = Convert.ToInt32(Prog_Ayarlar.Default.iDutyCycleMin);
                dutyCycleMax = Convert.ToInt32(Prog_Ayarlar.Default.iDutyCycleMax);
                frekansMin = Convert.ToInt32(Prog_Ayarlar.Default.iFrekansMin);
                frekansMax = Convert.ToInt32(Prog_Ayarlar.Default.iFrekansMax);
                tp12Min = Convert.ToInt32(Prog_Ayarlar.Default.iTP12Min);
                tp12Max = Convert.ToInt32(Prog_Ayarlar.Default.iTP12Max);
                tp13Min = Convert.ToInt32(Prog_Ayarlar.Default.iTP13Min);
                tp13Max = Convert.ToInt32(Prog_Ayarlar.Default.iTP13Max);
            }

            serialRxTimeout.Start();
            ProcessFrm.ProcessStart(stepState, FCT_STEP_MAX);
            if (stepState == 1) //Adım1 1. KARTIN PROGRAMLANMASI 
            {
                if (step == 1)
                {
                    logTut1(" ", "", "");
                    logTut1("1. KARTIN PROGRAMLANMASI", "", "");
                    logTut1(" ", "", "");
                    ModBusWriteSingleCoils(Y1, true);    //Birinci Kart Programlayıcı Anahtarlama
                    ModBusWriteSingleCoils(Y2, false);   //İkici Kart Programlayıcı Anahtarlama
                    Thread.Sleep(500);
                    if (ProgramlamaFrm.ProgramProduct(sap_no[1]))
                    {
                        logTut1("1.Kart Programlama Test", ":Passed ", "OK" + " Dönüş");
                        ProcessFrm.ProcessSuccess(stepState);
                    }
                    else
                    {
                        logTut1("1.Kart Programlama Test", ":Failed ", "NOK"+ " Dönüş");
                        ProcessFrm.ProcessFailed(stepState);
                    }
                }
            }
            else if (stepState == 2) //Adım2 2. KARTIN PROGRAMLANMASI 
            {
                if (step == 2)
                {
                    logTut2(" ", "", "");
                    logTut2("2. KARTIN PROGRAMLANMASI", "", "");
                    logTut2(" ", "", "");
                    ModBusWriteSingleCoils(Y1, false);  //Birinci Kart Programlayıcı Anahtarlama
                    ModBusWriteSingleCoils(Y2, true);   //İkici Kart Programlayıcı Anahtarlama
                    Thread.Sleep(500);
                    if (ProgramlamaFrm.ProgramProduct(sap_no[2]))
                    {
                        logTut2("2.Kart Programlama Test", ":Passed ", "OK" + " Dönüş");
                        ProcessFrm.ProcessSuccess(stepState);
                    }
                    else
                    {
                        logTut2("2.Kart Programlama Test", ":Failed ", "NOK" + " Dönüş");
                        ProcessFrm.ProcessFailed(stepState);
                    }
                }
            }
            else if (stepState == 3) //Adım3 1. KARTIN GÜÇ ÖLÇÜMLERİ 
            {
                if (step == 3)
                {
                    logTut1(" ", "", "");
                    logTut1("1. KARTIN GÜÇ ÖLÇÜMLERİ ", "", "");
                    logTut1(" ", "", "");
                    ModBusWriteSingleCoils(Y2, false);   //İkici Kart Programlayıcı Anahtarlama
                    ModBusWriteSingleCoils(M1, true);     // 1. Kartın Duty Cycle 
                    ModBusWriteSingleCoils(M2, false);     // 2. Kartın Duty Cycle 
                    ModBusWriteSingleCoils(M4, false);     // 2. Kartın Frekansı
                    Thread.Sleep(1000);

                    double tasiyici = ModBusReadHoldingRegisters(D10, 1);
                    Thread.Sleep(1000);
                    ModBusWriteSingleCoils(M3, true);     // 1. Kartın Duty Cycle 
                    Thread.Sleep(1500);
                    firstCardFrekans = ModBusReadHoldingRegisters(D30, 1);
                    firstCardDutyCycle = (tasiyici * 10000) / (1000000 / firstCardFrekans);
                    //ModBusWriteSingleCoils(M3, true);     // 1. Kartın Frekansı

                    if (firstCardDutyCycle >= dutyCycleMin && firstCardDutyCycle <= dutyCycleMax)
                    {
                        logTut1("1.Kart Güç Test", ":Passed ", firstCardDutyCycle + " %" + " Dönüş");
                        ProcessFrm.ProcessSuccess(stepState);
                    }
                    else
                    {
                        logTut1("1.Kart Güç Test", ":Failed ", firstCardDutyCycle  + " %" + " Dönüş");
                        ProcessFrm.ProcessFailed(stepState);
                    }
                }
            }
            else if (stepState == 4) //Adım4 1. KARTIN FREKANS ÖLÇÜMLERİ 
            {
                if (step == 4)
                {
                    logTut1(" ", "", "");
                    logTut1("1. KARTIN FREKANS ÖLÇÜMLERİ ", "", "");
                    logTut1(" ", "", "");
                    
                    if (firstCardFrekans >= frekansMin && firstCardFrekans <= frekansMax)
                    {
                        logTut1("1.Kart Frekans Test", ":Passed ", firstCardFrekans + " hz" + " Dönüş");
                        ProcessFrm.ProcessSuccess(stepState);
                    }
                    else
                    {
                        logTut1("1.Kart Frekans Test", ":Failed ", firstCardFrekans + " hz" + " Dönüş");
                        ProcessFrm.ProcessFailed(stepState);
                    }
                }
            }
            else if (stepState == 5) //Adım5 1. KARTIN TP12 ÖLÇÜMÜ
            {
                if (step == 5)
                {
                    logTut1(" ", "", "");
                    logTut1("1. KARTIN FREKANS TP12 ÖLÇÜMÜ ", "", "");
                    logTut1(" ", "", "");
                    firstCardTP12 = ModBusReadHoldingRegisters(D410, 1);
                    if (firstCardTP12 >= tp12Min && firstCardTP12 <= tp12Max)
                    {
                        logTut1("1.Kart TP12 Test", ":Passed ", firstCardTP12 / 100 + " V" + " Dönüş");
                        ProcessFrm.ProcessSuccess(stepState);
                    }
                    else
                    {
                        logTut1("1.Kart TP12 Test", ":Failed ", firstCardTP12 / 100 + " V" + " Dönüş");
                        ProcessFrm.ProcessFailed(stepState);
                    }
                }
            }
            else if (stepState == 6) //Adım6 1. KARTIN TP13 ÖLÇÜMÜ
            {
                if (step == 6)
                {
                    logTut1(" ", "", "");
                    logTut1("1. KARTIN FREKANS TP13 ÖLÇÜMÜ ", "", "");
                    logTut1(" ", "", "");
                    firstCardTP13 = ModBusReadHoldingRegisters(D415, 1);
                    if (firstCardTP13 >= tp13Min && firstCardTP13 <= tp13Max)
                    {
                        logTut1("1.Kart TP13 Test", ":Passed ", firstCardTP13 / 100 + " V" + " Dönüş");
                        ProcessFrm.ProcessSuccess(stepState);
                    }
                    else
                    {
                        logTut1("1.Kart TP13 Test", ":Failed ", firstCardTP13 / 100 + " V" + " Dönüş");
                        ProcessFrm.ProcessFailed(stepState);
                    }
                }
                ModBusWriteSingleCoils(M1, false);     // 1. Kartın Duty Cycle 
                ModBusWriteSingleCoils(M3, false);     // 1. Kartın Frekansı
                ModBusWriteSingleCoils(Y3, false);    //1.Kartın Beslemesi
            }
            else if (stepState == 7) //Adım7 2. KARTIN GÜÇ ÖLÇÜMLERİ 
            {
                if (step == 7)
                {
                    logTut2(" ", "", "");
                    logTut2("2. KARTIN GÜÇ ÖLÇÜMLERİ ", "", "");
                    logTut2(" ", "", "");
                    ModBusWriteSingleCoils(M1, false);     // 1. Kartın Duty Cycle 
                    ModBusWriteSingleCoils(M3, false);     // 1. Kartın Frekansı
                    ModBusWriteSingleCoils(M2, true);     // 2. Kartın Duty Cycle 
                    Thread.Sleep(1000);
                    double tasiyici2 = ModBusReadHoldingRegisters(D20, 1);
                    Thread.Sleep(1000);
                    ModBusWriteSingleCoils(M4, true);     // 2. Kartın Frekansı
                    Thread.Sleep(1500);
                    secondCardFrekans = ModBusReadHoldingRegisters(D40, 1);
                    secondCardDutyCycle = (tasiyici2 * 10000) / (1000000 / secondCardFrekans);
                    if (secondCardDutyCycle >= dutyCycleMin && secondCardDutyCycle <= dutyCycleMax)
                    {
                        logTut2("2.Kart Güç Test", ":Passed ", secondCardDutyCycle + " %" + " Dönüş");
                        ProcessFrm.ProcessSuccess(stepState);
                    }
                    else
                    {
                        logTut2("2.Kart Güç Test", ":Failed ", secondCardDutyCycle + " %" + " Dönüş");
                        ProcessFrm.ProcessFailed(stepState);
                    }
                }
            }
            else if (stepState == 8) //Adım8 2. KARTIN FREKANS ÖLÇÜMLERİ 
            {
                if (step == 8)
                {
                    logTut2(" ", "", "");
                    logTut2("2. KARTIN FREKANS ÖLÇÜMLERİ ", "", "");
                    logTut2(" ", "", "");
                    
                    if (secondCardFrekans >= frekansMin && secondCardFrekans <= frekansMax)
                    {
                        logTut2("2.Kart Frekans Test", ":Passed ", secondCardFrekans + " hz" + " Dönüş");
                        ProcessFrm.ProcessSuccess(stepState);
                    }
                    else
                    {
                        logTut2("2.Kart Frekans Test", ":Failed ", secondCardFrekans + " hz" + " Dönüş");
                        ProcessFrm.ProcessFailed(stepState);
                    }
                }
            }
            else if (stepState == 9) //Adım9 2. KARTIN TP12 ÖLÇÜMÜ
            {
                if (step == 9)
                {
                    logTut2(" ", "", "");
                    logTut2("2. KARTIN FREKANS TP12 ÖLÇÜMÜ ", "", "");
                    logTut2(" ", "", "");
                    secondCardTP12 = ModBusReadHoldingRegisters(D400, 1);
                    if (secondCardTP12 >= tp12Min && secondCardTP12 <= tp12Max)
                    {
                        logTut2("2.Kart TP12 Test", ":Passed ", secondCardTP12 / 100 + " V" + " Dönüş");
                        ProcessFrm.ProcessSuccess(stepState);
                    }
                    else
                    {
                        logTut2("2.Kart TP12 Test", ":Failed ", secondCardTP12 / 100 + " V" + " Dönüş");
                        ProcessFrm.ProcessFailed(stepState);
                    }
                }
            }
            else if (stepState == 10) //Adım10 2. KARTIN TP13 ÖLÇÜMÜ
            {
                if (step == 10)
                {
                    logTut2(" ", "", "");
                    logTut2("2. KARTIN FREKANS TP13 ÖLÇÜMÜ ", "", "");
                    logTut2(" ", "", "");
                    secondCardTP13 = ModBusReadHoldingRegisters(D405, 1);
                    if (secondCardTP13 >= tp13Min && secondCardTP13 <= tp13Max)
                    {
                        logTut2("2.Kart TP13 Test", ":Passed ", secondCardTP13 / 100 + " V" + " Dönüş");
                        ProcessFrm.ProcessSuccess(stepState);
                    }
                    else
                    {
                        logTut2("2.Kart TP13 Test", ":Failed ", secondCardTP13 / 100 + " V" + " Dönüş");
                        ProcessFrm.ProcessFailed(stepState);
                    }
                }
                ModBusWriteSingleCoils(M2, false);     // 2. Kartın Duty Cycle 
                ModBusWriteSingleCoils(M4, false);     // 2. Kartın Frekansı
                ModBusWriteSingleCoils(Y4, false);     //2.Kartın Beslemesi
                FCT_Success();
            }
            if (step > 0 && step < 10 && cardFailed == false)
                nextTimer.Start();
        }
         
        private void FCT_Success()
        {
            for (int i = 1; i <= FCT_CARD_NUMBER; i++)
            {
                if (urunlerUpdate(i))
                {
                    Thread.Sleep(500);
                    fctGenelInsert(i, "1");
                    //printerFunction(barcode50[i], i);
                    //printerFunction1(true, i);
                }
            }
            cardFailed = false;
            FCT_Finish();
            ModBusWriteSingleCoils(M5, true);
            CustomMessageBox.ShowMessage("FCT Testi Sonlandı. Lütfen Tekrar Başlayın!", customMessageBoxTitle, MessageBoxButtons.OK, CustomMessageBoxIcon.Information, Color.Green);
            componentsClear();
            ProcessFrm.Process_Clear();
            //deneme();
        }

        public void FCT_Fail()
        {
            for (int i = 1; i <= FCT_CARD_NUMBER; i++)
            {
                fctGenelInsert(i, "0");
                //printerFunction1(false, i);
            }
            cardFailed = true;
            errorCard = errorCard + 2;
            errorCardTxt.Text = Convert.ToString(errorCard);
            FCT_Finish();
            CustomMessageBox.ShowMessage("FCT Testi Başarısız Sonlandı. Lütfen Tekrar Başlayın!", customMessageBoxTitle, MessageBoxButtons.OK, CustomMessageBoxIcon.Information, Color.Red);
            cardFailed = false;
            ModBusWriteSingleCoils(M5, true);
            componentsClear();
            ProcessFrm.Process_Clear();
            //deneme();
        }

        public void FCT_Finish()
        {
            FCT_Clear();
            Verim();
        }

        private void FCT_Clear()
        {
            timersClear();
            variablesClear();
        }

        private void timersClear()
        {
            nextTimer.Stop();
            nextTimer.Enabled = false;
            timerAdmin.Stop();
            timerAdmin.Enabled = false;
            serialRxTimeout.Stop();
            serialRxTimeout.Enabled = false;
            timerEmergencyStop.Stop();
            timerEmergencyStop.Enabled = false;
            lblTimer2.BackColor = Color.Transparent;
            lblTimer2.Text = "OFF";
            timerInit.Start();
            lblTimer1.BackColor = Color.Green;
            lblTimer1.Text = "ON";
        }

        private void variablesClear()
        {
            saniyeState = false;
            step = 0;
            stepState = 0;
            adminTimerCounter = 0;
            timeoutTimerCounter = 0;
            saniyeTimerCounter = 0;
            fctSaniye = 0;

            yetki = 0;
            barcodeCounter = 0;
            for (int i = 1; i <= FCT_CARD_NUMBER; i++)
            {
                filePathTxt[i] = "";
                barcode50[i] = "";
                cardResult[i] = true;
                sap_no[i] = "";
            }
        }

        private void componentsClear()
        {
            btnFCTInit.BackColor = Color.Yellow;
            btnFCTInit.Text = "KART EKLEMEYE BAŞLAYABİLİRSİNİZ.";
            progressBarFCT.Value = 0;
        }
         
        private void Verim()
        {
            totalCard = totalCard + 2;
            totalCardTxt.Text = Convert.ToString(totalCard);
            verimTxt.Text = Convert.ToString(100 - ((float)((float)errorCard / totalCard)) * 100);
        }

        /****************************************************PRINTER*******************************************************************/
        private void printerFunction(object data, int cardNumber)  //PRINTER AKSİYON
        {
            try
            {
                string fullproductCode = (string)data;
                string company_no = string.Empty;
                string sap_no = string.Empty;
                string product_date = string.Empty;
                string index_no = string.Empty;
                string product_no = string.Empty;
                string card_type = string.Empty;
                string gerber_ver = string.Empty;
                string bom_ver = string.Empty;
                string hardware_ver = string.Empty;
                string software_ver = string.Empty;

                company_no = fullproductCode.Substring(0, 2);
                sap_no = fullproductCode.Substring(2, 10);
                product_date = fullproductCode.Substring(12, 4);
                index_no = fullproductCode.Substring(16, 6);
                product_no = fullproductCode.Substring(22, 14);
                card_type = fullproductCode.Substring(36, 2);
                gerber_ver = fullproductCode.Substring(38, 2);
                bom_ver = fullproductCode.Substring(40, 2);
                hardware_ver = fullproductCode.Substring(42, 4);
                software_ver = fullproductCode.Substring(46, 4);

                string test = "";
                string start = "^XA" + "^LH" + Ayarlar.Default.printerPos;
                string qr = "^BQN,2,2" + "^FDQA," + fullproductCode + "^FS";
                string s1 = company_no + index_no.Substring(0, 2);
                string s2 = index_no.Substring(2, 4);
                string s3 = product_no.Substring(0, 4);
                string s4 = product_no.Substring(4, 4);
                string s5 = product_no.Substring(8, 4);
                string s6 = product_no.Substring(12, 2) + card_type;

                string veri1 = "^FO60,10" + "^A0,15,15" + "^FD" + "P/N: " + sap_no + "^FS";   //İlki Pozisyon //İkincisi Boy-En
                string veri2 = "^FO60,35" + "^A0,15,15" + "^FD" + "S/N: " + s1 + "-" + s2 + "-" + s3 + "^FS";   //İlki Pozisyon //İkincisi Boy-En
                string veri3 = "^FO60,60" + "^A0,15,15" + "^FD" + "       " + s4 + "-" + s5 + "-" + s6 + "^FS";   //İlki Pozisyon //İkincisi Boy-En
                string veri4 = "^FO60,85" + "^A0,15,15" + "^FD" + "VER: " + hardware_ver + "." + software_ver + " G:" + gerber_ver + " B:" + bom_ver + " T:" + product_date + "^FS";   //İlki Pozisyon //İkincisi Boy-En
                string veri5 = "^FO110,110" + "^A0,30,30" + "^FD" + Convert.ToString(cardNumber) + "^FS";   //İlki Pozisyon //İkincisi Boy-En
                string end = "^XZ";

                test = start + qr + veri1 + veri2 + veri3 + veri4 + veri5 + end;

                //Get local print server
                var server = new LocalPrintServer();

                //Load queue for correct printer
                PrintQueue pq = server.GetPrintQueue(Ayarlar.Default.printerName, new string[0] { });
                PrintJobInfoCollection jobs = pq.GetPrintJobInfoCollection();
                foreach (PrintSystemJobInfo job in jobs)
                {
                    job.Cancel();
                }

                if (!RawPrinterHelper.SendStringToPrinter(Ayarlar.Default.printerName, test))
                {
                    ConsoleAppendLine("Printer Error1: ", Color.Red);
                }
            }
            catch (Exception ex)
            {
                ConsoleAppendLine("Printer Error2: " + ex.Message, Color.Red);
            }
        }

        private void printerFunction1(bool productState, int cardNumber)  //PRINTER AKSİYON
        {
            try
            {
                string test = "";
                string start = "^XA" + "^LH" + Ayarlar.Default.printerPos;
                string veri1 = "";
                if (productState)
                    veri1 = "^FO70,110" + "^A0,30,30" + "^FD" + "OK" + "^FS";   //İlki Pozisyon //İkincisi Boy-En
                else
                    veri1 = "^FO70,110" + "^A0,30,30" + "^FD" + "NOK" + "^FS";   //İlki Pozisyon //İkincisi Boy-En

                string veri2 = "^FO110,110" + "^A0,30,30" + "^FD" + Convert.ToString(cardNumber) + "^FS";   //İlki Pozisyon //İkincisi Boy-En
                string end = "^XZ";

                test = start + veri1 + veri2 + end;

                //Get local print server
                var server = new LocalPrintServer();

                //Load queue for correct printer
                PrintQueue pq = server.GetPrintQueue(Ayarlar.Default.printerName, new string[0] { });
                PrintJobInfoCollection jobs = pq.GetPrintJobInfoCollection();
                foreach (PrintSystemJobInfo job in jobs)
                {
                    job.Cancel();
                }

                if (!RawPrinterHelper.SendStringToPrinter(Ayarlar.Default.printerName, test))
                {
                    ConsoleAppendLine("Printer Error1: ", Color.Red);
                }
            }
            catch (Exception ex)
            {
                ConsoleAppendLine("Printer Error2: " + ex.Message, Color.Red);
            }
        }

        /****************************************************OTHER*******************************************************************/
        private void logTut1(string testName, string testResult, string testState)
        {
            try
            {
                if (logDosyaPath != "")
                {
                    List<string> lines = new List<string>();
                    lines = File.ReadAllLines(filePathTxt[1]).ToList();
                    lines.Add(testName + testResult + testState);
                    ConsoleAppendLine(testName + testResult + testState + "Eklendi", Color.Green);
                    File.WriteAllLines(filePathTxt[1], lines);
                }
                else
                {
                    ConsoleAppendLine("Dosya Yolu Boş Kalamaz", Color.Red);
                }
            }
            catch (Exception ex)
            {
                ConsoleAppendLine("sqlTextYaz: " + ex.Message, Color.Red);
            }
        }

        private void logTut2(string testName, string testResult, string testState)
        {
            try
            {
                if (logDosyaPath != "")
                {
                    List<string> lines = new List<string>();
                    lines = File.ReadAllLines(filePathTxt[2]).ToList();
                    lines.Add(testName + testResult + testState);
                    ConsoleAppendLine(testName + testResult + testState + "Eklendi", Color.Green);
                    File.WriteAllLines(filePathTxt[2], lines);
                }
                else
                {
                    ConsoleAppendLine("Dosya Yolu Boş Kalamaz", Color.Red);
                }
            }
            catch (Exception ex)
            {
                ConsoleAppendLine("sqlTextYaz: " + ex.Message, Color.Red);
            }
        }

        /****************************************************CONSOLE TEXT*******************************************************************/
        private void rtbConsole_TextChanged(object sender, EventArgs e)
        {
            RichTextBox rtb = sender as RichTextBox;
            rtb.SelectionStart = rtb.Text.Length;
            rtb.ScrollToCaret();
        }

        /*Kullanıcı Arayüzüne Yazı Yazılır*/
        public void ConsoleAppendLine(string text, Color color)
        {
            if (rtbConsole.InvokeRequired)
            {
                rtbConsole.Invoke(new Action(delegate ()
                {
                    rtbConsole.Select(rtbConsole.TextLength, 0);
                    rtbConsole.SelectionColor = color;
                    rtbConsole.AppendText(text + Environment.NewLine);
                    rtbConsole.Select(rtbConsole.TextLength, 0);
                    rtbConsole.SelectionColor = Color.White;
                }));
            }
            else
            {
                rtbConsole.Select(rtbConsole.TextLength, 0);
                rtbConsole.SelectionColor = color;
                rtbConsole.AppendText(text + Environment.NewLine);
                rtbConsole.Select(rtbConsole.TextLength, 0);
                rtbConsole.SelectionColor = Color.White;
            }
        }

        /*Kullanıcı Arayüzünde Bir Satır Boşluk Bırakılır*/
        public void ConsoleNewLine()
        {
            if (rtbConsole.InvokeRequired)
            {
                rtbConsole.Invoke(new Action(delegate ()
                {
                    rtbConsole.AppendText(Environment.NewLine);
                }));
            }
            else
            {
                rtbConsole.AppendText(Environment.NewLine);
            }
        }

        public void ConsoleClean()
        {
            if (rtbConsole.InvokeRequired)
            {
                rtbConsole.Invoke(new Action(delegate ()
                {
                    rtbConsole.Text = "";
                    rtbConsole.Select(rtbConsole.TextLength, 0);
                    rtbConsole.SelectionColor = Color.White;
                }));
            }
            else
            {
                rtbConsole.Text = "";
                rtbConsole.Select(rtbConsole.TextLength, 0);
                rtbConsole.SelectionColor = Color.White;
            }
        }

        /****************************************************PAGE CHANGE*******************************************************************/
        private void btnCikis_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAyar_Click(object sender, EventArgs e)
        {
            int num = (int)this.AyarFrm.ShowDialog();
        }

        private void btnProgAyar_Click(object sender, EventArgs e)
        {
            int num = (int)this.ProgAyarFrm.ShowDialog();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)    //YENİ
        {
            if (keyData != Keys.L)
                return false;
            if (this.yetki != 0)
            {
                timerAdmin.Stop();
                this.yetki = 0;
                this.yetkidegistir();
            }
            else
            {
                try { int num = (int)this.SifreFrm.ShowDialog(); }
                catch (Exception) { }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        public void yetkidegistir()
        {
            if (this.yetki == 0)
            {
                this.btnCikis.Enabled = false;
                this.btnAyar.Enabled = false;
                this.btnProgAyar.Enabled = false;

                this.btnCikis.BackColor = Color.Beige;
                this.btnAyar.BackColor = Color.Beige;
                this.btnProgAyar.BackColor = Color.Beige;
            }
            if (this.yetki == 1)
            {
                this.btnCikis.Enabled = true;
                this.btnAyar.Enabled = true;
                this.btnProgAyar.Enabled = true;

                this.btnCikis.BackColor = Color.Red;
                this.btnAyar.BackColor = Color.Red;
                this.btnProgAyar.BackColor = Color.Red;
                timerAdmin.Start();
            }
            if (this.yetki == 2)
            {
                this.btnCikis.Enabled = true;
                this.btnCikis.BackColor = Color.Red;
                this.btnAyar.BackColor = Color.Beige;
                this.btnProgAyar.BackColor = Color.Beige;
                timerAdmin.Start();
            }
        }

        private void timerAdmin_Tick_1(object sender, EventArgs e)
        {
            adminTimerCounter++;
            if (adminTimerCounter == 1)
            {
                adminTimerCounter = 0;
                timerAdmin.Stop();
                this.yetki = 0;
                this.yetkidegistir();
            }
        }

        private void serialRxTimeout_Tick(object sender, EventArgs e)
        {
            timeoutTimerCounter++;
            if (timeoutTimerCounter == 1)
            {
                ConsoleAppendLine("TIMEOUT_RX", Color.Red);
                timeoutTimerCounter = 0;
                serialRxTimeout.Stop();
                serialRxTimeout.Enabled = false;
                ProcessFrm.ProcessFailed(stepState);
            }
        }

        private void saniyeTimer_Tick(object sender, EventArgs e)
        {
            saniyeTimerCounter++;
            if (saniyeTimerCounter == 1)
            {
                saniyeTimerCounter = 0;
                fctTimerTxt.Text = Convert.ToString(++fctSaniye);
            }
        }

        bool saniyeState = false;
        int second = 0;
        int oldSecond = 0;
        private void saniyeThreadFunc()
        {
            for (; ; )
            {
                if (saniyeState)
                {
                    DateTime dt = DateTime.Now;
                    second = dt.Second;
                    if (second != oldSecond)
                    {
                        oldSecond = second;
                        fctTimerTxt.Text = Convert.ToString(++fctSaniye);
                    }
                    Thread.Sleep(1);
                }
            }
        }

        /*************************************************EXTRA**********************************************************************/
        private void tbBarcodeCurrent_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                /*
                if (ProgramlamaFrm.BarcodeControl(tbBarcodeCurrent.Text))
                {
                    sqlUpdateInitFCT();
                    ariza_kodu = "0";
                    urun_durum_no = "5";
                    tamir_edildi = "";
                    sqlUpdateLastFCT();
                }
                */
            }
        }

    }
}

