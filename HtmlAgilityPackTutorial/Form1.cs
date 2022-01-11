using System;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.IO;

namespace HtmlAgilityPackTutorial
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<string> UrlList = new List<string>();
        List<Urun> UrunList = new List<Urun>();
        OleDbConnection con;
        OleDbCommand cmd;
        OleDbDataReader dr;

        private void Form1_Load(object sender, EventArgs e)
        {
            var baseUrl = "https://www.markastok.com";
            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;//Türkçe Karakter Sorununu Gidermek İçin
            //string html = client.DownloadString(baseUrl);

            con = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\Users\ibrah\Desktop\HtmlAgilityPackTutorial - New\URL.xlsx; Extended Properties ='Excel 12.0 xml; HDR = YES; '");
            cmd = new OleDbCommand("SELECT * FROM [URL$]", con);
            con.Open();
            dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                UrlList.Add(dr["url"].ToString());
            }
            con.Close();

            foreach (var kisi in UrlList)
            {
                listBox4.Items.Add(baseUrl + kisi);
                string html = client.DownloadString(baseUrl + kisi);
                //Uri newUrl = new Uri(baseUrl + kisi);
                HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument(); dokuman.LoadHtml(html);
                string baslik = dokuman.DocumentNode.SelectNodes("//*[@id=\"product-name\"]")[0].InnerText;
                listBox1.Items.Add(baslik);
                string aciklama = dokuman.DocumentNode.SelectNodes("//*[@id=\"productRight\"]/div/div[6]/div[2]")[0].InnerText;
                listBox2.Items.Add(aciklama);
                string beden = dokuman.DocumentNode.SelectNodes("//*//*[@id=\"productRight\"]/div/div[4]/div[2]/div[2]/div")[0].InnerText;
                listBox3.Items.Add(beden);
                UrunList.Add(new Urun
                {
                    Baslik = baslik,
                    Aciklama = aciklama,
                    Beden = beden,
                    Link = baseUrl + kisi
                });
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

            for (int i = 1; i <= UrunList.Count; i++)
            {
                ((Range)sheet1.Cells[i, 1]).Value2 = UrunList[i - 1].Baslik;
                ((Range)sheet1.Cells[i, 2]).Value2 = UrunList[i - 1].Aciklama;
                ((Range)sheet1.Cells[i, 3]).Value2 = UrunList[i - 1].Beden;
            }

            //Mail Gönderme
            try
            {                
                Mail mail = new Mail();
                mail.konu = "Test Konu";
                mail.alici = "ibrahim.okuyucu@setup34.com.tr";
                mail.icerik = "Test İçerik";
                mail.attachKonum = @"C:\Users\ibrah\Desktop\HtmlAgilityPackTutorial - New\URL.xlsx";
                mail.Gonder();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void listBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}