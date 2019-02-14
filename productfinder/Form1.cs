using System;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using HtmlAgilityPack;

namespace productfinder
{
    using Excel = Microsoft.Office.Interop.Excel;
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           // var webGet = new HtmlWeb();

            //WebClient client = new WebClient();
            //client.Proxy = new WebProxy("201.236.143.242:56104", false);
            //if (textBox2.Text != "")
            //{
            //    WebClient client = new WebClient();
            //    // client.Proxy = new WebProxy('"' + textBox2.Text + '"', false);

            //    // client.Method = "GET";
            //}
            string output = textBox1.Text.Replace(" ", "+");
            WebClient wc = new WebClient();
            wc.Proxy = new WebProxy("201.236.143.242:56104");
            var page = wc.DownloadString("https://www.aliexpress.com/wholesale?SortType=create_desc&SearchText=" + '"' + output + '"' + "&page=" + 1 + "&isFavorite=y");

            var webGet = new HtmlWeb();



            var doc1 = webGet.Load(page); //webGet.Load("https://www.aliexpress.com/wholesale?SortType=create_desc&SearchText=" + '"' + output + '"' + "&page=" + 1 + "&isFavorite=y");
            var tag1 = doc1.DocumentNode.SelectNodes("//*[contains(@class,'ui-pagination-navi util-left')]");
            int k = 0;
            try
            {
                foreach (var node in tag1)
                {
                    string str = node.InnerText.Trim();
                    str = Regex.Replace(str, @"\s", "");
                    char ch = str[str.Length - 5];
                    str = Convert.ToString(ch);
                    k = Convert.ToInt32(str);


                }
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show("Change IP");
            }
            progressBar1.Maximum = k-1;
            progressBar1.Step = 1;
            progressBar1.Value = 0;
            for (int j = 0; j < k; j++) { 
            var doc = webGet.Load("https://www.aliexpress.com/wholesale?SortType=create_desc&SearchText=" + '"' + output + '"'+ "&page="+ j +"&isFavorite=y");

                if (j == 0)
                {
                    var tag = doc.DocumentNode.SelectNodes("//*[contains(@class,'pop-keyword-diva')]");
                    try { 
                    foreach (var node in tag)
                    {

                            label1.Text = node.InnerText;
                    }
                }
            catch (NullReferenceException ex)
                {
                    MessageBox.Show("Change IP");
                }

            }
            int i=0;
                
                var item = doc.DocumentNode.SelectNodes("//div[contains(@class,'info')]");
                try
                {
                    foreach (var node in item)
                    {
                        try
                        {
                            dataGridView1.Rows.Add(new object[] { node.SelectSingleNode(".//*[contains(@class,'history-item product')]").InnerText, node.SelectSingleNode(".//*[contains(@class,'history-item product')]").Attributes["href"].Value.Remove(0,2), node.SelectSingleNode(".//*[contains(@class,'order-num-a')]").InnerText.Replace("Orders (","").Replace(")","").Replace("Order (", ""), node.SelectSingleNode(".//*[contains(@class,'rate-num')]").InnerText.Replace("(", "").Replace(")",""), Regex.Replace(node.SelectSingleNode(".//*[contains(@class,'price price-m')]").InnerText.Replace("piece","").Replace("/",""), @"\s", "") });
                        }
                        catch (NullReferenceException ex)
                        {
                           
                        }
                        i = i + 1;
                    }
                }
                catch (NullReferenceException ex)
                {
                    MessageBox.Show("Not all Products are selected. Change IP");
                }
                progressBar1.Value = j;
                
                
            }
            button3.Visible = true;
            button2.Visible = true;
    }

        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                copyAlltoClipboard();
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Install MS Excel first and then try");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button3.Visible = false;
            button2.Visible = false;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.ColumnHeadersVisible = false;
            dataGridView1.Rows.Add("Product Name", "Product Link", "Number of Orders", "Number of Feedback", "Price");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form1 NewForm = new Form1();
            NewForm.Show();
            this.Dispose(false);
        }
    }
}
