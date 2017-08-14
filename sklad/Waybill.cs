using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace sklad
{
    public partial class Waybill : Form
    {
        int responsible, product, unit, no;
        string waybill, user, traffic, product_name, unit_name, date;

        private void button1_Click(object sender, EventArgs e)
        {            
            Excel.Application ExcelApp = new Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Columns.ColumnWidth = 15;
            Excel.Worksheet workSheet = (Excel.Worksheet)ExcelApp.ActiveSheet;
            Excel.Range range = workSheet.Range["B1", System.Type.Missing];
            range = workSheet.Range["A1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 11;
            range = workSheet.Range["B1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 37;
            range = workSheet.Range["C1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 8;
            range = workSheet.Range["D1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 12;
            range = workSheet.Range["E1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 12;
            range = workSheet.Range["G1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 18;
            //-------
            workSheet.get_Range("A1:B1").Merge();
            workSheet.get_Range("A2:B2").Merge();
            workSheet.get_Range("F1:G1").Merge();
            workSheet.get_Range("F2:G2").Merge();
            workSheet.get_Range("F3:G3").Merge();
            workSheet.get_Range("A4:B4").Merge();
            workSheet.get_Range("F4:G4").Merge();
            workSheet.get_Range("B8:G8").Merge();
            workSheet.get_Range("B9:G9").Merge();
            workSheet.get_Range("B10:D10").Merge();
            workSheet.get_Range("F10:G10").Merge();
            workSheet.get_Range("B11:D11").Merge();
            workSheet.get_Range("F11:G11").Merge();
            //-------
            workSheet.get_Range("A12:A13").Merge();
            workSheet.get_Range("B12:B13").Merge();
            workSheet.get_Range("C12:C13").Merge();
            workSheet.get_Range("D12:E12").Merge();
            workSheet.get_Range("F12:F13").Merge();
            workSheet.get_Range("G12:G13").Merge();
            //-------
            workSheet.get_Range("C25:D25").Merge();
            workSheet.get_Range("E26:G26").Merge();
            workSheet.get_Range("C27:D27").Merge();
            workSheet.get_Range("A28:B28").Merge();
            workSheet.get_Range("C28:D28").Merge();
            workSheet.get_Range("E28:G28").Merge();
            //-------
            range = workSheet.Range["B2", System.Type.Missing];
            range.EntireRow.RowHeight = 21;
            range = workSheet.Range["A10", System.Type.Missing];
            range.EntireRow.RowHeight = 30;
            ExcelApp.Cells[1, 1] = "'LGÇ' Müdirliginiň Guýulary düýpli we ýerasty bejeriji bölegi";
            ExcelApp.Cells[1, 6] = "A-5 görnüş";
            ExcelApp.Cells[2, 1] = "Kärhananyň ady";
            ExcelApp.Cells[2, 6] = "Türkmenistanyň Maliýe ministrliginiň";
            ExcelApp.Cells[3, 6] = "2011-nji ýylyň 19 awgustyndaky";
            ExcelApp.Cells[4, 1] = "Düzümindäki bölüm";
            ExcelApp.Cells[4, 6] = "82 belgili buýrugy bilen tassyklanyldy";
            ExcelApp.Cells[6, 3] = "Talapnama - ýan haty № "+waybill;
            if (traffic == "0")
            {
                ExcelApp.Cells[8, 2] = user;
                ExcelApp.Cells[10, 2] = "Аннаклычев Хакнепес Амангелдиевич";
            }
            else if (traffic == "1")
            {
                ExcelApp.Cells[8, 2] = "Аннаклычев Хакнепес Амангелдиевич";
                ExcelApp.Cells[10, 2] = user;
            }
            ExcelApp.Cells[8, 1] = "Kimiň üsti bilen ";
            ExcelApp.Cells[9, 2] = "harydy göýberijiniň  ady familiýasy we doly resmi salgysy ";
            ExcelApp.Cells[10, 1] = "Talap eden";
            ExcelApp.Cells[10, 5] = "Rugsat beren";
            ExcelApp.Cells[11, 2] = "harydy alyjynyň  ady familiýasy we doly resmi salgysy ";
            ExcelApp.Cells[11, 6] = "Rugsat beren ýolbaşçy";
            //-------
            ExcelApp.Cells[12, 1] = "sanaw belgisi";
            ExcelApp.Cells[12, 2] = "Maddy gymmatlyklarynyň ady";
            ExcelApp.Cells[12, 3] = "Ölçeg birligi";
            ExcelApp.Cells[12, 4] = "Mukdary";
            ExcelApp.Cells[13, 4] = "talap edileni";
            ExcelApp.Cells[13, 5] = "göýberileni";
            ExcelApp.Cells[12, 6] = "Bahasy - manat, teňňe";
            ExcelApp.Cells[12, 7] = "Goşmaça gymmaty üçin salgydy hasaba almazdan - manat, teňňe";
            //-------
            ExcelApp.Cells[25, 1] = "Göýbermäge rugsat berdim.Başlyk __________________________";
            ExcelApp.Cells[25, 5] = "Baş buhgalter";
            ExcelApp.Cells[26, 3] = "ady familiýasy";
            ExcelApp.Cells[26, 5] = "   wezipesi                             goly                         ady familiýasy";
            if (traffic == "0")
            {
                ExcelApp.Cells[27, 1] = "Göýberdim: "+user;
                ExcelApp.Cells[27, 5] = "Aldym: Аннаклычев Хакнепес";
            }
            else if (traffic == "1")
            {
                ExcelApp.Cells[27, 1] = "Göýberdim: Аннаклычев Хакнепес";
                ExcelApp.Cells[27, 5] = "Aldym: "+ user;
            }

            ExcelApp.Cells[27, 1] = "Göýberdim:  __________________      __________________________";
            ExcelApp.Cells[27, 5] = "Aldym: ___________  ________________";
            ExcelApp.Cells[28, 1] = "                                                     wezipesi                                   goly           ";
            ExcelApp.Cells[28, 3] = "ady familiýasy";
            ExcelApp.Cells[28, 5] = "   wezipesi                             goly                         ady familiýasy";
            //-------
            string dg = " ";
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    dg = dataGridView1[i, j].FormattedValue.ToString();
                    ExcelApp.Cells[j + 14, i + 1] = dg;
                }
            }
            workSheet.get_Range("A1:B1").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A3:B3").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("B8:G8").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("B10:D10").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("F10:G10").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("F25").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A12:G13").Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A12:G13").Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A12:G13").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A12:G13").Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A12:G13").Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A12:G13").Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //-------
            workSheet.get_Range("A14:G22").Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A14:G22").Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A14:G22").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A14:G22").Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A14:G22").Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A14:G22").Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //-------
            workSheet.get_Range("F1:G4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("A2").Font.Size = 14;
            workSheet.get_Range("A2").Font.Superscript= true;
            workSheet.get_Range("A2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("A4").Font.Size = 14;
            workSheet.get_Range("A4").Font.Superscript = true;
            workSheet.get_Range("A4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("C6").Font.Bold = true;
            workSheet.get_Range("B9").Font.Size = 14;
            workSheet.get_Range("B9").Font.Superscript = true;
            workSheet.get_Range("B9").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("B11").Font.Size = 14;
            workSheet.get_Range("B11").Font.Superscript = true;
            workSheet.get_Range("B11").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("F11").Font.Size = 14;
            workSheet.get_Range("F11").Font.Superscript = true;
            workSheet.get_Range("F11").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("A26").Font.Size = 14;
            workSheet.get_Range("A26").Font.Superscript = true;
            workSheet.get_Range("A26").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("C26").Font.Size = 14;
            workSheet.get_Range("C26").Font.Superscript = true;
            workSheet.get_Range("C26").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("E26").Font.Size = 14;
            workSheet.get_Range("E26").Font.Superscript = true;
            workSheet.get_Range("E26").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("A28").Font.Size = 14;
            workSheet.get_Range("A28").Font.Superscript = true;
            workSheet.get_Range("A28").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("C28").Font.Size = 14;
            workSheet.get_Range("C28").Font.Superscript = true;
            workSheet.get_Range("C28").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("E28").Font.Size = 14;
            workSheet.get_Range("E28").Font.Superscript = true;
            workSheet.get_Range("E28").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("A12:G22").WrapText = true;
            workSheet.get_Range("A14:A22").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("C12:G22").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("C12:G22").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            workSheet.get_Range("A12:G13").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("A12:G13").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            workSheet.get_Range("A10:G10").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("A10:G10").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            ExcelApp.Visible = true;
        }

        float product_quantity, price, summa;
        public Waybill(string _waybill, string _traffic)
        {
            InitializeComponent();
            this.waybill = _waybill;
            this.traffic = _traffic;
        }

        private void Waybill_Load(object sender, EventArgs e)
        {
            

            var columnNo = new DataGridViewColumn();
            columnNo.HeaderText = "No";
            columnNo.Name = "nomer";
            columnNo.CellTemplate = new DataGridViewTextBoxCell();

            var columnPName = new DataGridViewColumn();
            columnPName.HeaderText = "Maddy gymmatlygyn ady";
            columnPName.Name = "productName";
            columnPName.CellTemplate = new DataGridViewTextBoxCell();

            var columnUnit = new DataGridViewColumn();
            columnUnit.HeaderText = "Olceg birligin ady";
            columnUnit.Name = "unitName";
            columnUnit.CellTemplate = new DataGridViewTextBoxCell();

            var columnMT = new DataGridViewColumn();
            columnMT.HeaderText = "Talap edileni";
            columnMT.Name = "talap";
            columnMT.CellTemplate = new DataGridViewTextBoxCell();

            var columnMG = new DataGridViewColumn();
            columnMG.HeaderText = "goyberileni";
            columnMG.Name = "goyber";
            columnMG.CellTemplate = new DataGridViewTextBoxCell();

            var columnPrice = new DataGridViewColumn();
            columnPrice.HeaderText = "Bahasy";
            columnPrice.Name = "price";
            columnPrice.CellTemplate = new DataGridViewTextBoxCell();

            var columnAdditional = new DataGridViewColumn();
            columnAdditional.HeaderText = "Gosmaca gymmaty";
            columnAdditional.Name = "additional";
            columnAdditional.CellTemplate = new DataGridViewTextBoxCell();

            dataGridView1.Columns.Add(columnNo);
            dataGridView1.Columns.Add(columnPName);
            dataGridView1.Columns.Add(columnUnit);
            dataGridView1.Columns.Add(columnMT);
            dataGridView1.Columns.Add(columnMG);
            dataGridView1.Columns.Add(columnPrice);
            dataGridView1.Columns.Add(columnAdditional);
            //-------
            this.Text = "Talapnama - yan haty No "+waybill;
            ConnOpen respLoad = new ConnOpen();
            ConnOpen productLoad = new ConnOpen();
            ConnOpen unitLoad = new ConnOpen();
            ConnOpen userLoad = new ConnOpen();
            ConnOpen respForDGV = new ConnOpen();
            respLoad.connection.Open();
            SqlCommand commandResp = new SqlCommand("SELECT * FROM dbo.Responsibility WHERE waybill = '" + waybill+"' AND traffic = '"+traffic+"'", respLoad.connection);
            SqlDataReader readerResp = commandResp.ExecuteReader();
            readerResp.Read();
            responsible = Convert.ToInt32(readerResp["responsible"]);
            readerResp.Close();
            respLoad.connection.Close();
            //-------
            userLoad.connection.Open();
            SqlCommand commandUser = new SqlCommand("SELECT * FROM dbo.Users WHERE user_id = '"+responsible+"'",userLoad.connection);
            SqlDataReader readerUser = commandUser.ExecuteReader();
            readerUser.Read();        
            user = readerUser["fio"].ToString();
            readerUser.Close();
            userLoad.connection.Close();
            //-------
            respForDGV.connection.Open();
            productLoad.connection.Open();
            unitLoad.connection.Open();
            SqlCommand commandForDGV = new SqlCommand("SELECT * FROM dbo.Responsibility WHERE waybill = '" + waybill + "' AND traffic = '" + traffic + "'", respForDGV.connection);
            SqlDataReader readerForDGV = commandForDGV.ExecuteReader();
            SqlCommand commandProduct;
            SqlDataReader readerProduct;
            SqlCommand commandUnit;
            SqlDataReader readerUnit;
            int n=0;
            while(readerForDGV.Read())
            {
                n++;
                product = Convert.ToInt32(readerForDGV["product"]);
                commandProduct = new SqlCommand("SELECT * FROM dbo.Product WHERE product_id = '"+product+"'", productLoad.connection);
                readerProduct = commandProduct.ExecuteReader();
                readerProduct.Read();
                unit = Convert.ToInt32(readerProduct["product_unit"]);
                product_name = readerProduct["product_name"].ToString();
                price = Convert.ToSingle(readerProduct["product_price"]);
                readerProduct.Close();
                commandUnit = new SqlCommand("SELECT * FROM dbo.Unit WHERE unit_id = '" + unit + "'", unitLoad.connection);
                readerUnit = commandUnit.ExecuteReader();
                readerUnit.Read();
                unit_name = readerUnit["unit_name"].ToString();
                readerUnit.Close();
                product_quantity =Convert.ToSingle( readerForDGV["product_quantity"]);
                summa = product_quantity * price;
                dataGridView1.Rows.Add(n, product_name, unit_name, " ", product_quantity, price, summa);

            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            respForDGV.connection.Close();
            productLoad.connection.Close();
            unitLoad.connection.Close();
            //-------
            label1.Text = "Kimin usti bilen \t";
            label2.Text = "Talap eden \t";
            if(traffic=="0")
            {
                label2.Text += "Аннаклычев Хакнепес Амангелдиевич";
                label1.Text += user;
            }
            else if(traffic=="1")
            {
                label1.Text += "Аннаклычев Хакнепес Амангелдиевич";
                label2.Text += user;
            }
            
        }
    }
}
