﻿using System;
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
    public partial class Report : Form
    {
        Excel.Application ExcelApp = new Excel.Application();
        public Report()
        {
            InitializeComponent();
        }
        int product, unit;

        private void button2_Click(object sender, EventArgs e)
        {
            ConnOpen reportLoadM = new ConnOpen();
            reportLoadM.connection.Open();
            SqlCommand cReportM = new SqlCommand("SELECT DISTINCT product WHERE date<"+dateTimePicker1.Value.ToString(), reportLoadM.connection);
            reportLoadM.connection.Close();
            //-------
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Rows.RowHeight = 15;
            Excel.Worksheet workSheet = (Excel.Worksheet)ExcelApp.ActiveSheet;
            //-------
            Excel.Range range = workSheet.Range["A2", System.Type.Missing];
            range.EntireRow.RowHeight = 30;
            range = workSheet.Range["A3", System.Type.Missing];
            range.EntireRow.RowHeight = 45;
            range = workSheet.Range["A4", System.Type.Missing];
            range.EntireRow.RowHeight = 30;
            //-------
            range = workSheet.Range["A1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 4;
            range = workSheet.Range["B1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 25;
            range = workSheet.Range["C1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 6;
            range = workSheet.Range["D1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 11;
            range = workSheet.Range["E1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 8;
            range = workSheet.Range["F1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 11;
            //-------
            workSheet.get_Range("A2:A4").Merge();
            workSheet.get_Range("B2:B4").Merge();
            workSheet.get_Range("C2:C4").Merge();
            workSheet.get_Range("D2:D4").Merge();
            workSheet.get_Range("E2:F3").Merge();
            workSheet.get_Range("A26:B26").Merge();
            //-------
            int i = 1, g = 0;
            while (i<22)
            {
                g = i + 4;
                ExcelApp.Cells[g, 1] = i;
                i++;
            }
            ExcelApp.Cells[2, 1] = "T №";
            ExcelApp.Cells[2, 2] = "MADDY                                                                                  GYMMATLYKLARYŇ                                                                              ADY";
            ExcelApp.Cells[2, 3] = "Ölçeg birligi";
            ExcelApp.Cells[2, 4] = "Bahasy";
            ExcelApp.Cells[2, 5] = dateTimePicker1.Value.ToShortDateString()+ " ý              galyndysy";
            ExcelApp.Cells[4, 5] = "sany";
            ExcelApp.Cells[4, 6] = "jemi bahasy";
            ExcelApp.Cells[26, 1] = "Jemi";
            //-------
            workSheet.get_Range("A2").Orientation = 90;
            workSheet.get_Range("B2").Font.Bold = true;
            workSheet.get_Range("E2").Font.Bold = true;
            workSheet.get_Range("A26").Font.Bold = true;
            workSheet.get_Range("A26").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("A2:F4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("A2:F4").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            workSheet.get_Range("A2:F4").WrapText = true;
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //-------
            ExcelApp.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Otchet f1 = new Otchet(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            f1.ShowDialog();

        }

        private void Report_Load(object sender, EventArgs e)
        {

            string sProduct, sUnit;
            float product_quantity, price,  pValue = 0, rValue = 0, sum, sPrice;
            ConnOpen reportLoad = new ConnOpen();
            ConnOpen productLoad = new ConnOpen();
            ConnOpen unitLoad = new ConnOpen();
            ConnOpen userLoad = new ConnOpen();
            ConnOpen tLoad = new ConnOpen();
            reportLoad.connection.Open();
            productLoad.connection.Open();
            unitLoad.connection.Open();
            userLoad.connection.Open();
            tLoad.connection.Open();
            //Открыли все коннекты
            SqlCommand commandResponsible = new SqlCommand("SELECT * FROM dbo.Responsibility", reportLoad.connection);
            SqlDataReader readerResponsible = commandResponsible.ExecuteReader();
            SqlCommand commandProduct = new SqlCommand("SELECT * FROM dbo.Product WHERE product_flag = '" + 1 + "'", productLoad.connection);
            SqlDataReader readerProduct = commandProduct.ExecuteReader();
            SqlCommand commandUnit;
            SqlDataReader readerUnit;
            SqlCommand commandT;
            SqlDataReader readerT;
            //Создали команды и датаридеры
            var columnPName = new DataGridViewColumn();
            columnPName.HeaderText = "Название";
            columnPName.Name = "productName";
            columnPName.CellTemplate = new DataGridViewTextBoxCell();
            
            var columnPUnit = new DataGridViewColumn();
            columnPUnit.HeaderText = "Ед. Изм.";
            columnPUnit.Name = "productUnit";
            columnPUnit.CellTemplate = new DataGridViewTextBoxCell();

            var columnPPrice = new DataGridViewColumn();
            columnPPrice.HeaderText = "Цена за ед.";
            columnPPrice.Name = "productPrice";
            columnPPrice.CellTemplate = new DataGridViewTextBoxCell();

            var prih = new DataGridViewColumn();
            prih.HeaderText = "Приход";
            prih.Name = "prihod";
            prih.CellTemplate = new DataGridViewTextBoxCell();

            var rash = new DataGridViewColumn();
            rash.HeaderText = "Расход";
            rash.Name = "rashod";
            rash.CellTemplate = new DataGridViewTextBoxCell();

            var ostatok = new DataGridViewColumn();
            ostatok.HeaderText = "Остаток на ";
            ostatok.Name = "ostatok";
            ostatok.CellTemplate = new DataGridViewTextBoxCell();

            var summa = new DataGridViewColumn();
            summa.HeaderText = "Остаток";
            summa.Name = "summa";
            summa.CellTemplate = new DataGridViewTextBoxCell();

            var priceSumma = new DataGridViewColumn();
            priceSumma.HeaderText = "Итого цена";
            priceSumma.Name = "sumPrice";
            priceSumma.CellTemplate = new DataGridViewTextBoxCell();

            dataGridView1.Columns.Add(columnPName);
            dataGridView1.Columns.Add(columnPUnit);
            dataGridView1.Columns.Add(ostatok);
            dataGridView1.Columns.Add(columnPPrice);
            dataGridView1.Columns.Add(prih);
            dataGridView1.Columns.Add(rash);
            dataGridView1.Columns.Add(summa);
            dataGridView1.Columns.Add(priceSumma);
            //Добавили постоянные колонки
            while (readerProduct.Read())
            {
                product = Convert.ToInt32(readerProduct["product_id"]);
                sProduct = readerProduct["product_name"].ToString();
                unit = Convert.ToInt32(readerProduct["product_unit"]);
                price = Convert.ToSingle(readerProduct["product_price"]);
                product_quantity = Convert.ToSingle(readerProduct["product_quantity"]);
                //-------
                commandUnit = new SqlCommand("SELECT * FROM dbo.Unit WHERE unit_id = '" + unit + "'", unitLoad.connection);
                readerUnit = commandUnit.ExecuteReader();
                readerUnit.Read();
                sUnit = readerUnit["unit_name"].ToString();
                readerUnit.Close();
                //-------
                commandT = new SqlCommand("SELECT * FROM dbo.Responsibility WHERE product = '"+product+"'", tLoad.connection);
                readerT = commandT.ExecuteReader();
                while(readerT.Read())
                {
                    if(readerT["traffic"].ToString() == "0")
                    {
                        pValue += Convert.ToInt32(readerT["product_quantity"]);
                    }
                    if (readerT["traffic"].ToString() == "1")
                    {
                        rValue += Convert.ToInt32(readerT["product_quantity"]);
                    }
                }
                readerT.Close();
                sum = product_quantity + pValue - rValue;
                sPrice = sum * price;
                dataGridView1.Rows.Add(sProduct, sUnit, product_quantity, price, pValue, rValue, sum, sPrice);
                pValue = 0;
                rValue = 0;
            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            //while (readerResponsible.Read())
            //{
            //    product = Convert.ToInt32(readerResponsible["product"]);
            //    responsible = Convert.ToInt32(readerResponsible["responsible"]);
            //    traffic = Convert.ToInt32(readerResponsible["traffic"]);
            //    waybill = readerResponsible["waybill"].ToString();
            //    rProduct_quantity = Convert.ToSingle(readerResponsible["product_quantity"]);
            //    //-----
            //    commandProduct = new SqlCommand("SELECT * FROM dbo.Product WHERE product_id LIKE '%"+product+"'", productLoad.connection);
            //    readerProduct = commandProduct.ExecuteReader();
            //    readerProduct.Read();
            //    sProduct = readerProduct["product_name"].ToString();
            //    unit =Convert.ToInt32(readerProduct["product_unit"]);
            //    price = Convert.ToSingle(readerProduct["product_price"]);
            //    product_quantity = Convert.ToSingle(readerProduct["product_quantity"]);
            //    readerProduct.Close();
            //    //------
            //    commandUnit = new SqlCommand("SELECT * FROM dbo.Unit WHERE unit_id LIKE '%" + unit+"'", unitLoad.connection);
            //    readerUnit = commandUnit.ExecuteReader();
            //    readerUnit.Read();
            //    sUnit = readerUnit["unit_name"].ToString();
            //    readerUnit.Close();
            //    //------
            //    commandUser = new SqlCommand("SELECT * FROM dbo.Users WHERE user_id LIKE '%" + responsible + "'", userLoad.connection);
            //    readerUser = commandUser.ExecuteReader();
            //    readerUser.Read();
            //    sResponsible = readerUser["user_familija"].ToString() + " " + readerUser["user_imja"].ToString() + " " + readerUser["user_otchestvo"].ToString();
            //    readerUser.Close();
            //    this.Text += sProduct + sUnit;

            //    dataGridView1.Rows.Add(sProduct, sUnit, price);
            //}

            reportLoad.connection.Close();
            productLoad.connection.Close();
            unitLoad.connection.Close();
            userLoad.connection.Close();
            tLoad.connection.Close();
            //Закрыли коннекты

        }
    }
}
