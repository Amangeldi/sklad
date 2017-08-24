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
    public partial class Report : Form
    {
        Excel.Application ExcelApp = new Excel.Application();
        public Report()
        {
            InitializeComponent();
        }
        int product, unit;
        string[] c = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z","AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ","BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ" };
        string[] gh = { "G3:H3", "I3:J3", "K3:L3", "M3:N3", "O3:P3", "Q3:R3", "S3:T3", "U3:V3", "W3:X3", "Y3:Z3", "AA3:AB3", "AC3:AD3", "AE3:AF3", "AG3:AH3", "AI3:AJ3", "AK3:AL3", "AM3:AN3", "AO3:AP3", "AQ3:AR3", "AS3:AT3", "AU3:AV3", "AW3:AX3", "AY3:AZ3", "BA3:BB3", "BC3:BD3", "BE3:BF3" };

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Rows.RowHeight = 15;
            Excel.Worksheet workSheet = (Excel.Worksheet)ExcelApp.ActiveSheet;
            //-------
            string sProduct, sUnit, cell, resp = "", fi = "";
            float product_quantity, price, pValue = 0, rValue = 0, sum, sPrice;
            int girdeji = 0, cykdajyb = 0, cykdajys = 0;
            ConnOpen reportLoad = new ConnOpen();
            ConnOpen productLoad = new ConnOpen();
            ConnOpen unitLoad = new ConnOpen();
            ConnOpen userLoad = new ConnOpen();
            ConnOpen userLoad2 = new ConnOpen();
            ConnOpen userLoad3 = new ConnOpen();
            ConnOpen tLoad = new ConnOpen();
            reportLoad.connection.Open();
            productLoad.connection.Open();
            unitLoad.connection.Open();
            userLoad.connection.Open();
            userLoad2.connection.Open();
            userLoad3.connection.Open();
            tLoad.connection.Open();
            //Открыли все коннекты
            SqlCommand commandProduct = new SqlCommand("SELECT * FROM dbo.Product WHERE product_flag = '" + 1 + "'", productLoad.connection);
            SqlDataReader readerProduct = commandProduct.ExecuteReader();
            SqlCommand commandUnit;
            SqlDataReader readerUnit;
            SqlCommand commandT;
            SqlDataReader readerT;
            SqlCommand CQPU = new SqlCommand("SELECT DISTINCT waybill FROM dbo.Responsibility WHERE traffic = '0' AND date > '" + dateTimePicker1.Value.ToShortDateString() + "' AND date < '" + dateTimePicker2.Value.ToShortDateString() + "'", userLoad.connection);
            SqlDataReader RQPU = CQPU.ExecuteReader();
            SqlCommand CQPU2;
            SqlDataReader RQPU2;
            SqlCommand CUP;
            SqlDataReader RUP;
            //Создали команды и датаридеры
            int q = 7, u;
            while (RQPU.Read())
            {
                cell = gh[girdeji];
                workSheet.get_Range(cell).Merge();
                girdeji++;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                u = q + 1;
                ExcelApp.Cells[4, q] = "sany";
                ExcelApp.Cells[4, u] = "jemi bahasy";
                userLoad2.connection.Close();
                userLoad2.connection.Open();
                CQPU2 = new SqlCommand("SELECT * FROM dbo.Responsibility WHERE waybill = '"+RQPU["waybill"].ToString()+"'", userLoad2.connection);
                RQPU2 = CQPU2.ExecuteReader();
                RQPU2.Read();
                resp = RQPU2["responsible"].ToString();
                CUP = new SqlCommand("SELECT * FROM dbo.Users WHERE user_id = '" + resp +"'", userLoad3.connection);
                RUP = CUP.ExecuteReader();
                RUP.Read();
                fi = RUP["fio"].ToString();
                fi = fi.Remove(fi.IndexOf(' ') + 2);
                ExcelApp.Cells[3, q] =RUP["place_of_work"].ToString()+" " + fi + " Nakl № " + RQPU["waybill"].ToString();
                RUP.Close();
                RQPU2.Close();
                userLoad2.connection.Close();
                cell = c[q - 1] + "4:" + c[q - 1] + "26";
                userLoad2.connection.Close();
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                cell = c[q] + "4:" + c[q] + "26";
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                q = q + 2;
            }
            userLoad2.connection.Close();
            userLoad3.connection.Close();
            u = q + 1;
            ExcelApp.Cells[4, q] = "sany";
            ExcelApp.Cells[4, u] = "jemi bahasy";
            cell = c[q - 1] + "2:" + c[q] + "3";
            workSheet.get_Range(cell).Merge();
            ExcelApp.Cells[2, q] = "GIRDEJILERIŇ                                                                JEMI";
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell = c[q - 1] + "4:" + c[q - 1] + "26";
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell = c[q] + "4:" + c[q] + "26";
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            q = q + 2;
            girdeji = girdeji * 2+5;
            userLoad.connection.Close();
            
            userLoad.connection.Open();
            userLoad2.connection.Open();
            userLoad3.connection.Open();
            SqlCommand CQRU = new SqlCommand("SELECT DISTINCT waybill FROM dbo.Responsibility WHERE traffic = '1' AND date > '" + dateTimePicker1.Value.ToShortDateString() + "' AND date < '" + dateTimePicker2.Value.ToShortDateString() + "'", userLoad.connection);
            SqlDataReader RQRU = CQRU.ExecuteReader();
            SqlCommand CQRU2;
            SqlDataReader RQRU2;
            SqlCommand CUR;
            SqlDataReader RUR;
            while (RQRU.Read())
            {
                cell = gh[(girdeji - 5) / 2 + 1 + cykdajys];
                workSheet.get_Range(cell).Merge();
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                cykdajys++;
                u = q + 1;
                ExcelApp.Cells[4, q] = "sany";
                ExcelApp.Cells[4, u] = "jemi bahasy";
                cell = c[q - 1] + "4:" + c[q - 1] + "26";
                userLoad2.connection.Close();
                userLoad2.connection.Open();
                CQRU2 = new SqlCommand("SELECT * FROM dbo.Responsibility WHERE waybill = '" + RQRU["waybill"].ToString() + "'", userLoad2.connection);
                RQRU2 = CQRU2.ExecuteReader();
                RQRU2.Read();
                resp = RQRU2["responsible"].ToString();
                CUR = new SqlCommand("SELECT * FROM dbo.Users WHERE user_id = '" + resp + "'", userLoad3.connection);
                RUR = CUR.ExecuteReader();
                RUR.Read();
                fi = RUR["fio"].ToString();
                fi = fi.Remove(fi.IndexOf(' ') + 2);
                ExcelApp.Cells[3, q] = RUR["place_of_work"].ToString() + " " + fi + " Nakl № " + RQRU["waybill"].ToString();
                RUR.Close();
                RQRU2.Close();
                userLoad2.connection.Close();
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                cell = c[q] + "4:" + c[q] + "26";
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                q = q + 2;
            }
            u = q + 1;
            ExcelApp.Cells[4, q] = "sany";
            ExcelApp.Cells[4, u] = "jemi bahasy";
            cell = c[q - 1] + "2:" + c[q] + "3";
            workSheet.get_Range(cell).Merge();
            ExcelApp.Cells[2, q] = "ÇYKDAJYLARYŇ JEMI";
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell = c[q - 1] + "4:" + c[q - 1] + "26";
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell = c[q] + "4:" + c[q] + "26";
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //-------
            q = q + 2;
            u = q + 1;
            ExcelApp.Cells[4, q] = "sany";
            ExcelApp.Cells[4, u] = "jemi bahasy";
            cell = c[q - 1] + "2:" + c[q] + "3";
            workSheet.get_Range(cell).Merge();
            ExcelApp.Cells[2, q] = dateTimePicker2.Value.ToShortDateString() + " ý              galyndysy"; ;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell = c[q - 1] + "4:" + c[q - 1] + "26";
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell = c[q] + "4:" + c[q] + "26";
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell = "G2:" + c[q] + "4";
            workSheet.get_Range(cell).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range(cell).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            workSheet.get_Range(cell).WrapText = true;
            //-------
            cykdajyb = girdeji+3;
            cykdajys = cykdajyb + cykdajys*2-1;
            userLoad.connection.Close();
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
            cell ="G2:" + c[girdeji] + "2";
            workSheet.get_Range(cell).Merge();
            cell = c[cykdajyb] + "2:" + c[cykdajys] + "2";
            workSheet.get_Range(cell).Merge();
            //-------
            int i = 1, g = 0, n=5;
            while (i<22)
            {
                g = i + 4;
                ExcelApp.Cells[g, 1] = i;
                i++;
            }
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
                commandT = new SqlCommand("SELECT * FROM dbo.Responsibility WHERE product = '" + product + "' AND date < '" + dateTimePicker1.Value.ToShortDateString()+"'", tLoad.connection);
                readerT = commandT.ExecuteReader();
                while (readerT.Read())
                {
                    if (readerT["traffic"].ToString() == "0" && readerT.HasRows == true)
                    {
                        pValue += Convert.ToInt32(readerT["product_quantity"]);
                    }
                    if (readerT["traffic"].ToString() == "1" && readerT.HasRows == true)
                    {
                        rValue += Convert.ToInt32(readerT["product_quantity"]);
                    }
                }
                readerT.Close();
                sum = product_quantity + pValue - rValue;
                sPrice = sum * price;
                ExcelApp.Cells[n, 2] = sProduct;
                ExcelApp.Cells[n, 3] = sUnit;
                ExcelApp.Cells[n, 4] = price;
                ExcelApp.Cells[n, 5] = sum;
                ExcelApp.Cells[n, 6] = sPrice;
                n++;
                pValue = 0;
                rValue = 0;
            }
            reportLoad.connection.Close();
            productLoad.connection.Close();
            unitLoad.connection.Close();
            userLoad.connection.Close();
            tLoad.connection.Close();
            //-------
            n = cykdajyb + 1;
            ExcelApp.Cells[2, 1] = "T №";
            ExcelApp.Cells[2, 2] = "MADDY                                                                                  GYMMATLYKLARYŇ                                                                              ADY";
            ExcelApp.Cells[2, 3] = "Ölçeg birligi";
            ExcelApp.Cells[2, 4] = "Bahasy";
            ExcelApp.Cells[2, 5] = dateTimePicker1.Value.ToShortDateString()+ " ý              galyndysy";
            ExcelApp.Cells[4, 5] = "sany";
            ExcelApp.Cells[4, 6] = "jemi bahasy";
            ExcelApp.Cells[26, 1] = "Jemi";
            ExcelApp.Cells[2, 7] = "G  i  r  d  e  j  i";
            ExcelApp.Cells[2, n] = "Ç  y  k  d  a  j  y";
            //-------
            workSheet.get_Range("A2").Orientation = 90;
            workSheet.get_Range("B2").Font.Bold = true;
            workSheet.get_Range("E2").Font.Bold = true;
            workSheet.get_Range("A26").Font.Bold = true;
            cell = "G2:" + c[girdeji] + "2";
            workSheet.get_Range(cell).Font.Size = 18;
            cell = c[cykdajyb] + "2:" + c[cykdajys] + "2";
            workSheet.get_Range(cell).Font.Size = 18;
            cell = "G2:" + c[girdeji] + "2";
            workSheet.get_Range(cell).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range(cell).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cell = c[cykdajyb] + "2:" + c[cykdajys] + "2";
            workSheet.get_Range(cell).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range(cell).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            workSheet.get_Range("A26").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("A2:F4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.get_Range("A2:F4").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            workSheet.get_Range("A2:F4").WrapText = true;
            workSheet.get_Range("D5:D26").NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            workSheet.get_Range("F5:F26").NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range("A2:F26").Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell = "G2:" + c[girdeji] + "2";
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell = c[cykdajyb] + "2:" + c[cykdajys] + "2";
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.get_Range(cell).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
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
                    if(readerT["traffic"].ToString() == "0" && readerT.HasRows == true)
                    {
                        pValue += Convert.ToInt32(readerT["product_quantity"]);
                    }
                    if (readerT["traffic"].ToString() == "1" && readerT.HasRows == true)
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
