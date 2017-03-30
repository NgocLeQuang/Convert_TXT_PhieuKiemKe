using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ConvertExcelToTXT
{
    public partial class Form1 : Form
    {
        DataTable kq = new DataTable();
        String filename = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel files |*.xls;*.xlsx";
            //openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    

                    //chọn sheet
                    ChonSheet cs = new ChonSheet();
                    cs.SetComboBox(DatatableFroExcelFile.GetListSheetNameFromFileExcel(openFileDialog1.FileName));
                    if (cs.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        dataGridView1.DataSource = kq = DatatableFroExcelFile.exceldata(openFileDialog1.FileName, cs.strSheet);
                        filename = openFileDialog1.SafeFileName;
                        button2.Enabled = true;
                    }


                        
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        String getcharacter(int n, String str)
        {
            String kq = "";
            for (int i = 1; i <= n; i++)
            {
                kq = kq.Insert(kq.Length, str);
            }
            
            return kq;
        }

        private String ThemKyTubatKySTRPhiaSau(String input, int iByte, string str)
        {

            if (input.Length > iByte)
                return input.Remove(iByte);
            if (input.Length == iByte)
                return input;

            return input.Insert(input.Length, getcharacter(iByte - input.Length, str));
        }

        private String ThemKyTubatKyPhiatruoc(String input, int iByte, string str)
        {
            if (input.Length >= iByte)
                return input.Substring(input.Length - iByte, iByte );

            return input.Insert(0, getcharacter(iByte - input.Length, str));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                File.Delete(@".\" + filename + ".txt");

                for (int i = 0; i < kq.Rows.Count; i++)
                {
                    String log = "";

                    log += ThemKyTubatKySTRPhiaSau(textBox1.Text + kq.Rows[i][0].ToString(), 80, " "); //Đường dẫn và thư mục file ảnh, " " 80bte
                    log += ",";
                    log += ThemKyTubatKyPhiatruoc(kq.Rows[i][1].ToString(), 2, "0");//Mã Tỉnh. Phải đủ 2 byte
                    log += ",";
                    log += ThemKyTubatKySTRPhiaSau(kq.Rows[i][2].ToString(), 4, " ");//Mã văn phòng. 4 byte Nếu không đủ byte, xuất space phía sau
                    log += ",";
                    log += ThemKyTubatKyPhiatruoc(kq.Rows[i][3].ToString(), 6, " ");//Mã số bảo hiểm. 6byte Nếu không đủ byte, xuất space phía trước
                    log += ",";
                    log += kq.Rows[i][4].ToString();//Năm nhật 1 byte, Ngoài số này thì là lỗi -- chưa có phương thức để kiểm tra lỗi
                    log += ",";
                    log += kq.Rows[i][5].ToString();//Ngày tháng năm, 6 byte, check logic? -- Không hiểu check logic, có lấp đầy 6 byte không?
                    log += ",";
                    log += ThemKyTubatKyPhiatruoc(kq.Rows[i][6].ToString().Length > 4 ? "    " : kq.Rows[i][6].ToString(), 4, " ");//SỨC KHỎE. 4 byte, Nếu không đủ byte, xuất space phía trước
                    log += ",";
                    log += ThemKyTubatKyPhiatruoc(kq.Rows[i][7].ToString().Length > 4 ? "    " : kq.Rows[i][7].ToString(), 4, " ");//LƯƠNG HƯU, 4 byte, Nếu không đủ byte, xuất space phía trước
                    log += ",";
                    log += (kq.Rows[i][8].ToString().Length > 7 ? "9999999" : kq.Rows[i][8].ToString());//TiỀN LƯƠNG, Nếu quá 7 kí tự xuất 9999999, Không có dữ liệu thì để trống
                    log += ",";
                    log += (kq.Rows[i][9].ToString().Length > 7 ? "9999999" : kq.Rows[i][9].ToString());//9. ĐiỀU CHỈNH, Nếu quá 7 kí tự xuất 9999999, Không có dữ liệu thì để trống
                    log += ",";

                    LogFile.WriteLog(filename + ".txt", log);
                }
                Process.Start(filename + ".txt");
            }
            catch (Exception)
            {
                MessageBox.Show("Có lỗi xẩy ra. Có thể file excel không đúng chuẩn file của dự án Takahashi. Vui lòng liên hệ NAMDT. Thanks!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            


            
        }

        private void excitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button1_Click(null, null);
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Chương trình Convert excel to txt\r\n@Copyright: Đặng Thanh Nam - namdt@mptelecom.com.vn", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button2.Enabled = false;
        }

        private void label3_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("mailto:namdt@mptelecom.com.vn");
            }
            catch (Exception)
            {
                
                
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show(ThemKyTubatKyPhiatruoc("abcd", 4, "*"));
        }
    }
}
