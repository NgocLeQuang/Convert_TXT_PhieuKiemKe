using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ConvertExcelToTXT
{

    public partial class Convert : Form
    {
        private DataTable _kq = new DataTable();
        public string Filename = "";

        public Convert()
        {
            InitializeComponent();
        }

        public bool Check_Version()
        {
            string version =(from w in Global.db.tbl_Versions where w.IDProject == Global.IDProject select w.IDVersion).FirstOrDefault();
            if (lb_Version.Text == version)
            {
                return true;
            }
            try
            {
                MessageBox.Show("Bạn đang sử dụng phiên bản cũ, Vui lòng cập nhật phiên bản mới từ server!","Thông báo!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Process.Start(Global.UrlUpdateVersion);
                Application.Exit();
            }
            catch
            {
                MessageBox.Show("Error updated, not get new Version!");
            }
            return false;
        }
       
        string getcharacter(int n, string str)
        {
            string kq = "";
            for (int i = 1; i <= n; i++)
            {
                kq = kq.Insert(kq.Length, str);
            }

            return kq;
        }

        private string ThemKyTubatKyStrPhiaSau(string input, int iByte, string str)
        {
            if (input.Length > iByte)
                return input.Substring(0, iByte);
            if (input.Length == iByte)
                return input;

            return input.Insert(input.Length, getcharacter(iByte - input.Length, str));
        }

        private string ThemKyTuPhiaTruocVaBoKyTuPhiaSau(string input, int ibyte, string str)
        {
            if (input.Length >= ibyte)
                return input.Substring(0, ibyte);
            return input.Insert(0, getcharacter(ibyte - input.Length, str));
        }

        //private string ThemKyTubatKyPhiatruoc(string input, int iByte, string str)
        //{
        //    if (input.Length >= iByte)
        //        return input.Substring(input.Length - iByte, iByte);

        //    return input.Insert(0, getcharacter(iByte - input.Length, str));
        //}

        //private string ThemKyTu0PhiaTruocVaBoKyTuPhiaSau(string input, int ibyte, string str)
        //{
        //    if (input.Length >= ibyte)
        //        return input.Substring(0, ibyte);
        //    if(string.IsNullOrEmpty(input))
        //        return input.Insert(0, getcharacter(ibyte, " "));
        //    return input.Insert(0, getcharacter(ibyte - input.Length, str));
        //}

        ////Nếu số 0 ở trước thì đổi lại thành space. 2 số 0 thành 2 space. số 0 phía sau và 1 kí tự khác 0 phía trước thì giữ nguyên số 0. Ví dụ: '01' thành ' 1', '00' thành '  ', '10' thành '10'
        //private string EditString(string strInput)
        //{
        //    var kq = strInput;
        //    if (kq.Length != 2)
        //        return kq;
        //    if (kq[0].ToString() != "0")
        //        return kq;
        //    return kq.Replace('0', ' ');
        //}

        //private string ThemKyTuBatKyPhiaTruocVaSpaceKhiKhongCoGi(string strInput, int iByte, string str)
        //{
        //    if (string.IsNullOrEmpty(strInput))
        //    {
        //        return ThemKyTubatKyPhiatruoc(strInput, iByte, " ");
        //    }
        //    return ThemKyTubatKyPhiatruoc(strInput, iByte, str);
        //}


        //private string ThemKyTuTruongTien(string input)
        //{
        //    if (input.IndexOf("?") >= 0)
        //        return "???????????";
        //    if (string.IsNullOrEmpty(input))
        //        return ThemKyTubatKyPhiatruoc(input, 11, " ");
        //    return ThemKyTubatKyPhiatruoc(input, 11, "0");
        //}

        private void Convert_txt()
        {
            try
            {
                File.Delete(@".\" + Filename + ".txt");
                File.Delete(@".\" + Filename + "_WithPart" + ".txt");
                int dem = 1;
                for (int i = 0; i < _kq.Rows.Count; i += dem)
                {
                    string logd1 = "",logd2="";
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][1].ToString(), 2, " "); //Trường cố định
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][2].ToString(), 6, " ");//trường No
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][3].ToString(), 4, " ");//truong code
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][4].ToString(), 10, " ");//truong 04
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][5].ToString(), 15, " ");//truong 05
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][6].ToString(), 2, " ");//truong stt
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][7].ToString(), 30, " ");//truong 07
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][8].ToString(), 25, " ");//truong 08
                    logd1 += ThemKyTuPhiaTruocVaBoKyTuPhiaSau(_kq.Rows[i][9]+".00", 10, "0");//truong 9 - số lượng
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][10].ToString(), 2, " ");//truong 10
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][11].ToString(), 2, " ");//truong 11-mặc định
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][12].ToString(), 3, " ");//truong 12- phân loại
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][13].ToString(), 20, " ");
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][14].ToString(), 30, " ");
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][15].ToString(), 3, " ");
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][16].ToString(), 8, " ");
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][17].ToString(), 20, " ");
                    logd1 += ThemKyTubatKyStrPhiaSau(_kq.Rows[i][18].ToString(), 58, " ");

                    LogFile.WriteLog(Filename + ".txt", logd1);

                    logd2 = logd1 + ThemKyTubatKyStrPhiaSau(_kq.Rows[i][19].ToString(), 35, " ");
                    LogFile.WriteLog(Filename+"_WithPart" + ".txt", logd2);

                }
                Process.Start(Filename + ".txt");
                Process.Start(Filename + "_WithPart" + ".txt");
            }
            catch (Exception)
            {
                MessageBox.Show("Có lỗi xẩy ra. Có thể file excel không đúng chuẩn file của dự án Phiếu kiểm kê.", "Lỗi",MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btn_browser_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = @"Excel files |*.xls;*.xlsx";
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
                    if (cs.ShowDialog() == DialogResult.OK)
                    {
                        dataGridView1.DataSource =
                            _kq = DatatableFroExcelFile.exceldata(openFileDialog1.FileName, cs.strSheet);
                        Filename = openFileDialog1.SafeFileName;
                        btn_convert.Enabled = true;

                        txt_batch.Text = "";
                        this.Text = Global.NameProgram + "       " +openFileDialog1.FileName;
                        lb_partfile.Text = cs.strSheet;
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(@"Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }
        private void btn_convert_Click(object sender, EventArgs e)
        {
           
            if (Check_Version())
            {
                //if (string.IsNullOrEmpty(txt_batch.Text))
                //{
                //    if (MessageBox.Show("Bạn chưa nhập đường dẫn batch, Bạn vẫn muỗn convert!", "Cảnh bảo",MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                //        Convert_txt();
                //}
                //else
                //{
                    Convert_txt();
                //}
            }
        }

        private void excitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btn_browser_Click(null, null);
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            lb_Version.Text = Global.Version;
            Text = Global.NameProgram;
        }
    }
}
