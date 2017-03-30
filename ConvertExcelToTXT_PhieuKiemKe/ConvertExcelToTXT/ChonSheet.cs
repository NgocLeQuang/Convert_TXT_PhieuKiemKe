using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ConvertExcelToTXT
{
    public partial class ChonSheet : Form
    {
        public String strSheet;
        public ChonSheet()
        {
            InitializeComponent();
        }

        public void SetComboBox(List<String> lstr)
        {
            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(lstr.ToArray());
        }

        private void ChonSheet_Load(object sender, EventArgs e)
        {
            if (comboBox1.Items.Count > 0)
                comboBox1.Text = comboBox1.Items[0].ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            strSheet = comboBox1.Text;
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }
    }
}
