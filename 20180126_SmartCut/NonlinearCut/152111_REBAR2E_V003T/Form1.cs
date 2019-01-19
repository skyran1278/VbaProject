using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace REBAR2ET
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {



            if (_TXTETABS.TextLength == 0 || _COL_REBAR.TextLength == 0)
            {
                MessageBox.Show("請輸入檔名..");
                Application.Exit();
            }
            else
            {
                MessageBox.Show("程式開始...");

                string name1 = _TXTETABS.Text;
                string name2 = _COL_REBAR.Text;
                string name3 = _BEAM_REBAR.Text;


                string hinge_col_i = _COL_hinge_i.Text;
                string hinge_col_j = _COL_hinge_j.Text;
                string hinge_beaml_i = _BEAM_hinge_i.Text;
                string hinge_beaml_j = _BEAM_hinge_j.Text;


                string COL_COVER = _COL_COVER.Text;
                string BEAM_COVRE = _BEAM_COVER.Text;


                DB db = new DB(name1, name2, name3, hinge_col_i, hinge_col_j, hinge_beaml_i, hinge_beaml_j, COL_COVER, BEAM_COVRE);


                READTXTE2K Read_ETABS_TXT = new READTXTE2K(db);
                Read_ETABS_TXT.readinput(name1, db);


                READ_COL Read_COL_TXT = new READ_COL(db);
                Read_COL_TXT.readinput(name2, db);


                READ_BEAM Read_BEAM_TXT = new READ_BEAM(db);
                Read_BEAM_TXT.readinput(name3, db);

                WRITEe2K WriteE2K_TXT = new WRITEe2K();

                WriteE2K_TXT.write_out_e2k(db);


                MessageBox.Show("產生完成...");
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Cursor Files|*.e2k|Cursor Files|*.$2k|All files (*.*)|*.*";
            openFileDialog1.Title = "Select a ETABS File";

            openFileDialog1.ShowDialog();
            this._TXTETABS.Text = openFileDialog1.SafeFileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Cursor Files|*.col|Cursor Files|*.xls|All files (*.*)|*.*";
            openFileDialog2.Title = "Select a TXT File";
   
            openFileDialog2.ShowDialog();
            this._COL_REBAR.Text = openFileDialog2.SafeFileName;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Cursor Files|*.beam|Cursor Files|*.xls|All files (*.*)|*.*";
            openFileDialog2.Title = "Select a TXT File";

            openFileDialog2.ShowDialog();
            this._BEAM_REBAR.Text = openFileDialog2.SafeFileName;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
