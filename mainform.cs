using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;

namespace proeval
{
    public partial class mainform : Form
    {
        public string datapath = System.Environment.CurrentDirectory + "/";
        public mainform()
        {
            InitializeComponent();
            updatecelllist();
        }
        List<string> GetAllFileNames(string path, string pattern = "*.dat")
        {
            List<FileInfo> folder = new DirectoryInfo(path).GetFiles(pattern).ToList();
            return folder.Select(x => x.Name).ToList();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            cellform subform2 = new cellform();
            subform2.Owner = this;
            subform2.ShowDialog();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (!(listBox1.SelectedIndex >= 0))
                    return;
                string cellname = this.listBox1.SelectedItem.ToString() + ".dat";
                DirectoryInfo file = new DirectoryInfo(datapath + cellname);
                DialogResult result = MessageBox.Show("确认删除该单元评价记录？", "警告", MessageBoxButtons.OKCancel);
                if (File.Exists(datapath + cellname) && result == DialogResult.OK)
                {
                    File.Delete(datapath + cellname);
                }

                updatecelllist();
            }
            catch { }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItems.Count == 1)
            {
                cellform subform2 = new cellform();
                subform2.Owner = this;
                string cellname = this.listBox1.SelectedItem.ToString();
                string a = subform2.classcell.cellCEO;
                FileStream fs = new FileStream(datapath + cellname + ".dat", FileMode.Open);
                BinaryFormatter bf = new BinaryFormatter();
                subform2.classcell = bf.Deserialize(fs) as cell;
                fs.Close();
                subform2.form2update();
                subform2.ShowDialog();
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        public void updatecelllist()
        {
            listBox1.Items.Clear();
            foreach (var filename in GetAllFileNames(datapath))
            {
                listBox1.Items.Add(filename.Remove(filename.Length - 4, 4));
            }
        }

        private void mainform_Load(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string S = textBox1.Text;
            this.listBox1.Items.Clear();
            foreach (var filename in GetAllFileNames(datapath))
            {
                if (filename.Contains(S))
                { this.listBox1.Items.Add(filename.Remove(filename.Length - 4, 4)); }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                this.updatecelllist();
            }
        }
        private Size beforeResizeSize = Size.Empty;

        protected override void OnResizeBegin(EventArgs e)
        {
            base.OnResizeBegin(e);
            beforeResizeSize = this.Size;
        }
        protected override void OnResizeEnd(EventArgs e)
        {
            base.OnResizeEnd(e);
            //窗口resize之后的大小
            Size endResizeSize = this.Size;
            //获得变化比例
            float percentWidth = (float)endResizeSize.Width / beforeResizeSize.Width;
            float percentHeight = (float)endResizeSize.Height / beforeResizeSize.Height;
            foreach (Control control in this.Controls)
            {
                if (control is DataGridView) continue;
                //按比例改变控件大小
                control.Width = (int)(control.Width * percentWidth);
                control.Height = (int)(control.Height * percentHeight);
                //为了不使控件之间覆盖 位置也要按比例变化
                control.Left = (int)(control.Left * percentWidth);
                control.Top = (int)(control.Top * percentHeight);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            //Walkthrough m1 = new Walkthrough();
            //m1.Main2();

        }

    }
}
