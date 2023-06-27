using System;
using System.Windows.Forms;

namespace proeval
{
    public partial class comment : Form
    {
        public string text;
        public comment()
        {
            InitializeComponent();
        }

        private void comment_Load(object sender, EventArgs e)
        {
            textBox1.Text = text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            cellform father;
            father = (cellform)(this.Owner);
            father.classcell.expertcomment = textBox1.Text;
            this.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
