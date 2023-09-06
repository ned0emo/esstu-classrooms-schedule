using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rasp
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
          
        }
        public bool flagger=false;
        private void button1_Click(object sender, EventArgs e)
        {
            flagger = true;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //comboBox1.SelectedIndex = -1;
            this.Close();
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
              //  comboBox1.SelectedIndex = -1;
        }
    }
}
