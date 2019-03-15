using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kursovayaa
{
    public partial class Result : Form
    {

        public Result()
        {
            InitializeComponent();
        }

        public string Ploshad
        {
            get { return textBoxS.Text; }
            set { textBoxS.Text = value; }
        }

        private void textBoxData_TextChanged(object sender, EventArgs e)
        {

        }

        private void Result_Load(object sender, EventArgs e)
        {
            textBoxData.Text = DateTime.Now.ToString("dd.MM.yyyy");
            
        }

        private void textBoxS_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
