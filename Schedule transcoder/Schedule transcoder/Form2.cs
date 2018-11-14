using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Schedule_transcoder
{
    public partial class Form2 : Form
    {
        StreamWriter myWriter = null;
        Form1 myForm = new Form1();

        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {

                myWriter = new StreamWriter("Update Key.txt");

                myWriter.WriteLine(txtKey.Text.Substring(0, txtKey.Text.Length - 2));

                myWriter.Close();

                Application.Restart();

            }
            catch
            {
                MessageBox.Show("Update Key invalid\nContact  Software Engineer.", "Error");
            }
        }
        static string HexStringToString(string hexString)
        {
            var sb = new StringBuilder();

            try
            {
                if (hexString == null || (hexString.Length & 1) == 1)
                {

                }

                for (int i = 0; i < hexString.Length; i += 2)
                {
                    string hexChar = hexString.Substring(i, 2);
                    sb.Append((char)Convert.ToByte(hexChar, 16));
                }
            }
            catch
            {
                MessageBox.Show("Invalid Update Key.","Error");
                Application.ExitThread();
            }
            
            return sb.ToString();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.ExitThread();
        }
    }
}
