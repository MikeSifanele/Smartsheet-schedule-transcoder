using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Schedule_transcoder
{
    public class Initial_Components
    {
        StreamReader myReader = null;
        Form2 myForm = new Form2();
        Form1 myForm1 = new Form1();

        public void Load()
        {
            try
            {
                myReader = new StreamReader("Update Key.txt");

                if (Convert.ToDateTime(DateTime.Now.ToShortDateString()) >= Convert.ToDateTime(_ToString(myReader.ReadLine())))
                {
                    myReader.Close();

                    DialogResult dialogResult = MessageBox.Show($"Application requires update.\nDo you wish to update?", "Warning", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        myForm.Show();

                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        Application.ExitThread();
                    }
                }
            }
            catch
            {
                MessageBox.Show("Invalid Update Key.", "Error");
                Application.ExitThread();
            }
            
        }
        static string _ToString(string _String)
        {
            if (_String == null || (_String.Length & 1) == 1)
            {

            }
            var sb = new StringBuilder();
            for (int i = 0; i < _String.Length; i += 2)
            {
                string _Char = _String.Substring(i, 2);
                sb.Append((char)Convert.ToByte(_Char, 16));
            }
            return sb.ToString();
        }
    }
}
