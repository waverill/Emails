using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using System.Collections.Specialized;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace Emails
{
    public partial class Emailer : Form
    {
        private List<IavaInterface> IAVAs;
        private List<string> IAVA_content;

        public Emailer()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text=@"C:\Users\waverill\IAVAs";
            textBox2.Text = @"C:\Users\waverill\IAVAs_Final";
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int size = textBox1.Text.Length + 1;
            this.IAVAs = new List<IavaInterface>();
            this.IAVA_content = new List<string>();
            try
            {
                var can_files = Directory.GetFiles(textBox1.Text);
                List<string> files = new List<string>();
                int count = 0;
                foreach (string f in can_files)
                {
                    if (f.Substring(f.Length - 5) == ".docx" && f.Substring(size, 8) == "Precoord")
                    {
                        files.Add(f);
                        count++;
                    }
                }
                string msg = "Found " + count + " IAVA files.";
                MessageBox.Show(msg);
                int c = 0;
                string msg2 = "";
                Regex iava = new Regex(@"(2012-[A-B]-[0-9]{4})");
                Regex title = new Regex(@"^Precoord 2012-[A-B]-[0-9]{4} (.*)$");
                if (count == 0)
                {
                    msg2 = "Could not find any Precoord (.docx) files in selected directory.  Please try again.";
                }
                foreach (string f in files)
                {
                    msg += "\n\n" + c + ") " + f.Substring(size);
                    c++;
                    try
                    {
                      /*  Microsoft.Office.Interop.Word.Application wordObject = new Microsoft.Office.Interop.Word.Application();
                        Document doc = wordObject.Documents.Open(f);
                        string t = doc.Name;
                        doc.ActiveWindow.Selection.WholeStory();
                        doc.ActiveWindow.Selection.Copy();
                        IDataObject data = System.Windows.Forms.Clipboard.GetDataObject();
                        String temp = data.GetData(System.Windows.Forms.DataFormats.StringFormat).ToString();
                        System.Windows.Forms.Clipboard.SetDataObject(string.Empty);*/
                        DocxTextReader docxReader = new DocxTextReader(f);
                        string temp = docxReader.getDocumentText();
                        Match m = iava.Match(f.Substring(size));
                        Match m2 = title.Match(f.Substring(size));
                        string iav_title = m2.Groups[1].Captures[0].ToString();
                        iav_title = iav_title.Substring(0, iav_title.Length - 5);
                        IAVA_content.Add(temp);
                        // msg2 +="\n" + doc.Name;
                        IavaInterface II = new IavaInterface(m.Value, iav_title);
                        IAVAs.Add(II);
                    //    msg2+=c+ ") " + m.Value + " - " + iav_title + "\r\n";
                        msg2 += II.iava_ack_template();
                        msg2 += temp;
                      
                        string temp2 = II.iava_ack_template();
                        temp2 += temp;
                        II.Make_IAVA_Pre_ACK(temp2);
                        //wordObject.Quit();
                    }
                    catch (Exception j)
                    {
                        MessageBox.Show(j.Message);
                    }
                }
               
                textBox3.Text = msg2;

               // System.Windows.Forms.MessageBox.Show(msg);
                          
            }
            catch (Exception l)
            {
                MessageBox.Show(l.Message);
            }

           

        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\Users\waverill\IAVAs_Final";
            openFileDialog1.Filter = "text files (*.txt)|*.txt";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog1.FileName;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // this.IAVAs[0].DoFinalAcks2(textBox2.Text); 
            int size = textBox1.Text.Length + 1;
            try
            {
                var can_files = Directory.GetFiles(textBox1.Text);
                List<string> files = new List<string>();
                int count = 0;
                foreach (string f in can_files)
                {
                    if (f.Substring(f.Length - 5) == ".docx" && f.Substring(size, 5) == "2012-")
                    {
                        files.Add(f);
                        count++;
                    }
                }
                string msg = "Found " + count + " IAVA files.";
                MessageBox.Show(msg);
                Regex iava = new Regex(@"(2012-[A-B]-[0-9]{4})");
                Regex title = new Regex(@"^2012-[A-B]-[0-9]{4} (.*)$");
                foreach (string f in files)
                {
                   try
                    {
                        DocxTextReader docxReader = new DocxTextReader(f);
                        string temp = docxReader.getDocumentText();
                        Match m = iava.Match(f.Substring(size));
                        Match m2 = title.Match(f.Substring(size));
                        string iav_title = m2.Groups[1].Captures[0].ToString();
                        iav_title = iav_title.Substring(0, iav_title.Length - 5);
                        IavaInterface II = new IavaInterface(m.Value, iav_title);
                        II.DoFinalAcks3(temp, textBox2.Text);                       
                    }
                    catch (Exception j)
                    {
                        MessageBox.Show(j.Message);
                    }
                }
            }
            catch (Exception l)
            {
                MessageBox.Show(l.Message);
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
