using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string file = "";
        string find = "";
        string replaceWith = "";
        List<string> repAll = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = "This program will replace text in any .docx file or any format that can be opened by notepad.\r\nTo do all files in a directory, select \"Batch\".";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = this.openFileDialog1.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = this.openFileDialog1.FileName;
            }
        }

        private void button1_Click1(object sender, EventArgs e)
        {
            DialogResult result = this.folderBrowserDialog1.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            file = textBox2.Text;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                label2.Text = "Select Folder";
                label4.Visible = true;
                comboBox1.Visible = true;
                button1.Click += new EventHandler(button1_Click1);
                button1.Click -= new EventHandler(button1_Click);
            }
            else if (checkBox1.Checked != true)
            {
                label2.Text = "Select File";
                label4.Visible = false;
                comboBox1.Visible = false;
                button1.Click += new EventHandler(button1_Click);
                button1.Click -= new EventHandler(button1_Click1);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "")
            {
                if (checkBox1.Checked == true)
                {
                    file = pathTrim(file);
                    if (Directory.Exists(file) == true)
                    {
                        string ext = extTrim(comboBox1.Text);

                        string[] filePaths = Directory.GetFiles(@file, "*" + ext);
                        if (filePaths == null)
                        {
                            MessageBox.Show("There are no files with that extension.\r\n(Format must be \".abc\" with only alphanumeric numbers");
                        }
                        else if (filePaths != null)
                        {
                            foreach (string blip in filePaths)
                            {
                                replace(blip, find, replaceWith);
                            }
                        }


                    }
                    else if (Directory.Exists(file) != true)
                    {
                        MessageBox.Show("This folder does not exist");
                    }


                }
                else if (checkBox1.Checked != true)
                {
                    if (File.Exists(file) == true)
                    {
                        replace(file, find, replaceWith);

                    }
                    else if (File.Exists(file) != true)
                    {
                        MessageBox.Show("This File does not exist");
                    }
                }


            }
            MessageBox.Show("Done.");
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            find = textBox3.Text;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            replaceWith = textBox4.Text;
        }
        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private string pathTrim(string path)
        {
            char[] rem = { '\\', ' ' };
            path = path.TrimEnd(rem);
            path = path + "\\";
            return path;
        }

        private string extTrim(string ext)
        {

            ext = Regex.Replace(ext, @"[\W]", "");
            ext = "." + ext;
            return ext;
        }

        private void replace(string rfile, string rfind, string rrep)
        {
            if (rfile.Contains("~$") != true)
            {
                string ext = Path.GetExtension(rfile);
                if (ext == ".docx")
                {

                    try
                    {
                        string docText = "";
                        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(rfile, true))
                        {
                            using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                            {
                                docText = sr.ReadToEnd();
                            }

                            Regex regexText = new Regex(rfind);
                            docText = regexText.Replace(docText, rrep);


                            using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                            {
                                sw.Write(docText);
                            }
                            foreach (FooterPart footer in wordDoc.MainDocumentPart.FooterParts)
                            {
                                string footerText = null;
                                using (StreamReader sr = new StreamReader(footer.GetStream()))
                                {
                                    footerText = sr.ReadToEnd();
                                }
                                footerText = regexText.Replace(footerText, rrep);
                                using (StreamWriter sw = new StreamWriter(footer.GetStream(FileMode.Create)))
                                {
                                    sw.Write(footerText);
                                }
                            }

                            foreach (HeaderPart header in wordDoc.MainDocumentPart.HeaderParts)
                            {
                                string headerText = null;
                                using (StreamReader sr = new StreamReader(header.GetStream()))
                                {
                                    headerText = sr.ReadToEnd();
                                }


                                headerText = regexText.Replace(headerText, rrep);

                                using (StreamWriter sw = new StreamWriter(header.GetStream(FileMode.Create)))
                                {
                                    sw.Write(headerText);
                                }
                            }


                        }

                    }
                    catch
                    {
                        MessageBox.Show(rfile + "could not be modified");
                    }
                }

                else if (ext != ".docx")
                {
                    try
                    {
                        string txt = File.ReadAllText(rfile);

                        txt.Replace(rfind, rrep);

                        File.WriteAllText(rfind, txt);
                    }
                    catch
                    {
                        MessageBox.Show("The text could not be replaced in \"" + rfile + "\". It may not be an accepted file type.");
                    }
                }



            }
        }
    }
}
