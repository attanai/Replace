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
        //Globals
        string file = "";
        string find = "";
        string replaceWith = "";
        
        //List of strings, in case Batch is selected.
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
            //Description
            textBox1.Text = "This program will replace text in any .docx file or any format that can be opened by notepad.\r\nTo do all files in a directory, select \"Batch\".";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Get the File, if batch is not checked.
            DialogResult result = this.openFileDialog1.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = this.openFileDialog1.FileName;
            }
        }

        private void button1_Click1(object sender, EventArgs e)
        {
            //Get the folder is batch is checked
            DialogResult result = this.folderBrowserDialog1.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //Set the global file
            file = textBox2.Text;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //Is Batch checked?
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
            //Is there a file listed?
            if (textBox2.Text != string.Empty;)
            {
                //Batch?
                if (checkBox1.Checked == true)
                {
                    file = pathTrim(file);
                    
                    //Check if the file actually exists.
                    if (Directory.Exists(file) == true)
                    {
                        //Get the extension
                        string ext = extTrim(comboBox1.Text)
                        
                        //List all of the files in the directory
                        string[] filePaths = Directory.GetFiles(@file, "*" + ext);
                        
                        //Make sure that there are actually some.
                        if (filePaths == null)
                        {
                            MessageBox.Show("There are no files with that extension.\r\n(Format must be \".abc\" with only alphanumeric numbers");
                        }
                        else if (filePaths != null)
                        {
                            //Run the replace
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
                //Not a batch
                else if (checkBox1.Checked != true)
                {
                    if (File.Exists(file) == true)
                    {
                        //Run the replace
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
            //Set Find
            find = textBox3.Text;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            //Set ReplaceWith
            replaceWith = textBox4.Text;
        }
        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        
        ///
        ///Trims a folder path to ensure that it uses the correct format
        ///
        private string pathTrim(string path)
        {
            char[] rem = { '\\', ' ' };
            path = path.TrimEnd(rem);
            path = path + "\\";
            return path;
        }

        ///
        ///Get the extension of a file
        ///
        private string extTrim(string ext)
        {
            ext = Regex.Replace(ext, @"[\W]", "");
            ext = "." + ext;
            return ext;
        }

        ///The actual replacement method
        private void replace(string rfile, string rfind, string rrep)
        {
            //Make sure that it's a real file, and not one of those MSWord temp files
            if (rfile.Contains("~$") != true)
            {
                string ext = Path.GetExtension(rfile);
                //If it's a word doc
                if (ext == ".docx")
                {

                    try
                    {
                        //Get the actual text
                        string docText = "";
                        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(rfile, true))
                        {
                            using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                            {
                                docText = sr.ReadToEnd();
                            }

                            Regex regexText = new Regex(rfind);
                            //The actual replacement
                            docText = regexText.Replace(docText, rrep);


                            //Rewrite the file
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

                //If it's not a word doc, do a normal text replace
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
