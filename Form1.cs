using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using HtmlAgilityPack;

namespace ConvertHTML
{
    public partial class Form1 : Form
    {
        string[] fileNames;
        string[] safeFileNames;
        string saveLocation;
        Word.Application wrdApp = new Word.Application();


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Open Word Document";
            ofd.Filter = "Word Document|*.docx";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileNames = ofd.FileNames;
                char[] charactersToTrim = ofd.SafeFileNames.GetValue(0).ToString().ToCharArray();
                saveLocation = fileNames.GetValue(0).ToString().TrimEnd(charactersToTrim);
                string pathString = System.IO.Path.Combine(saveLocation, "HTMLfiles");
                System.IO.Directory.CreateDirectory(pathString);
                string displayNames = "";
                safeFileNames = ofd.SafeFileNames;
                foreach (string entry in ofd.SafeFileNames)
                {
                    displayNames += entry + ", ";
                }

                //Removes the last space and ',' characters
                displayNames = displayNames.Substring(0, displayNames.Length - 2);
                textBox1.ResetText();
                textBox1.AppendText(displayNames);
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (fileNames != null && saveLocation != null)
            {
                //Just a counter to get the safe file name below
                int i = 0;
                foreach (string name in fileNames)
                {
                    try
                    {
                        // removes last 5 characters to get rid of the '.docx' extension
                        string htmlFileName = safeFileNames.GetValue(i).ToString().Substring(0, safeFileNames.GetValue(i).ToString().Length - 5);
                        // 10 corresponds to Filtered HTML document type in word
                        try
                        {
                            this.wrdApp.Documents.Open(name).SaveAs(saveLocation + "\\HTMLFiles\\" + htmlFileName, 10);
                        }
                        catch (Exception e2)
                        {
                            MessageBox.Show("Something went wrong, try again");
                        }
                        finally
                        {
                            this.wrdApp.Documents.Close();
                            scrapeHTMLDocument(saveLocation + "\\HTMLFiles\\" + htmlFileName);
                            i++;
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        MessageBox.Show("File is Open Elsewhere!");
                        //counter still goes up one here in order to keep flow going in the event that one of the files is open
                        i++;
                    }

                }
                MessageBox.Show("All Done!");
            }
            else
            {
                MessageBox.Show("Ensure word files are selected!");
            }
        }
        public void scrapeHTMLDocument(string address)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            //string htmlFileText = System.IO.File.ReadAllText(address + ".htm");
            htmlDoc.Load(address + ".htm");
            var nodeCollection1 = htmlDoc.DocumentNode.SelectNodes("//style");
            var nodeCollection2 = htmlDoc.DocumentNode.SelectNodes("//span");
            if (nodeCollection1 != null )
            {
                foreach (var style in nodeCollection1)
                {
                    style.ParentNode.RemoveChild(style);
                }
               
            }
            if (nodeCollection2 != null)
            {
                foreach (var spanNode in nodeCollection2)
                {
                    if (spanNode.InnerText == null || spanNode.InnerHtml == null)
                        spanNode.ParentNode.RemoveChild(spanNode);
                }
            }
            htmlDoc.Save(address + ".htm");
            //var writeHTML = new System.IO.StreamWriter(address + ".htm",false);
            //writeHTML.Write(htmlFileText);
           // writeHTML.Close();
        }
    }
}
