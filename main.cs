using System;
using System.Windows.Forms;
using W = Microsoft.Office.Interop.Word;

namespace testWord
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Word Documents|*.docx";
                openFileDialog.Title = "Select a Word Document";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    W.Application oWord = null;
                    W.Document oDoc = null;
                    try
                    {
                        string path = openFileDialog.FileName;
                        oWord = new W.Application();
                        oDoc = oWord.Documents.Open(path);
                        foreach (W.Paragraph oPar in oDoc.Paragraphs)
                        {
                            if (!oPar.Range.Text.Equals("\r"))
                            {
                                lblFromWord.Text += oPar.Range.Text;
                            }
                        }
                        oWord.Visible = false;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        if (oDoc != null)
                        {
                            oDoc.Close();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
                        }
                        if (oWord != null)
                        {
                            oWord.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWord);
                        }
                    }
                }
            }
        }
    }
}