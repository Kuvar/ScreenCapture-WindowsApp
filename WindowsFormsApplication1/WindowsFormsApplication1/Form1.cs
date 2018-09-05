using System;
using System.Drawing.Imaging;
using System.Windows.Forms;
using ScreenShotDemo;
using Xceed.Words.NET;
using Image = Xceed.Words.NET.Image;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        int counter = 0;
        public Form1()
        {
            InitializeComponent();
            RegisterHotKeyClass registerHotKey=new RegisterHotKeyClass();
            registerHotKey.Keys = Keys.PrintScreen;
            registerHotKey.ModKey = 0;
            registerHotKey.WindowHandle = this.Handle;
            registerHotKey.HotKey += new RegisterHotKeyClass.HotKeyPass(registerHotKeyHandler);
            registerHotKey.StarHotKey();
        }

        void registerHotKeyHandler()
        {
            if (radioButton1.Checked)
            {
                SaveToImage();
            }
            else if (radioButton2.Checked)
            {
                SaveToWordFile();
            }
            else
            {
                MessageBox.Show("Please Select Source Type","InforMation",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
        }

        public void SaveToImage()
        {
            try
            {
                string fileName = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\" + "Capture Screen" + "\\" + "Image" + "\\" + DateTime.Now.ToString("dd-MM-yyyy") + "\\";
                if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\" + "Capture Screen" + "\\" + "Image" + "\\" + DateTime.Now.ToString("dd-MM-yyyy") + "\\"))
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\" + "Capture Screen" + "\\" + "Image" + "\\" + DateTime.Now.ToString("dd-MM-yyyy") + "\\");
                }
                ScreenCapture sc = new ScreenCapture();
                // capture entire screen, and save it to a file
                System.Drawing.Image img = sc.CaptureScreen();
                // display image in a Picture control named imageDisplay
                pictureBox1.Image = img;
                // capture this window, and save it;
                //sc.CaptureWindowToFile(this.Handle, Guid.NewGuid()+".gif", ImageFormat.Gif);
                img.Save(fileName + "\\" + "Screenshot" + Guid.NewGuid() + ".jpg", ImageFormat.Gif);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SaveToWordFile()
        {
            ScreenCapture sc = new ScreenCapture();
            // capture entire screen, and save it to a file
            System.Drawing.Image img1 = sc.CaptureScreen();
            // display image in a Picture control named imageDisplay
            pictureBox1.Image = img1;
            // capture this window, and save it;
            //sc.CaptureWindowToFile(this.Handle, Guid.NewGuid()+".gif", ImageFormat.Gif);
            img1.Save("Screenshot" + ".jpg", ImageFormat.Gif);
            counter++;
            string fileName = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\" + "Capture Screen" + "\\";
            if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\" + "Capture Screen" + "\\"))
            {
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\" + "Capture Screen" + "\\");
            }
            if (!File.Exists(fileName + DateTime.Now.ToString("dd-MM-yyyy") + ".docx"))
            {
                var docCreate = DocX.Create(fileName + DateTime.Now.ToString("dd-MM-yyyy") + ".docx");
                docCreate.Save();
            }
            var docload = DocX.Load(fileName + DateTime.Now.ToString("dd-MM-yyyy") + ".docx");
            Image img = docload.AddImage("Screenshot.jpg");
            Picture p = img.CreatePicture();
            p.Width = 690;
            p.Height = 365;
            Xceed.Words.NET.Paragraph par = docload.InsertParagraph(Convert.ToString(counter));
            par.AppendPicture(p);
            docload.Save();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
        }
    }
}
