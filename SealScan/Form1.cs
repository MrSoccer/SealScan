using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.IO;
using System.Diagnostics;



namespace SealScan
{
    public partial class SealScan : Form
    {
        string sourceFile;
        public SealScan()
        {
            InitializeComponent();
            WindowState = FormWindowState.Minimized;
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            WqlEventQuery insertQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_LogicalDisk'");

            ManagementEventWatcher insertWatcher = new ManagementEventWatcher(insertQuery);
            insertWatcher.EventArrived += new EventArrivedEventHandler(DeviceInsertedEvent);
            insertWatcher.Start();

            WqlEventQuery removeQuery = new WqlEventQuery("SELECT * FROM __InstanceDeletionEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_LogicalDisk'");
            ManagementEventWatcher removeWatcher = new ManagementEventWatcher(removeQuery);
            removeWatcher.EventArrived += new EventArrivedEventHandler(DeviceRemovedEvent);
            removeWatcher.Start();

            if (scannerConnected())
            {
                this.Show();
                if (this.WindowState == FormWindowState.Minimized)
                {
                    this.WindowState = FormWindowState.Normal;
                }
            }
            else
            {
                Hide();
                WindowState = FormWindowState.Minimized;
            }
        }

        private bool scannerConnected()
        {
            bool bRC = false;
            var drives = DriveInfo.GetDrives();
            foreach (var drive in drives)
            {
                if (drive.DriveType == DriveType.Removable)
                {
                    Console.WriteLine(drive.Name);
                    this.sourceFile = drive.Name + @"Scanned Barcodes\BARCODES.TXT";

                    if (verifyBarcodeFile(sourceFile))
                    {
                        bRC = true;
                        break;
                    }
                }
            }
            return bRC;
        }

        public new void Activate()
        {

            if (!InvokeRequired)
            {
                base.Activate();
            }
            else
            {
                this.Invoke((MethodInvoker)(() => base.Activate()));
            }
        }
        public new void Show()
        {

            if (!InvokeRequired)
            {
                base.Show();
            }
            else
            {
                this.Invoke((MethodInvoker)(() => base.Show()));
            }
        }

        public new void Hide()
        {
            if (!InvokeRequired)
                base.Hide();
            else
                this.Invoke((MethodInvoker)(() => base.Hide()));
        }

        private void DeviceInsertedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            String tempDeviceID = "";
            String resultDeviceID = "";

            foreach (var property in instance.Properties)
            {
                Console.WriteLine(property.Name + " = " + property.Value);
                if (property.Name.Equals("DeviceID"))
                {
                    tempDeviceID = (string)property.Value;
                    this.sourceFile = tempDeviceID + @"\Scanned Barcodes\BARCODES.TXT";
                    //string sourceFile = tempDeviceID + @"\Scanned Barcodes\BARCODES.TXT";

                    if (verifyBarcodeFile(sourceFile))
                    {
                        try
                        {
                            this.Show();
                            if (this.WindowState == FormWindowState.Minimized)
                            {
                                this.WindowState = FormWindowState.Normal;
                            }
                            //Activate();
                        }
                        catch
                        {

                        }
                    }

                }
                else if (property.Name.Equals("VolumeName") && property.Value.Equals("CS3000") && tempDeviceID != "")
                {
                    resultDeviceID = tempDeviceID;
                }
            }
        }

        private bool verifyBarcodeFile(string sourceFile)
        {
            return File.Exists(sourceFile);
        }

        private void CopyFileToAppDirectory(string sourceFile)
        {
            //string sourceFile = tempDeviceID + @"\Scanned Barcodes\BARCODES.TXT";
            if (!Directory.Exists(@"C:\Temp"))
            {
                Directory.CreateDirectory(@"C:\Temp");
            }
            string destFile = @"c:\Temp\BarCodes.txt";
            File.Copy(sourceFile, destFile, true);
        }

        private void DeviceRemovedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            //String tempDeviceID = "";

            foreach (var property in instance.Properties)
            {
                Console.WriteLine(property.Name + " = " + property.Value);
                if (property.Name.Equals("VolumeName"))
                {
                    if (property.Value.Equals("CS3000"))
                    {
                        Hide();
                        WindowState = FormWindowState.Minimized;
                        break;
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (orderNumberValid() && (vesselNumberValid()))
            {
                CopyFileToAppDirectory(this.sourceFile);
                List<string> barcodeList = Scanner.ScannedBarcodes(this.sourceFile);


                DocClass docClass = new DocClass();

                DateTime dateValue = DateTime.Now;
                string timeStamp = dateValue.ToString("ddMMMyy_HH_mm");

                string templateFile = Properties.Settings.Default.FormTemplate;
                string tempFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\SealReports\SealReport_" + timeStamp + ".docx";
                docClass.CopyMasterFile(templateFile, tempFile);
                docClass.UpdateDateStamp(tempFile);
                docClass.UpdateOrderNumber(tempFile, textBoxOrderNum.Text);
                docClass.UpdateVesselNumber(tempFile, textBoxVesselNum.Text);
                docClass.UpdateWordDocument(tempFile, barcodeList);

#if !DEBUG
                File.Delete(sourceFile);
#endif
                textBoxVesselNum.Text = "";
                textBoxOrderNum.Text = "";
                Hide();
                WindowState = FormWindowState.Minimized;

                System.Diagnostics.Process.Start("winword.exe", tempFile);
            }
        }

        private bool vesselNumberValid()
        {
            bool bRC = false;
            if (textBoxVesselNum.Text != "")
            {
                bRC = true;
            }
            else
            {
                MessageBox.Show("Vessel Number must be entered");
            }
            return bRC;
        }

        private bool orderNumberValid()
        {
            bool bRC = false;
            int result;

            if (textBoxOrderNum.Text == "")
            {
                MessageBox.Show("Order number can't be blank");
            }
            else if (textBoxOrderNum.Text.Length != 7)
            {
                MessageBox.Show("Order number must be 7 digits long");
            }
            else if (!int.TryParse(textBoxOrderNum.Text, out result))
            {
                MessageBox.Show("Order Number field may only contain numbers");
            }
            else
            {
                bRC = true;
            }
            return bRC;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBoxOrderNum_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxVesselNum_TextChanged(object sender, EventArgs e)
        {

        }
    }


    // Called by DriveDetector when removable device in inserted


}
