using PortableDeviceApiLib;
using PortableDevices;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace USBLoadApplication
{
    public partial class Form1 : Form
    {
        private PortableDevices.PortableDevice DetectedUSB;

        public Form1()
        {
            InitializeComponent();
            Task.Run(() =>
            {
                bool LastValue = false;
                while (true)
                {
                    bool NewValue = CheckForMTP();
                    if (LastValue != NewValue)
                    {
                        if (NewValue)
                        {// onplugin one shot
                            LoadFiles();
                        }
                    }
                    LastValue = NewValue;
                }
            });
        }

        private void LoadFiles()
        {
            PortableDeviceManager deviceManager = new PortableDeviceManager();
            deviceManager.RefreshDeviceList();
            uint numberOfDevices = 1;
            deviceManager.GetDevices(null, ref numberOfDevices);
            string temp1 = "";
            deviceManager.GetDevices(ref temp1, ref numberOfDevices);
            if (temp1 != "")
            {
                PortableDevices.PortableDevice Device = new PortableDevices.PortableDevice(temp1);
                PortableDeviceFolder root = Device.GetContents();
                PortableDeviceFolder Folder = root.Files.First() as PortableDeviceFolder;
                foreach (string file in Directory.EnumerateFiles("Files"))
                {
                    try
                    {
                        Device.TransferContentToDeviceFromStream(Path.GetFileName(file), new MemoryStream(File.ReadAllBytes(file), false), Folder.Id);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error On upload :" + ex.ToString());
                    }
                }
                MessageBox.Show("Uploaded Files");
            }
        }

        private bool CheckForMTP()
        {
            PortableDeviceManager deviceManager = new PortableDeviceManager();
            deviceManager.RefreshDeviceList();
            uint numberOfDevices = 1;
            deviceManager.GetDevices(null, ref numberOfDevices);
            string temp1 = "";
            deviceManager.GetDevices(ref temp1, ref numberOfDevices);
            if (temp1 != "")
            {
                return true;
            }
            return false;
        }
    }
}