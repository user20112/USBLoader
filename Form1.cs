using PortableDeviceApiLib;
using PortableDevices;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Automation;
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
                        else
                        {//unplug one shot
                            CloseModalWindows();
                        }
                    }
                    LastValue = NewValue;
                }
            });
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

        private void CloseModalWindows()
        {
            // get the main window
            AutomationElement root = AutomationElement.FromHandle(Process.GetCurrentProcess().MainWindowHandle);
            if (root == null)
                return;
            // it should implement the Window pattern
            if (!root.TryGetCurrentPattern(WindowPattern.Pattern, out object pattern))
                return;
            WindowPattern window = (WindowPattern)pattern;
            if (window.Current.WindowInteractionState != WindowInteractionState.ReadyForUserInteraction)
            {
                // get sub windows
                foreach (AutomationElement element in root.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)))
                {
                    // hmmm... is it really a window?
                    if (element.TryGetCurrentPattern(WindowPattern.Pattern, out pattern))
                    {
                        // if it's ready, try to close it
                        WindowPattern childWindow = (WindowPattern)pattern;
                        if (childWindow.Current.WindowInteractionState == WindowInteractionState.ReadyForUserInteraction)
                        {
                            childWindow.Close();
                        }
                    }
                }
            }
        }

        private void LoadFiles()
        {
            PortableDeviceManager deviceManager = new PortableDeviceManager();
            deviceManager.RefreshDeviceList();
            uint numberOfDevices = 1;
            deviceManager.GetDevices(null, ref numberOfDevices);
            string[] deviceIds;
            string temp1 = "";
            deviceIds = new string[numberOfDevices];
            deviceManager.GetDevices(ref temp1, ref numberOfDevices);
            PortableDevices.PortableDevice Drive = null;
            try
            {
                if (temp1 != "")
                {
                    PortableDevices.PortableDevice Device = new PortableDevices.PortableDevice(temp1);
                    try
                    {
                        Device.Connect();
                        string temp = Device.FriendlyName;
                        PortableDeviceFolder root = Device.GetContents();
                        PortableDeviceFolder folder = root.Files.First() as PortableDeviceFolder;
                        foreach (string file in Directory.EnumerateFiles("Files"))
                        {
                            Device.TransferContentToDeviceFromStream(Path.GetFileName(file), new MemoryStream(File.ReadAllBytes(file)), folder.Id);
                        }
                    }
                    catch (Exception ex)
                    {
                        Device.Disconnect();
                        MessageBox.Show("Unable to load files Has this one already been loaded?");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hit Exception:" + ex.ToString());
            }
            MessageBox.Show("Offload Complete! please unplug drive");
        }
    }
}