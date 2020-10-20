using PortableDeviceApiLib;
using System.Collections.ObjectModel;

namespace PortableDevices
{
    public class PortableDeviceCollection : Collection<PortableDevice>
    {
        private readonly PortableDeviceManager _deviceManager;

        public PortableDeviceCollection()
        {
            _deviceManager = new PortableDeviceManager();
        }

        public void Refresh()
        {
            _deviceManager.RefreshDeviceList();
            // Determine how many WPD devices are connected
            uint count = 1;
            _deviceManager.GetDevices(null, ref count);
            if (count == 0) return;
            // Retrieve the device id for each connected device
            string[] deviceIds = new string[count];
            string temp = "";
            _deviceManager.GetDevices(ref temp, ref count);
            foreach (var deviceId in deviceIds)
            {
                Add(new PortableDevice(deviceId));
            }
        }
    }
}