using System.Collections.Generic;
using DirectShowLib;

namespace FishingFloatAlertSharp
{
    internal static class DirectShowVideoDevices
    {
        /// <summary>DirectShow Video Input Device. <paramref name="Index"/>는 OpenCV DSHOW 인덱스와 동일한 전역 순서입니다.</summary>
        public static IEnumerable<(int Index, string Name, string DevicePath)> EnumerateVideoInputDevices()
        {
            var i = 0;
            foreach (DsDevice d in DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice))
            {
                yield return (i, (d.Name ?? "").Trim(), d.DevicePath ?? "");
                i++;
            }
        }

        public static int FindIndexByDevicePath(string devicePath)
        {
            if (string.IsNullOrEmpty(devicePath))
                return -1;
            foreach (var (index, _, path) in EnumerateVideoInputDevices())
            {
                if (string.Equals(path, devicePath, StringComparison.OrdinalIgnoreCase))
                    return index;
            }

            return -1;
        }
    }
}
