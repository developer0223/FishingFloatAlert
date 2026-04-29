using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using Windows.Devices.Enumeration;

namespace FishingFloatAlertSharp
{
    public partial class MainWindow : Window
    {
        /// <summary>
        /// 선택된 카메라 식별자. <c>-1</c>은 플레이스홀더(「카메라를 선택해주세요」) 또는 미선택.
        /// 0 이상이면 해당 카메라 항목(장치 ID는 선택 항목의 <see cref="CameraComboEntry.DeviceInformationId"/>).
        /// </summary>
        private int selectedCameraId = -1;

        private readonly ObservableCollection<CameraComboEntry> _cameraComboEntries = new();

        public MainWindow()
        {
            InitializeComponent();
            CameraCombo.ItemsSource = _cameraComboEntries;
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            ApplyOuterWindowSizeForClientDips(1600, 900);
        }

        /// <summary>
        /// 제목 표시줄·테두리를 제외한 클라이언트 영역이 (<paramref name="clientWidthDip"/>, <paramref name="clientHeightDip"/>) DIP가 되도록 창 전체 크기를 맞춥니다.
        /// </summary>
        private void ApplyOuterWindowSizeForClientDips(double clientWidthDip, double clientHeightDip)
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd == IntPtr.Zero)
                return;

            var dpi = VisualTreeHelper.GetDpi(this);
            int clientPxW = (int)Math.Round(clientWidthDip * dpi.DpiScaleX);
            int clientPxH = (int)Math.Round(clientHeightDip * dpi.DpiScaleY);

            int style = GetWindowLong(hwnd, GWL_STYLE);
            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);

            var rect = new RECT
            {
                Left = 0,
                Top = 0,
                Right = clientPxW,
                Bottom = clientPxH
            };

            if (!AdjustWindowRectEx(ref rect, style, bMenu: false, exStyle))
                return;

            int outerPxW = rect.Right - rect.Left;
            int outerPxH = rect.Bottom - rect.Top;

            double outerW = outerPxW / dpi.DpiScaleX;
            double outerH = outerPxH / dpi.DpiScaleY;

            Width = outerW;
            Height = outerH;
            MinWidth = MaxWidth = outerW;
            MinHeight = MaxHeight = outerH;
        }

        private async void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            await RefreshCameraListAsync();
        }

        private async void RefreshCamerasButton_OnClick(object sender, RoutedEventArgs e)
        {
            await RefreshCameraListAsync();
        }

        private async Task RefreshCameraListAsync()
        {
            RefreshCamerasButton.IsEnabled = false;
            try
            {
                var selector = DeviceInformation.FindAllAsync(DeviceClass.VideoCapture);
                var devices = await selector.AsTask();

                _cameraComboEntries.Clear();
                _cameraComboEntries.Add(new CameraComboEntry(
                    selectedId: -1,
                    name: "카메라를 선택해주세요",
                    deviceInformationId: null));

                foreach (var d in devices.OrderBy(x => x.Name))
                {
                    var index = _cameraComboEntries.Count - 1;
                    _cameraComboEntries.Add(new CameraComboEntry(
                        selectedId: index,
                        name: d.Name,
                        deviceInformationId: d.Id));
                }

                CameraCombo.SelectedIndex = 0;
                selectedCameraId = -1;
            }
            finally
            {
                RefreshCamerasButton.IsEnabled = true;
            }
        }

        private void CameraCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CameraCombo.SelectedItem is CameraComboEntry entry)
                selectedCameraId = entry.SelectedId;
            else
                selectedCameraId = -1;
        }

        private sealed class CameraComboEntry
        {
            public CameraComboEntry(int selectedId, string name, string? deviceInformationId)
            {
                SelectedId = selectedId;
                Name = name;
                DeviceInformationId = deviceInformationId;
            }

            /// <summary>콤보 식별용. <c>-1</c>은 미선택 플레이스홀더.</summary>
            public int SelectedId { get; }

            public string Name { get; }

            /// <summary>WinRT 캡처용 장치 ID. 플레이스홀더일 때는 <c>null</c>.</summary>
            public string? DeviceInformationId { get; }
        }

        private const int GWL_STYLE = -16;
        private const int GWL_EXSTYLE = -20;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AdjustWindowRectEx(ref RECT lpRect, int dwStyle, bool bMenu, int dwExStyle);

        private static int GetWindowLong(IntPtr hwnd, int nIndex)
        {
            if (IntPtr.Size == 8)
                return unchecked((int)(long)GetWindowLongPtr64(hwnd, nIndex));
            return GetWindowLong32(hwnd, nIndex);
        }
    }
}
