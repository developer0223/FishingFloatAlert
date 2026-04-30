using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Input;
using System.Windows.Threading;
using Windows.Devices.Enumeration;
using Windows.Graphics.Imaging;
using Windows.Media.Capture;
using Windows.Media.Capture.Frames;
using Windows.Media.Devices;
using Windows.Media.MediaProperties;
using Windows.Security.Cryptography;

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

        private MediaCapture? _mediaCapture;
        private MediaFrameReader? _frameReader;
        private WriteableBitmap? _previewBitmap;

        /// <summary>타겟 색 ± 허용 범위에 맞는 픽셀만 흰색으로 표시한 미리보기용 비트맵.</summary>
        private WriteableBitmap? _maskPreviewBitmap;

        private int _previewWidth;
        private int _previewHeight;
        private int _previewBlitScheduled;
        private long _lastNullSoftwareBitmapLogTick;
        private long _lastFrameProcessErrorLogTick;
        private bool _suppressCameraSelectionChanged;

        /// <summary>DirectShow + OpenCV 미리보기(OBS Virtual Camera 등 WinRT 미노출 장치).</summary>
        private OpenCvSharp.VideoCapture? _openCvCapture;

        private DispatcherTimer? _openCvPreviewTimer;

        private DispatcherTimer? _clockTimer;

        /// <summary>미리보기에 겹친 ROI(오버레이 좌표, DIP). 실제 프레임은 <see cref="TryGetPreviewContentRect"/> 안에 맞춰 둡니다.</summary>
        private Rect _previewRoiRectDips;

        private bool _previewRoiHasPlacement;
        private bool _previewRoiDragging;
        private PreviewRoiHitKind _previewRoiActiveHit;
        private Point _previewRoiDragMouseStart;
        private Rect _previewRoiDragRectStart;
        private Rect _previewRoiDragContentBounds;

        /// <summary>원본 프레임(<see cref="_previewWidth"/>×<see cref="_previewHeight"/>)에서 ROI의 왼쪽·위·너비·높이 비율(0~1).</summary>
        private double _previewRoiSourceNormX;

        private double _previewRoiSourceNormY;
        private double _previewRoiSourceNormWidth;
        private double _previewRoiSourceNormHeight;

        /// <summary><see cref="_previewRoiSourceNormX"/> 등을 현재 원본 해상도로 환산한 정수 사각형(픽셀).</summary>
        private Int32Rect _previewRoiSourcePixels;

        private const double PreviewRoiMinSizeDip = 32;
        private const double PreviewRoiCornerHitDip = 14;
        private const double PreviewRoiEdgeHitDip = 10;

        private Color _targetColor = Color.FromRgb(0, 255, 0);

        private bool _eyedropperActive;

        private bool _suppressTargetColorPickerEvents;

        /// <summary>타겟 RGB 각 채널에서 허용할 최대 차이(0=일치만).</summary>
        private int _colorChannelTolerance = 12;

        /// <summary>ROI 유사색이 이 개수 이하일 때만 지속 비프합니다.</summary>
        private const int FloatAlarmSustainMatchMax = 5;

        private volatile bool _floatAlarmSustainRequested;

        private volatile bool _floatAlarmBeepWorkerShutdown;

        private Thread? _floatAlarmBeepThread;

        private readonly object _floatAlarmBeepThreadLock = new();

        private bool _tabMaskPreviewHold;

        private bool _suppressPersistUserSettings;

        private enum PreviewRoiHitKind
        {
            None,
            Move,
            ResizeNorth,
            ResizeNorthEast,
            ResizeEast,
            ResizeSouthEast,
            ResizeSouth,
            ResizeSouthWest,
            ResizeWest,
            ResizeNorthWest,
        }

        public MainWindow()
        {
            InitializeComponent();
            CameraDeviceComboBox.ItemsSource = _cameraComboEntries;
            LoadUserSettingsFromFile();
            ApplyTargetColorToPicker();
            SyncColorToleranceFromSlider();
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
            StartClockTimer();
            EnsureFloatAlarmSustainBeepWorker();
            await RefreshCameraListAsync();
        }

        private void MainWindow_OnClosing(object? sender, CancelEventArgs e)
        {
            StopClockTimer();
            StopFloatAlarmSustainBeepWorker();
            PersistUserSettingsToFile();
            StopPreviewSync();
        }

        private void MainWindow_OnDeactivated(object? sender, EventArgs e)
        {
            if (!_tabMaskPreviewHold)
                return;
            _tabMaskPreviewHold = false;
            UpdatePreviewImageSourceForCurrentMode();
        }

        private void MainWindow_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Tab || e.IsRepeat)
                return;
            _tabMaskPreviewHold = true;
            UpdatePreviewImageSourceForCurrentMode();
            e.Handled = true;
        }

        private void MainWindow_OnPreviewKeyUp(object sender, KeyEventArgs e)
        {
            var isTab = e.Key == Key.Tab || (e.Key == Key.System && e.SystemKey == Key.Tab);
            if (!isTab)
                return;
            if (!_tabMaskPreviewHold)
                return;
            _tabMaskPreviewHold = false;
            UpdatePreviewImageSourceForCurrentMode();
            e.Handled = true;
        }

        private static string GetUserSettingsConfigPath()
            => Path.Combine(AppContext.BaseDirectory, "settings", "config.txt");

        private void LoadUserSettingsFromFile()
        {
            _suppressPersistUserSettings = true;
            try
            {
                var path = GetUserSettingsConfigPath();
                if (!File.Exists(path))
                    return;

                foreach (var rawLine in File.ReadAllLines(path))
                {
                    var line = rawLine.Trim();
                    if (line.Length == 0 || line.StartsWith('#'))
                        continue;
                    var eq = line.IndexOf('=');
                    if (eq <= 0)
                        continue;
                    var key = line[..eq].Trim();
                    var val = line[(eq + 1)..].Trim();
                    if (key.Equals("TARGET_COLOR", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            var conv = System.Windows.Media.ColorConverter.ConvertFromString(val);
                            if (conv is Color c)
                                _targetColor = c;
                        }
                        catch
                        {
                            // ignore bad color
                        }
                    }
                    else if (key.Equals("COLOR_TOLERANCE", StringComparison.OrdinalIgnoreCase))
                    {
                        if (int.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out var tol))
                        {
                            tol = Math.Clamp(tol, 0, 50);
                            _colorChannelTolerance = tol;
                            if (ColorToleranceSlider is not null)
                                ColorToleranceSlider.Value = tol;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogWarning("설정 파일 읽기 실패(settings/config.txt)", ex);
            }
            finally
            {
                _suppressPersistUserSettings = false;
            }
        }

        private void PersistUserSettingsToFile()
        {
            if (_suppressPersistUserSettings)
                return;
            try
            {
                var dir = Path.Combine(AppContext.BaseDirectory, "settings");
                Directory.CreateDirectory(dir);
                var path = GetUserSettingsConfigPath();
                var c = _targetColor;
                var hex = $"#{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}";
                var text =
                    $"TARGET_COLOR={hex}{Environment.NewLine}COLOR_TOLERANCE={_colorChannelTolerance.ToString(CultureInfo.InvariantCulture)}{Environment.NewLine}";
                File.WriteAllText(path, text);
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogWarning("설정 파일 저장 실패(settings/config.txt)", ex);
            }
        }

        private void StartClockTimer()
        {
            StopClockTimer();
            _clockTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _clockTimer.Tick += ClockTimer_OnTick;
            _clockTimer.Start();
            UpdateClockText();
        }

        private void StopClockTimer()
        {
            if (_clockTimer is null)
                return;
            _clockTimer.Stop();
            _clockTimer.Tick -= ClockTimer_OnTick;
            _clockTimer = null;
        }

        private void ClockTimer_OnTick(object? sender, EventArgs e) => UpdateClockText();

        private void UpdateClockText()
        {
            ClockTextBlock.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private async void RefreshCamerasButton_OnClick(object sender, RoutedEventArgs e)
        {
            await RefreshCameraListAsync();
        }

        private async Task RefreshCameraListAsync()
        {
            await StopPreviewAsync();

            string? keepDeviceId = null;
            string? keepName = null;
            if (CameraDeviceComboBox.SelectedItem is CameraComboEntry cur && cur.SelectedId >= 0)
            {
                keepDeviceId = cur.DeviceInformationId;
                keepName = cur.Name;
            }

            RefreshCamerasButton.IsEnabled = false;
            _suppressCameraSelectionChanged = true;
            try
            {
                var rows = await EnumerateCameraRowsForComboAsync();

                _cameraComboEntries.Clear();
                _cameraComboEntries.Add(new CameraComboEntry(
                    selectedId: -1,
                    name: "카메라를 선택해주세요",
                    deviceInformationId: null));

                foreach (var (name, id) in rows)
                {
                    var index = _cameraComboEntries.Count - 1;
                    _cameraComboEntries.Add(new CameraComboEntry(
                        selectedId: index,
                        name: name,
                        deviceInformationId: id));
                }

                CameraComboEntry? restored = null;
                if (!string.IsNullOrEmpty(keepDeviceId))
                {
                    restored = _cameraComboEntries.FirstOrDefault(e =>
                        e.DeviceInformationId != null
                        && string.Equals(e.DeviceInformationId, keepDeviceId, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(e.Name, keepName ?? "", StringComparison.OrdinalIgnoreCase))
                        ?? _cameraComboEntries.FirstOrDefault(e =>
                            e.DeviceInformationId != null
                            && string.Equals(e.DeviceInformationId, keepDeviceId, StringComparison.OrdinalIgnoreCase));
                }

                if (restored != null)
                    CameraDeviceComboBox.SelectedItem = restored;
                else
                    CameraDeviceComboBox.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogError("카메라 목록 새로고침", ex);
            }
            finally
            {
                _suppressCameraSelectionChanged = false;
                RefreshCamerasButton.IsEnabled = true;
            }

            if (CameraDeviceComboBox.SelectedItem is CameraComboEntry entry)
                selectedCameraId = entry.SelectedId;
            else
                selectedCameraId = -1;

            await RestartPreviewForSelectionAsync();
        }

        /// <summary>
        /// WinRT에 노출되는 비디오 캡처 장치를 여러 쿼리로 합쳐 엽니다. (가상 카메라가 한쪽 쿼리에만 나오는 경우 대비)
        /// </summary>
        private static async Task<List<DeviceInformation>> EnumerateVideoCaptureDevicesMergedAsync()
        {
            var mediaSelector = MediaDevice.GetVideoCaptureSelector();
            var list0 = await DeviceInformation.FindAllAsync(mediaSelector).AsTask();
            var list1 = await DeviceInformation.FindAllAsync(DeviceClass.VideoCapture).AsTask();

            IReadOnlyList<DeviceInformation> list2 = Array.Empty<DeviceInformation>();
            try
            {
                // KSCATEGORY_VIDEO_CAMERA (USB/가상 카메라 일부가 여기에만 잡힘).
                // 주의: {65E8773D-...} 는 KSCATEGORY_CAPTURE(오디오)라 마이크·라인이 섞이면 안 됨.
                const string kscVideoCameraAqs =
                    "System.Devices.InterfaceClassGuid:=\"{E5323777-E976-4F5B-9B55-B94699C46E44}\" AND System.Devices.InterfaceEnabled:=System.StructuredQueryType.Boolean#True";
                list2 = await DeviceInformation.FindAllAsync(kscVideoCameraAqs).AsTask();
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogWarning("KSCATEGORY_VIDEO_CAMERA 추가 열거 실패(무시)", ex);
            }

            var byId = list0
                .Concat(list1)
                .Concat(list2)
                .GroupBy(d => d.Id, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .Where(d => !IsLikelyAudioCaptureDeviceName(d.Name))
                .ToList();

            // 동일 표시 이름·다른 인터페이스 Id 중복(예: DESKTOP-… 여러 줄) 정리
            return byId
                .GroupBy(d => (d.Name ?? "").Trim(), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderBy(x => x.Id, StringComparer.OrdinalIgnoreCase).First())
                .OrderBy(d => d.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>잘못된 열거가 섞였을 때를 대비한 보험(오디오 입력 표시명).</summary>
        private static bool IsLikelyAudioCaptureDeviceName(string? name)
        {
            if (string.IsNullOrEmpty(name))
                return false;
            var n = name.Trim();
            if (n.Contains("Realtek HD Audio", StringComparison.OrdinalIgnoreCase))
                return true;
            if (n.Contains("Mic input", StringComparison.OrdinalIgnoreCase))
                return true;
            if (n.Contains("Line input", StringComparison.OrdinalIgnoreCase))
                return true;
            if (n.Contains("Stereo input", StringComparison.OrdinalIgnoreCase))
                return true;
            if (string.Equals(n, "USB Audio", StringComparison.OrdinalIgnoreCase))
                return true;
            if (n.Contains("USBC Headset", StringComparison.OrdinalIgnoreCase))
                return true;
            return n.Contains("Headset", StringComparison.OrdinalIgnoreCase)
                   && !n.Contains("Camera", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>WinRT 목록 + DirectShow 전용(이름이 WinRT와 겹치지 않는) 장치.</summary>
        private static async Task<List<(string Name, string Id)>> EnumerateCameraRowsForComboAsync()
        {
            var winRt = await EnumerateVideoCaptureDevicesMergedAsync();
            var winNameSet = new HashSet<string>(
                winRt.Select(d => (d.Name ?? "").Trim()),
                StringComparer.OrdinalIgnoreCase);

            var rows = new List<(string Name, string Id)>();
            foreach (var d in winRt.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
                rows.Add((d.Name, d.Id));

            try
            {
                // OpenCV DSHOW는 장치 경로 문자열이 아니라 열거 순서의 정수 인덱스로 여는 것이 안정적입니다.
                foreach (var (index, name, path) in DirectShowVideoDevices.EnumerateVideoInputDevices())
                {
                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(path))
                        continue;
                    if (winNameSet.Contains(name.Trim()))
                        continue;
                    rows.Add((name.Trim(), "dshowidx:" + index.ToString(CultureInfo.InvariantCulture)));
                }
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogWarning("DirectShow 비디오 입력 열거 실패", ex);
            }

            return rows;
        }

        private async void CameraDeviceComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressCameraSelectionChanged)
                return;

            if (CameraDeviceComboBox.SelectedItem is CameraComboEntry entry)
                selectedCameraId = entry.SelectedId;
            else
                selectedCameraId = -1;

            await RestartPreviewForSelectionAsync();
        }

        private async Task RestartPreviewForSelectionAsync()
        {
            await StopPreviewAsync();
            if (selectedCameraId < 0)
            {
                await ClearPreviewSurfaceAsync();
                return;
            }

            if (CameraDeviceComboBox.SelectedItem is not CameraComboEntry { DeviceInformationId: { } deviceId })
            {
                await ClearPreviewSurfaceAsync();
                return;
            }

            try
            {
                await StartPreviewAsync(deviceId);
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogError("카메라 미리보기 시작 실패", ex);
                await StopPreviewAsync();
                PreviewImage.Source = null;
            }
        }

        private async Task StopPreviewAsync()
        {
            await StopOpenCvPreviewAsync();

            if (_frameReader != null)
            {
                _frameReader.FrameArrived -= OnPreviewFrameArrived;
                try
                {
                    await _frameReader.StopAsync();
                }
                catch (Exception ex)
                {
                    AppDiagnostics.LogWarning("MediaFrameReader.StopAsync 중 예외(무시)", ex);
                }

                _frameReader.Dispose();
                _frameReader = null;
                // 일부 드라이버는 Stop 직후 즉시 Dispose하면 다음 오픈이 실패할 수 있어 짧게 양보합니다.
                await Task.Delay(120);
            }

            if (_mediaCapture != null)
            {
                try
                {
                    _mediaCapture.Dispose();
                }
                catch (Exception ex)
                {
                    AppDiagnostics.LogWarning("MediaCapture.Dispose 중 예외(무시)", ex);
                }

                _mediaCapture = null;
                await Task.Delay(80);
            }

            _previewBitmap = null;
            _maskPreviewBitmap = null;
            _previewWidth = 0;
            _previewHeight = 0;
            await ClearPreviewSurfaceAsync();
        }

        private Task ClearPreviewSurfaceAsync()
        {
            return Dispatcher.InvokeAsync(() =>
            {
                PreviewImage.Source = null;
                ResetPreviewRoiOverlay();
            }).Task;
        }

        private void StopPreviewSync()
        {
            StopOpenCvPreviewSync();

            try
            {
                if (_frameReader != null)
                {
                    _frameReader.FrameArrived -= OnPreviewFrameArrived;
                    _frameReader.StopAsync().AsTask().ConfigureAwait(false).GetAwaiter().GetResult();
                    _frameReader.Dispose();
                    _frameReader = null;
                    Thread.Sleep(120);
                }
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogWarning("미리보기 동기 중지(MediaFrameReader) 예외(무시)", ex);
            }

            try
            {
                _mediaCapture?.Dispose();
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogWarning("미리보기 동기 중지(MediaCapture) 예외(무시)", ex);
            }

            _mediaCapture = null;
            Thread.Sleep(80);
            _previewBitmap = null;
            _maskPreviewBitmap = null;
            _previewWidth = 0;
            _previewHeight = 0;
            try
            {
                Dispatcher.Invoke(() =>
                {
                    PreviewImage.Source = null;
                    ResetPreviewRoiOverlay();
                }, DispatcherPriority.Send);
            }
            catch
            {
                // 창 종료 중 Dispatcher 무효일 수 있음
            }
        }

        private Task StopOpenCvPreviewAsync()
        {
            return Dispatcher.InvokeAsync(() =>
            {
                if (_openCvPreviewTimer != null)
                {
                    _openCvPreviewTimer.Stop();
                    _openCvPreviewTimer.Tick -= OnOpenCvPreviewTick;
                    _openCvPreviewTimer = null;
                }

                _openCvCapture?.Dispose();
                _openCvCapture = null;
            }).Task;
        }

        private void StopOpenCvPreviewSync()
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    if (_openCvPreviewTimer != null)
                    {
                        _openCvPreviewTimer.Stop();
                        _openCvPreviewTimer.Tick -= OnOpenCvPreviewTick;
                        _openCvPreviewTimer = null;
                    }

                    _openCvCapture?.Dispose();
                    _openCvCapture = null;
                }, DispatcherPriority.Send);
            }
            catch
            {
                // 창 종료 중
            }
        }

        private async Task StartOpenCvDirectShowPreviewAsync(string deviceIdMarker)
        {
            await StopOpenCvPreviewAsync();

            int camIndex;
            if (deviceIdMarker.StartsWith("dshowidx:", StringComparison.OrdinalIgnoreCase))
            {
                var s = deviceIdMarker["dshowidx:".Length..];
                if (!int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out camIndex) || camIndex < 0)
                    throw new InvalidOperationException("DirectShow 장치 인덱스가 올바르지 않습니다: " + deviceIdMarker);
            }
            else if (deviceIdMarker.StartsWith("dshowpath:", StringComparison.OrdinalIgnoreCase))
            {
                var path = Uri.UnescapeDataString(deviceIdMarker["dshowpath:".Length..]);
                camIndex = DirectShowVideoDevices.FindIndexByDevicePath(path);
                if (camIndex < 0)
                    throw new InvalidOperationException("DirectShow 경로에 해당하는 인덱스를 찾을 수 없습니다: " + path);
                AppDiagnostics.LogWarning("dshowpath: 형식은 dshowidx: 로 대체되었습니다. 목록 새로고침을 권장합니다.");
            }
            else
            {
                throw new InvalidOperationException("지원하지 않는 DirectShow 장치 Id: " + deviceIdMarker);
            }

            await Dispatcher.InvokeAsync(() =>
            {
                PreviewImage.Source = null;
                _previewBitmap = null;
                _maskPreviewBitmap = null;
                _previewWidth = 0;
                _previewHeight = 0;

                var cap = new OpenCvSharp.VideoCapture(camIndex, OpenCvSharp.VideoCaptureAPIs.DSHOW);
                if (!cap.IsOpened())
                {
                    cap.Dispose();
                    cap = new OpenCvSharp.VideoCapture(camIndex, OpenCvSharp.VideoCaptureAPIs.ANY);
                }

                if (!cap.IsOpened())
                    throw new InvalidOperationException("OpenCV로 비디오 입력을 열 수 없습니다. DirectShow index=" + camIndex);

                TryRequestOpenCvCaptureResolution(cap);

                _openCvCapture = cap;
                _openCvPreviewTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(33) };
                _openCvPreviewTimer.Tick += OnOpenCvPreviewTick;
                _openCvPreviewTimer.Start();
            });
        }

        private void OnOpenCvPreviewTick(object? sender, EventArgs e)
        {
            if (_openCvCapture is null || !_openCvCapture.IsOpened())
                return;

            var cap = _openCvCapture;

            using var frame = new OpenCvSharp.Mat();
            if (!cap.Read(frame) || frame.Empty())
                return;

            if (frame.Channels() == 1)
                OpenCvSharp.Cv2.CvtColor(frame, frame, OpenCvSharp.ColorConversionCodes.GRAY2BGR);

            using var bgra = new OpenCvSharp.Mat();
            OpenCvSharp.Cv2.CvtColor(frame, bgra, OpenCvSharp.ColorConversionCodes.BGR2BGRA);

            var w = bgra.Cols;
            var h = bgra.Rows;
            var stride = w * 4;
            var len = stride * h;
            var pixels = new byte[len];
            if (bgra.IsContinuous())
                Marshal.Copy(bgra.Data, pixels, 0, len);
            else
            {
                for (var r = 0; r < h; r++)
                    Marshal.Copy(bgra.Ptr(r), pixels, r * stride, stride);
            }

            if (_previewBitmap == null || w != _previewWidth || h != _previewHeight)
            {
                _previewBitmap = new WriteableBitmap(w, h, 96, 96, PixelFormats.Pbgra32, null);
                _maskPreviewBitmap = new WriteableBitmap(w, h, 96, 96, PixelFormats.Pbgra32, null);
                _previewWidth = w;
                _previewHeight = h;
                RefreshPreviewRoiLayout();
            }

            _previewBitmap.Lock();
            try
            {
                var dstStride = _previewBitmap.BackBufferStride;
                if (dstStride == stride)
                    Marshal.Copy(pixels, 0, _previewBitmap.BackBuffer, len);
                else
                {
                    for (var row = 0; row < h; row++)
                        Marshal.Copy(pixels, row * stride, _previewBitmap.BackBuffer + row * dstStride, Math.Min(stride, dstStride));
                }

                _previewBitmap.AddDirtyRect(new Int32Rect(0, 0, w, h));
            }
            finally
            {
                _previewBitmap.Unlock();
            }

            ProcessFloatDetection(pixels, w, h);
            ApplyPreviewDisplayModeAfterFrame(pixels, w, h);
        }

        private async Task StartPreviewAsync(string deviceId)
        {
            if (deviceId.StartsWith("dshowidx:", StringComparison.OrdinalIgnoreCase)
                || deviceId.StartsWith("dshowpath:", StringComparison.OrdinalIgnoreCase))
            {
                await StartOpenCvDirectShowPreviewAsync(deviceId);
                return;
            }

            var capture = await InitializeMediaCaptureAsync(deviceId);
            _mediaCapture = capture;

            var frameSource = await PrepareFrameSourceWithBestFormatAsync(capture);
            if (frameSource == null)
            {
                AppDiagnostics.LogError($"카메라 프레임 소스 없음(FrameSources 비어 있음). deviceId={deviceId}");
                await StopPreviewAsync();
                return;
            }

            await Dispatcher.InvokeAsync(() =>
            {
                PreviewImage.Source = null;
                _previewBitmap = null;
                _maskPreviewBitmap = null;
                _previewWidth = 0;
                _previewHeight = 0;
            });

            // MJPEG/YUY2 등에서 SoftwareBitmap이 비는 경우가 많아, 파이프라인이 BGRA8로 디코드한 프레임을 요청합니다.
            MediaFrameReader reader;
            try
            {
                reader = await capture.CreateFrameReaderAsync(frameSource, MediaEncodingSubtypes.Bgra8);
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogWarning("BGRA8 FrameReader 생성 실패, 기본 소스로 재시도", ex);
                reader = await capture.CreateFrameReaderAsync(frameSource);
            }

            reader.AcquisitionMode = MediaFrameReaderAcquisitionMode.Realtime;
            reader.FrameArrived += OnPreviewFrameArrived;
            try
            {
                await reader.StartAsync();
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogError("MediaFrameReader.StartAsync 실패", ex);
                reader.Dispose();
                throw;
            }

            _frameReader = reader;
        }

        /// <summary>가능하면 공유 읽기 전용으로 열어 다른 앱과 동시 사용 여지를 둡니다. 미지원/실패 시 전용 모드로 재시도합니다.</summary>
        private static async Task<MediaCapture> InitializeMediaCaptureAsync(string deviceId)
        {
            var sharedSettings = new MediaCaptureInitializationSettings
            {
                VideoDeviceId = deviceId,
                StreamingCaptureMode = StreamingCaptureMode.Video,
                SharingMode = MediaCaptureSharingMode.SharedReadOnly,
                // GPU 전용 버퍼만 주는 드라이버에서 SoftwareBitmap 이 비는 경우를 줄입니다.
                MemoryPreference = MediaCaptureMemoryPreference.Cpu
            };

            var capShared = new MediaCapture();
            try
            {
                await capShared.InitializeAsync(sharedSettings);
                return capShared;
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogWarning("MediaCapture SharedReadOnly 초기화 실패 → ExclusiveControl 시도", ex);
                capShared.Dispose();
            }

            var exclusiveSettings = new MediaCaptureInitializationSettings
            {
                VideoDeviceId = deviceId,
                StreamingCaptureMode = StreamingCaptureMode.Video,
                SharingMode = MediaCaptureSharingMode.ExclusiveControl,
                MemoryPreference = MediaCaptureMemoryPreference.Cpu
            };

            var capExclusive = new MediaCapture();
            try
            {
                await capExclusive.InitializeAsync(exclusiveSettings);
                return capExclusive;
            }
            catch (Exception ex)
            {
                AppDiagnostics.LogError("MediaCapture ExclusiveControl 초기화 실패", ex);
                capExclusive.Dispose();
                throw;
            }
        }

        /// <summary>
        /// 모든 프레임 소스·지원 포맷 중 픽셀 수가 가장 큰 조합을 고르고 <see cref="MediaFrameSource.SetFormatAsync"/>로 적용합니다.
        /// (기존: 첫 VideoPreview 핀만 선택 + 포맷 미설정 → 드라이버 기본 640×480 등으로 표시되는 경우가 많음)
        /// </summary>
        private static async Task<MediaFrameSource?> PrepareFrameSourceWithBestFormatAsync(MediaCapture capture)
        {
            if (capture.FrameSources.Count == 0)
                return null;

            MediaFrameSource? bestSource = null;
            MediaFrameFormat? bestFormat = null;
            long bestArea = -1;
            var bestTypeRank = int.MinValue;
            var bestFps = -1.0;

            foreach (var source in capture.FrameSources.Values)
            {
                var typeRank = MediaStreamTypeRank(source.Info.MediaStreamType);
                foreach (var format in source.SupportedFormats)
                {
                    var vf = format.VideoFormat;
                    if (vf is null)
                        continue;

                    long area = (long)vf.Width * vf.Height;
                    var fps = FrameRateToFps(format.FrameRate);

                    var better = area > bestArea
                        || (area == bestArea && typeRank > bestTypeRank)
                        || (area == bestArea && typeRank == bestTypeRank && fps > bestFps);

                    if (!better)
                        continue;

                    bestArea = area;
                    bestTypeRank = typeRank;
                    bestFps = fps;
                    bestSource = source;
                    bestFormat = format;
                }
            }

            if (bestSource is null)
                return capture.FrameSources.Values.FirstOrDefault();

            if (bestFormat is not null)
            {
                try
                {
                    await bestSource.SetFormatAsync(bestFormat).AsTask();
                }
                catch (Exception ex)
                {
                    AppDiagnostics.LogWarning("카메라 고해상도 포맷 적용 실패(드라이버 기본 해상도 유지)", ex);
                }
            }

            return bestSource;
        }

        private static int MediaStreamTypeRank(MediaStreamType t) => t switch
        {
            MediaStreamType.VideoRecord => 10,
            MediaStreamType.VideoPreview => 5,
            _ => 0,
        };

        private static double FrameRateToFps(MediaRatio? r)
        {
            if (r is null || r.Denominator == 0)
                return 0;
            return r.Numerator / (double)r.Denominator;
        }

        /// <summary>DirectShow 기본값(종종 320×240~640×480)보다 높은 해상도를 순차 요청합니다.</summary>
        private static void TryRequestOpenCvCaptureResolution(OpenCvSharp.VideoCapture cap)
        {
            (int w, int h)[] targets = [(1920, 1080), (1280, 720), (960, 540), (800, 600)];
            using var probe = new OpenCvSharp.Mat();
            foreach (var (tw, th) in targets)
            {
                cap.Set(OpenCvSharp.VideoCaptureProperties.FrameWidth, tw);
                cap.Set(OpenCvSharp.VideoCaptureProperties.FrameHeight, th);
                if (!cap.Read(probe) || probe.Empty())
                    continue;

                if (probe.Width * probe.Height >= 1280 * 720 * 0.9)
                    return;
            }
        }

        private void OnPreviewFrameArrived(MediaFrameReader sender, object args)
        {
            if (Interlocked.CompareExchange(ref _previewBlitScheduled, 1, 0) != 0)
                return;

            using var frameRef = sender.TryAcquireLatestFrame();
            if (frameRef?.VideoMediaFrame is not { } vmf)
            {
                Interlocked.Exchange(ref _previewBlitScheduled, 0);
                return;
            }

            SoftwareBitmap? surfaceCopy = null;
            try
            {
                var sbInput = vmf.SoftwareBitmap;
                if (sbInput == null && vmf.Direct3DSurface != null)
                {
                    try
                    {
                        surfaceCopy = SoftwareBitmap.CreateCopyFromSurfaceAsync(vmf.Direct3DSurface)
                            .AsTask()
                            .ConfigureAwait(false)
                            .GetAwaiter()
                            .GetResult();
                        sbInput = surfaceCopy;
                    }
                    catch (Exception ex)
                    {
                        AppDiagnostics.LogWarning("Direct3DSurface → SoftwareBitmap 복사 실패", ex);
                    }
                }

                if (sbInput is null)
                {
                    var now = Environment.TickCount64;
                    if (now - _lastNullSoftwareBitmapLogTick > 2000)
                    {
                        _lastNullSoftwareBitmapLogTick = now;
                        var sub = frameRef.Format?.Subtype ?? "(unknown)";
                        var hasD3d = vmf.Direct3DSurface != null;
                        AppDiagnostics.LogWarning(
                            $"프레임에 SoftwareBitmap 없음 (subtype={sub}, Direct3DSurface={(hasD3d ? "있음" : "없음")}). 드라이버/포맷 확인.");
                    }

                    Interlocked.Exchange(ref _previewBlitScheduled, 0);
                    return;
                }

                byte[]? pixels;
                int w;
                int h;
                try
                {
                    using var converted = SoftwareBitmap.Convert(sbInput, BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
                    w = converted.PixelWidth;
                    h = converted.PixelHeight;
                    var capacity = (uint)(w * h * 4);
                    var buffer = new Windows.Storage.Streams.Buffer(capacity) { Length = capacity };
                    converted.CopyToBuffer(buffer);
                    CryptographicBuffer.CopyToByteArray(buffer, out pixels);
                    if (pixels == null || pixels.Length == 0)
                    {
                        Interlocked.Exchange(ref _previewBlitScheduled, 0);
                        return;
                    }
                }
                catch (Exception ex)
                {
                    var t = Environment.TickCount64;
                    if (t - _lastFrameProcessErrorLogTick > 2000)
                    {
                        _lastFrameProcessErrorLogTick = t;
                        AppDiagnostics.LogError("카메라 프레임 변환/복사", ex);
                    }

                    Interlocked.Exchange(ref _previewBlitScheduled, 0);
                    return;
                }

                var frameBytes = pixels!;
                DispatchPreviewFrame(frameBytes, w, h);
            }
            finally
            {
                surfaceCopy?.Dispose();
            }
        }

        private void DispatchPreviewFrame(byte[] frameBytes, int w, int h)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Render, () =>
            {
                try
                {
                    if (_previewBitmap == null || w != _previewWidth || h != _previewHeight)
                    {
                        _previewBitmap = new WriteableBitmap(w, h, 96, 96, PixelFormats.Pbgra32, null);
                        _maskPreviewBitmap = new WriteableBitmap(w, h, 96, 96, PixelFormats.Pbgra32, null);
                        _previewWidth = w;
                        _previewHeight = h;
                        RefreshPreviewRoiLayout();
                    }

                    _previewBitmap.Lock();
                    try
                    {
                        int dstStride = _previewBitmap.BackBufferStride;
                        int srcStride = w * 4;
                        if (dstStride == srcStride)
                        {
                            Marshal.Copy(frameBytes, 0, _previewBitmap.BackBuffer, frameBytes.Length);
                        }
                        else
                        {
                            for (int row = 0; row < h; row++)
                            {
                                Marshal.Copy(
                                    frameBytes,
                                    row * srcStride,
                                    _previewBitmap.BackBuffer + row * dstStride,
                                    Math.Min(srcStride, dstStride));
                            }
                        }

                        _previewBitmap.AddDirtyRect(new Int32Rect(0, 0, w, h));
                    }
                    finally
                    {
                        _previewBitmap.Unlock();
                    }

                    ProcessFloatDetection(frameBytes, w, h);
                    ApplyPreviewDisplayModeAfterFrame(frameBytes, w, h);
                }
                catch (Exception ex)
                {
                    AppDiagnostics.LogError("미리보기 UI 갱신", ex);
                }
                finally
                {
                    Interlocked.Exchange(ref _previewBlitScheduled, 0);
                }
            });
        }

        private void RoiOverlayCanvas_OnSizeChanged(object sender, SizeChangedEventArgs e) => RefreshPreviewRoiLayout();

        private void RoiOverlayCanvas_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (RoiOverlayCanvas is null)
                return;

            if (_eyedropperActive)
            {
                var eyedropPoint = e.GetPosition(RoiOverlayCanvas);
                if (TrySamplePreviewPixelColor(eyedropPoint, out var picked))
                {
                    var c = picked;
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, () =>
                    {
                        _suppressTargetColorPickerEvents = true;
                        try
                        {
                            _targetColor = c;
                            if (TargetColorPickerControl is not null)
                                TargetColorPickerControl.SelectedColor = c;
                        }
                        finally
                        {
                            _suppressTargetColorPickerEvents = false;
                        }

                        PersistUserSettingsToFile();
                        EyedropperToggleButton.IsChecked = false;
                    });
                }

                e.Handled = true;
                return;
            }

            if (!TryGetPreviewContentRect(out var content))
                return;

            if (!_previewRoiHasPlacement)
            {
                _previewRoiRectDips = CreateDefaultPreviewRoi(content);
                _previewRoiHasPlacement = true;
                UpdatePreviewRoiVisual();
            }

            var p = e.GetPosition(RoiOverlayCanvas);
            var hit = HitTestPreviewRoi(p, _previewRoiRectDips);
            if (hit == PreviewRoiHitKind.None)
                return;

            _previewRoiDragging = true;
            _previewRoiActiveHit = hit;
            _previewRoiDragMouseStart = p;
            _previewRoiDragRectStart = _previewRoiRectDips;
            _previewRoiDragContentBounds = content;
            RoiOverlayCanvas.CaptureMouse();
            e.Handled = true;
        }

        private void RoiOverlayCanvas_OnMouseMove(object sender, MouseEventArgs e)
        {
            if (RoiOverlayCanvas is null)
                return;

            if (_eyedropperActive)
            {
                RoiOverlayCanvas.Cursor = Cursors.Cross;
                return;
            }

            if (!_previewRoiDragging)
            {
                if (_previewRoiHasPlacement && TryGetPreviewContentRect(out _))
                {
                    var p = e.GetPosition(RoiOverlayCanvas);
                    RoiOverlayCanvas.Cursor = CursorForPreviewRoiHit(HitTestPreviewRoi(p, _previewRoiRectDips));
                }
                else
                {
                    RoiOverlayCanvas.Cursor = Cursors.Arrow;
                }

                return;
            }

            if (e.LeftButton != MouseButtonState.Pressed)
            {
                FinishPreviewRoiMouseDrag();
                return;
            }

            var cur = e.GetPosition(RoiOverlayCanvas);
            _previewRoiRectDips = ComputePreviewRoiDuringDrag(cur);
            UpdatePreviewRoiVisual();
            e.Handled = true;
        }

        private void RoiOverlayCanvas_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_previewRoiDragging)
                return;
            FinishPreviewRoiMouseDrag();
            e.Handled = true;
        }

        private void RoiOverlayCanvas_OnLostMouseCapture(object sender, MouseEventArgs e) => FinishPreviewRoiMouseDrag();

        private void FinishPreviewRoiMouseDrag()
        {
            if (!_previewRoiDragging)
                return;
            _previewRoiDragging = false;
            _previewRoiActiveHit = PreviewRoiHitKind.None;
            RoiOverlayCanvas?.ReleaseMouseCapture();
            if (TryGetPreviewContentRect(out var content))
                _previewRoiRectDips = ClampPreviewRoiToContent(_previewRoiRectDips, content);
            UpdatePreviewRoiVisual();
        }

        private void ResetPreviewRoiOverlay()
        {
            _previewRoiHasPlacement = false;
            _previewRoiDragging = false;
            _previewRoiActiveHit = PreviewRoiHitKind.None;
            _previewRoiRectDips = Rect.Empty;
            RoiOverlayCanvas?.ReleaseMouseCapture();
            ResetFloatAlarmTracking();
            UpdatePreviewRoiVisual();
        }

        private void RefreshPreviewRoiLayout()
        {
            if (RoiOverlayCanvas is null || !TryGetPreviewContentRect(out var content))
                return;

            if (!_previewRoiHasPlacement)
            {
                if (_previewWidth > 0 && _previewHeight > 0)
                {
                    _previewRoiRectDips = CreateDefaultPreviewRoi(content);
                    _previewRoiHasPlacement = true;
                }
            }
            else
            {
                _previewRoiRectDips = ClampPreviewRoiToContent(_previewRoiRectDips, content);
            }

            UpdatePreviewRoiVisual();
        }

        private bool TryGetPreviewContentRect(out Rect contentRect)
        {
            contentRect = default;
            if (RoiOverlayCanvas is null)
                return false;

            var aw = RoiOverlayCanvas.ActualWidth;
            var ah = RoiOverlayCanvas.ActualHeight;
            if (aw < 1 || ah < 1)
                return false;

            if (_previewWidth <= 0 || _previewHeight <= 0)
            {
                contentRect = new Rect(0, 0, aw, ah);
                return true;
            }

            var s = Math.Min(aw / _previewWidth, ah / _previewHeight);
            var dw = _previewWidth * s;
            var dh = _previewHeight * s;
            var ox = (aw - dw) * 0.5;
            var oy = (ah - dh) * 0.5;
            contentRect = new Rect(ox, oy, dw, dh);
            return true;
        }

        private static Rect CreateDefaultPreviewRoi(Rect content)
        {
            var cw = Math.Max(PreviewRoiMinSizeDip, content.Width * 0.35);
            var ch = Math.Max(PreviewRoiMinSizeDip, content.Height * 0.35);
            cw = Math.Min(cw, content.Width);
            ch = Math.Min(ch, content.Height);
            var lx = content.Left + (content.Width - cw) * 0.5;
            var ly = content.Top + (content.Height - ch) * 0.5;
            return new Rect(lx, ly, cw, ch);
        }

        private static Rect ClampPreviewRoiToContent(Rect r, Rect content)
        {
            var rw = Math.Min(Math.Max(r.Width, PreviewRoiMinSizeDip), content.Width);
            var rh = Math.Min(Math.Max(r.Height, PreviewRoiMinSizeDip), content.Height);
            var rx = Math.Max(content.Left, Math.Min(r.Left, content.Right - rw));
            var ry = Math.Max(content.Top, Math.Min(r.Top, content.Bottom - rh));
            return new Rect(rx, ry, rw, rh);
        }

        private PreviewRoiHitKind HitTestPreviewRoi(Point p, Rect r)
        {
            if (r.Width < 1 || r.Height < 1)
                return PreviewRoiHitKind.None;

            var cl = PreviewRoiCornerHitDip;
            var nearLeft = p.X <= r.Left + cl;
            var nearRight = p.X >= r.Right - cl;
            var nearTop = p.Y <= r.Top + cl;
            var nearBottom = p.Y >= r.Bottom - cl;

            if (nearTop && nearLeft)
                return PreviewRoiHitKind.ResizeNorthWest;
            if (nearTop && nearRight)
                return PreviewRoiHitKind.ResizeNorthEast;
            if (nearBottom && nearLeft)
                return PreviewRoiHitKind.ResizeSouthWest;
            if (nearBottom && nearRight)
                return PreviewRoiHitKind.ResizeSouthEast;

            var inH = p.X >= r.Left && p.X <= r.Right;
            var inV = p.Y >= r.Top && p.Y <= r.Bottom;

            if (nearTop && inH)
                return PreviewRoiHitKind.ResizeNorth;
            if (nearBottom && inH)
                return PreviewRoiHitKind.ResizeSouth;
            if (nearLeft && inV)
                return PreviewRoiHitKind.ResizeWest;
            if (nearRight && inV)
                return PreviewRoiHitKind.ResizeEast;

            if (inH && inV)
                return PreviewRoiHitKind.Move;

            return PreviewRoiHitKind.None;
        }

        private static Cursor CursorForPreviewRoiHit(PreviewRoiHitKind hit) => hit switch
        {
            PreviewRoiHitKind.Move => Cursors.SizeAll,
            PreviewRoiHitKind.ResizeNorth or PreviewRoiHitKind.ResizeSouth => Cursors.SizeNS,
            PreviewRoiHitKind.ResizeEast or PreviewRoiHitKind.ResizeWest => Cursors.SizeWE,
            PreviewRoiHitKind.ResizeNorthWest or PreviewRoiHitKind.ResizeSouthEast => Cursors.SizeNWSE,
            PreviewRoiHitKind.ResizeNorthEast or PreviewRoiHitKind.ResizeSouthWest => Cursors.SizeNESW,
            _ => Cursors.Arrow,
        };

        private Rect ComputePreviewRoiDuringDrag(Point currentMouse)
        {
            var dm = currentMouse - _previewRoiDragMouseStart;
            var r0 = _previewRoiDragRectStart;
            var b = _previewRoiDragContentBounds;
            const double minS = PreviewRoiMinSizeDip;

            double L = r0.Left;
            double T = r0.Top;
            double R = r0.Right;
            double B = r0.Bottom;

            switch (_previewRoiActiveHit)
            {
                case PreviewRoiHitKind.Move:
                    L += dm.X;
                    T += dm.Y;
                    R += dm.X;
                    B += dm.Y;
                    break;

                case PreviewRoiHitKind.ResizeNorthWest:
                    L += dm.X;
                    T += dm.Y;
                    break;
                case PreviewRoiHitKind.ResizeNorthEast:
                    R += dm.X;
                    T += dm.Y;
                    break;
                case PreviewRoiHitKind.ResizeSouthWest:
                    L += dm.X;
                    B += dm.Y;
                    break;
                case PreviewRoiHitKind.ResizeSouthEast:
                    R += dm.X;
                    B += dm.Y;
                    break;

                case PreviewRoiHitKind.ResizeNorth:
                    T += dm.Y;
                    break;
                case PreviewRoiHitKind.ResizeSouth:
                    B += dm.Y;
                    break;
                case PreviewRoiHitKind.ResizeWest:
                    L += dm.X;
                    break;
                case PreviewRoiHitKind.ResizeEast:
                    R += dm.X;
                    break;

                default:
                    return _previewRoiRectDips;
            }

            if (_previewRoiActiveHit == PreviewRoiHitKind.Move)
            {
                var w = R - L;
                var h = B - T;
                L = Math.Max(b.Left, Math.Min(L, b.Right - w));
                T = Math.Max(b.Top, Math.Min(T, b.Bottom - h));
                return new Rect(L, T, w, h);
            }

            L = Math.Min(L, R - minS);
            T = Math.Min(T, B - minS);
            L = Math.Max(L, b.Left);
            T = Math.Max(T, b.Top);
            R = Math.Max(R, L + minS);
            B = Math.Max(B, T + minS);
            R = Math.Min(R, b.Right);
            B = Math.Min(B, b.Bottom);

            if (R <= L + minS - 0.0001 || B <= T + minS - 0.0001)
                return ClampPreviewRoiToContent(_previewRoiDragRectStart, b);

            return new Rect(L, T, R - L, B - T);
        }

        private void UpdatePreviewRoiVisual()
        {
            if (PreviewRoiRectangle is null || RoiOverlayCanvas is null)
                return;

            if (!_previewRoiHasPlacement || _previewRoiRectDips.Width < 1 || _previewRoiRectDips.Height < 1)
            {
                PreviewRoiRectangle.Visibility = Visibility.Collapsed;
                ClearPreviewRoiSourceMapping();
                ResetFloatAlarmTracking();
                return;
            }

            PreviewRoiRectangle.Visibility = Visibility.Visible;
            PreviewRoiRectangle.Width = _previewRoiRectDips.Width;
            PreviewRoiRectangle.Height = _previewRoiRectDips.Height;
            Canvas.SetLeft(PreviewRoiRectangle, _previewRoiRectDips.Left);
            Canvas.SetTop(PreviewRoiRectangle, _previewRoiRectDips.Top);
            SyncPreviewRoiOverlayToSourceFrame();
        }

        private void ClearPreviewRoiSourceMapping()
        {
            _previewRoiSourceNormX = 0;
            _previewRoiSourceNormY = 0;
            _previewRoiSourceNormWidth = 0;
            _previewRoiSourceNormHeight = 0;
            _previewRoiSourcePixels = default;
        }

        /// <summary>
        /// 1280×720 뷰 안의 Uniform 영상 표시 영역 대비 ROI(DIP)를 원본 프레임 비율·픽셀로 갱신합니다.
        /// </summary>
        private void SyncPreviewRoiOverlayToSourceFrame()
        {
            if (!_previewRoiHasPlacement
                || _previewRoiRectDips.Width < 1
                || _previewRoiRectDips.Height < 1
                || _previewWidth <= 0
                || _previewHeight <= 0
                || !TryGetPreviewContentRect(out var content)
                || content.Width < 1
                || content.Height < 1)
            {
                ClearPreviewRoiSourceMapping();
                return;
            }

            var r = ClampPreviewRoiToContent(_previewRoiRectDips, content);
            _previewRoiSourceNormX = (r.Left - content.Left) / content.Width;
            _previewRoiSourceNormY = (r.Top - content.Top) / content.Height;
            _previewRoiSourceNormWidth = r.Width / content.Width;
            _previewRoiSourceNormHeight = r.Height / content.Height;

            var x = (int)Math.Round(_previewRoiSourceNormX * _previewWidth);
            var y = (int)Math.Round(_previewRoiSourceNormY * _previewHeight);
            var w = (int)Math.Round(_previewRoiSourceNormWidth * _previewWidth);
            var h = (int)Math.Round(_previewRoiSourceNormHeight * _previewHeight);

            x = Math.Clamp(x, 0, Math.Max(0, _previewWidth - 1));
            y = Math.Clamp(y, 0, Math.Max(0, _previewHeight - 1));
            w = Math.Clamp(w, 1, _previewWidth - x);
            h = Math.Clamp(h, 1, _previewHeight - y);
            _previewRoiSourcePixels = new Int32Rect(x, y, w, h);
        }

        private void TargetColorPickerControl_OnSelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            if (_suppressTargetColorPickerEvents)
                return;
            if (e.NewValue.HasValue)
            {
                _targetColor = e.NewValue.Value;
                PersistUserSettingsToFile();
            }
        }

        private void EyedropperToggleButton_OnChecked(object sender, RoutedEventArgs e) => _eyedropperActive = true;

        private void EyedropperToggleButton_OnUnchecked(object sender, RoutedEventArgs e) => _eyedropperActive = false;

        private void ColorToleranceSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            SyncColorToleranceFromSlider();
        }

        private void SyncColorToleranceFromSlider()
        {
            if (ColorToleranceSlider is null)
                return;
            var v = (int)Math.Round(ColorToleranceSlider.Value);
            v = Math.Clamp(v, 0, 50);
            _colorChannelTolerance = v;
            if (ColorToleranceValueText is not null)
                ColorToleranceValueText.Text = v.ToString(CultureInfo.InvariantCulture);
            PersistUserSettingsToFile();
        }

        private void ResetFloatAlarmTracking() => UpdateFloatAlarmSustainRequested(false);

        private void EnsureFloatAlarmSustainBeepWorker()
        {
            lock (_floatAlarmBeepThreadLock)
            {
                if (_floatAlarmBeepThread is { IsAlive: true })
                    return;
                _floatAlarmBeepWorkerShutdown = false;
                _floatAlarmBeepThread = new Thread(FloatAlarmSustainBeepThreadProc)
                {
                    IsBackground = true,
                    Name = "FloatAlarmSustainBeep",
                };
                _floatAlarmBeepThread.Start();
            }
        }

        private void StopFloatAlarmSustainBeepWorker()
        {
            _floatAlarmSustainRequested = false;
            _floatAlarmBeepWorkerShutdown = true;
            Thread? t;
            lock (_floatAlarmBeepThreadLock)
            {
                t = _floatAlarmBeepThread;
            }

            try
            {
                t?.Join(millisecondsTimeout: 1200);
            }
            catch
            {
                // ignore
            }

            lock (_floatAlarmBeepThreadLock)
            {
                if (ReferenceEquals(_floatAlarmBeepThread, t))
                    _floatAlarmBeepThread = null;
            }
        }

        private void FloatAlarmSustainBeepThreadProc()
        {
            const int toneMs = 380;
            while (!_floatAlarmBeepWorkerShutdown)
            {
                if (_floatAlarmSustainRequested)
                {
                    try
                    {
                        Beep(1320, toneMs);
                    }
                    catch
                    {
                        // ignore
                    }
                }
                else
                {
                    Thread.Sleep(20);
                }
            }
        }

        private void UpdateFloatAlarmSustainRequested(bool sustain)
        {
            _floatAlarmSustainRequested = sustain;
            EnsureFloatAlarmSustainBeepWorker();
        }

        private void UpdateFloatAlarmSustainFromMatchCount(long matchCount) =>
            UpdateFloatAlarmSustainRequested(matchCount <= FloatAlarmSustainMatchMax);

        private bool IsMaskPreviewMode() =>
            _tabMaskPreviewHold || (PreviewViewMaskRadio?.IsChecked == true);

        private void PreviewViewModeRadio_OnChecked(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded)
                return;
            UpdatePreviewImageSourceForCurrentMode();
        }

        private void UpdatePreviewImageSourceForCurrentMode()
        {
            if (PreviewImage is null)
                return;
            if (IsMaskPreviewMode() && _maskPreviewBitmap is not null)
                PreviewImage.Source = _maskPreviewBitmap;
            else if (_previewBitmap is not null)
                PreviewImage.Source = _previewBitmap;
        }

        /// <summary>BGRA 버퍼를 <see cref="ProcessFloatDetection"/>과 동일한 기준으로 이진(흰/검) 마스크로 그립니다.</summary>
        private void BlitColorMaskFromBgra(WriteableBitmap mask, byte[] bgra, int w, int h)
        {
            if (bgra.Length < w * h * 4 || mask.PixelWidth != w || mask.PixelHeight != h)
                return;

            var tr = _targetColor.R;
            var tg = _targetColor.G;
            var tb = _targetColor.B;
            var tol = _colorChannelTolerance;
            var srcStride = w * 4;

            mask.Lock();
            try
            {
                var dstStride = mask.BackBufferStride;
                var p0 = mask.BackBuffer;
                for (var y = 0; y < h; y++)
                {
                    var rowBase = y * srcStride;
                    var dstRow = p0 + y * dstStride;
                    for (var x = 0; x < w; x++)
                    {
                        var si = rowBase + x * 4;
                        var bb = bgra[si];
                        var gg = bgra[si + 1];
                        var rr = bgra[si + 2];
                        var write = (Math.Abs(rr - tr) <= tol && Math.Abs(gg - tg) <= tol && Math.Abs(bb - tb) <= tol)
                            ? unchecked((int)0xffffffff)
                            : unchecked((int)0xff000000);
                        Marshal.WriteInt32(dstRow + x * 4, write);
                    }
                }

                mask.AddDirtyRect(new Int32Rect(0, 0, w, h));
            }
            finally
            {
                mask.Unlock();
            }
        }

        private void ApplyPreviewDisplayModeAfterFrame(byte[] bgra, int w, int h)
        {
            if (_previewBitmap is null || w != _previewWidth || h != _previewHeight)
                return;
            if (IsMaskPreviewMode() && _maskPreviewBitmap is not null)
                BlitColorMaskFromBgra(_maskPreviewBitmap, bgra, w, h);
            UpdatePreviewImageSourceForCurrentMode();
        }

        /// <summary>
        /// ROI 내 유사색 픽셀이 <see cref="FloatAlarmSustainMatchMax"/>개 이하이면 지속 비프, 그보다 많으면 비프를 끕니다. BGRA, <paramref name="w"/>×<paramref name="h"/>.
        /// </summary>
        private void ProcessFloatDetection(byte[] bgra, int w, int h)
        {
            if (bgra.Length < w * h * 4)
            {
                UpdateFloatAlarmSustainRequested(false);
                return;
            }

            if (!_previewRoiHasPlacement || _previewRoiSourcePixels.Width < 1 || _previewRoiSourcePixels.Height < 1)
            {
                UpdateFloatAlarmSustainRequested(false);
                return;
            }

            var roi = _previewRoiSourcePixels;
            if (roi.X < 0 || roi.Y < 0 || roi.X + roi.Width > w || roi.Y + roi.Height > h)
            {
                UpdateFloatAlarmSustainRequested(false);
                return;
            }

            var tr = _targetColor.R;
            var tg = _targetColor.G;
            var tb = _targetColor.B;
            var tol = _colorChannelTolerance;
            var stride = w * 4;

            long matchCount = 0;
            for (var py = roi.Y; py < roi.Y + roi.Height; py++)
            {
                for (var px = roi.X; px < roi.X + roi.Width; px++)
                {
                    var i = py * stride + px * 4;
                    var bb = bgra[i];
                    var gg = bgra[i + 1];
                    var rr = bgra[i + 2];
                    if (Math.Abs(rr - tr) <= tol && Math.Abs(gg - tg) <= tol && Math.Abs(bb - tb) <= tol)
                        matchCount++;
                }
            }

            UpdateFloatAlarmSustainFromMatchCount(matchCount);
        }

        private void ApplyTargetColorToPicker()
        {
            if (TargetColorPickerControl is null)
                return;
            _suppressTargetColorPickerEvents = true;
            try
            {
                TargetColorPickerControl.SelectedColor = _targetColor;
            }
            finally
            {
                _suppressTargetColorPickerEvents = false;
            }
        }

        private bool TrySamplePreviewPixelColor(Point canvasPoint, out Color color)
        {
            color = default;
            if (_previewBitmap is null || _previewWidth <= 0 || _previewHeight <= 0)
                return false;
            if (!TryGetPreviewContentRect(out var content) || content.Width < 1 || content.Height < 1)
                return false;
            if (!content.Contains(canvasPoint))
                return false;

            var nx = (canvasPoint.X - content.Left) / content.Width;
            var ny = (canvasPoint.Y - content.Top) / content.Height;
            var bx = (int)Math.Clamp(Math.Floor(nx * _previewWidth), 0, _previewWidth - 1);
            var by = (int)Math.Clamp(Math.Floor(ny * _previewHeight), 0, _previewHeight - 1);

            var bytes = new byte[4];
            try
            {
                // WriteableBitmap에 Lock 후 CopyPixels를 호출하면 렌더/합성 스레드와 교착될 수 있어 Lock 없이 읽습니다.
                _previewBitmap.CopyPixels(new Int32Rect(bx, by, 1, 1), bytes, 4, 0);
            }
            catch
            {
                return false;
            }

            color = Color.FromArgb(bytes[3], bytes[2], bytes[1], bytes[0]);
            return true;
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

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool Beep(int frequency, int duration);

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
