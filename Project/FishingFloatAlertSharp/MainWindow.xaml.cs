using System;
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
        private int _previewWidth;
        private int _previewHeight;
        private int _previewBlitScheduled;
        private long _lastNullSoftwareBitmapLogTick;
        private long _lastFrameProcessErrorLogTick;

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

        private void MainWindow_OnClosing(object? sender, CancelEventArgs e)
        {
            StopPreviewSync();
        }

        private async void RefreshCamerasButton_OnClick(object sender, RoutedEventArgs e)
        {
            await RefreshCameraListAsync();
        }

        private async Task RefreshCameraListAsync()
        {
            await StopPreviewAsync();
            RefreshCamerasButton.IsEnabled = false;
            try
            {
                // MediaDevice 셀렉터: USB 웹캠·OBS Virtual Camera 등 WinRT에 노출되는 캡처 장치를 더 폭넓게 포함합니다.
                var selector = MediaDevice.GetVideoCaptureSelector();
                var devices = (await DeviceInformation.FindAllAsync(selector).AsTask())
                    .GroupBy(d => d.Id)
                    .Select(g => g.First())
                    .ToList();

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
            catch (Exception ex)
            {
                AppDiagnostics.LogError("카메라 목록 새로고침", ex);
            }
            finally
            {
                RefreshCamerasButton.IsEnabled = true;
            }
        }

        private async void CameraCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CameraCombo.SelectedItem is CameraComboEntry entry)
                selectedCameraId = entry.SelectedId;
            else
                selectedCameraId = -1;

            await RestartPreviewForSelectionAsync();
        }

        private async Task RestartPreviewForSelectionAsync()
        {
            await StopPreviewAsync();
            if (selectedCameraId < 0)
                return;
            if (CameraCombo.SelectedItem is not CameraComboEntry { DeviceInformationId: { } deviceId })
                return;

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
            _previewWidth = 0;
            _previewHeight = 0;
        }

        private void StopPreviewSync()
        {
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
            _previewWidth = 0;
            _previewHeight = 0;
        }

        private async Task StartPreviewAsync(string deviceId)
        {
            var capture = await InitializeMediaCaptureAsync(deviceId);
            _mediaCapture = capture;

            var frameSource = PickFrameSource(capture);
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

        private static MediaFrameSource? PickFrameSource(MediaCapture capture)
        {
            if (capture.FrameSources.Count == 0)
                return null;

            foreach (var s in capture.FrameSources.Values)
            {
                if (s.Info.MediaStreamType == MediaStreamType.VideoPreview)
                    return s;
            }

            foreach (var s in capture.FrameSources.Values)
            {
                if (s.Info.MediaStreamType == MediaStreamType.VideoRecord)
                    return s;
            }

            MediaFrameSource? best = null;
            var bestArea = 0;
            foreach (var s in capture.FrameSources.Values)
            {
                var area = s.SupportedFormats
                    .Where(f => f.VideoFormat != null)
                    .Select(f => (int)(f.VideoFormat!.Width * f.VideoFormat.Height))
                    .DefaultIfEmpty(0)
                    .Max();
                if (area > bestArea)
                {
                    bestArea = area;
                    best = s;
                }
            }

            return best ?? capture.FrameSources.Values.FirstOrDefault();
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
                        _previewWidth = w;
                        _previewHeight = h;
                        PreviewImage.Source = _previewBitmap;
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
