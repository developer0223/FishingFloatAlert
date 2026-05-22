package com.fakeeyes.fishingfloatalert;

import android.Manifest;
import android.content.Context;
import android.graphics.Bitmap;
import android.graphics.ImageFormat;
import android.graphics.Rect;
import android.os.Build;
import android.hardware.camera2.CameraAccessException;
import android.hardware.camera2.CameraCaptureSession;
import android.hardware.camera2.CameraCharacteristics;
import android.hardware.camera2.CameraDevice;
import android.hardware.camera2.CameraManager;
import android.hardware.camera2.CaptureRequest;
import android.hardware.camera2.params.StreamConfigurationMap;
import android.media.Image;
import android.media.ImageReader;
import android.os.Handler;
import android.os.HandlerThread;
import android.util.Range;
import android.util.Size;
import android.view.Surface;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.core.content.ContextCompat;
import androidx.fragment.app.FragmentActivity;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.Set;
import java.util.concurrent.CopyOnWriteArraySet;

public final class CameraPreviewManager {

    private static final long TARGET_FRAME_INTERVAL_MS = 25L;
    private static final long CAMERA_CLOSE_DELAY_MS = 250L;

    private static final CameraPreviewManager INSTANCE = new CameraPreviewManager();

    private final Set<CameraFrameListener> listeners = new CopyOnWriteArraySet<>();
    private final Handler mainHandler = new Handler(android.os.Looper.getMainLooper());

    @Nullable
    private Context appContext;
    private boolean permissionGranted;
    @Nullable
    private FragmentActivity hostActivity;

    @Nullable
    private HandlerThread cameraThread;
    @Nullable
    private Handler cameraHandler;
    @Nullable
    private ImageReader imageReader;
    @Nullable
    private CameraDevice cameraDevice;
    @Nullable
    private CameraCaptureSession captureSession;
    @Nullable
    private String cameraId;
    private int previewRotationDegrees;
    private long lastFrameDeliveredMs;
    private float zoomRatio = 1f;
    private volatile boolean filterModeEnabled;

    private final Runnable delayedCloseRunnable = this::closeCamera;

    private CameraPreviewManager() {
    }

    @NonNull
    public static CameraPreviewManager getInstance() {
        return INSTANCE;
    }

    public void init(@NonNull Context context) {
        appContext = context.getApplicationContext();
    }

    public void setPermissionGranted(boolean granted) {
        permissionGranted = granted;
        if (granted) {
            maybeOpenCamera();
        } else {
            closeCamera();
        }
    }

    public void addListener(@NonNull CameraFrameListener listener, @NonNull FragmentActivity activity) {
        mainHandler.removeCallbacks(delayedCloseRunnable);
        listeners.add(listener);
        hostActivity = activity;
        maybeOpenCamera();
    }

    public void removeListener(@NonNull CameraFrameListener listener) {
        listeners.remove(listener);
        if (listeners.isEmpty()) {
            hostActivity = null;
            mainHandler.postDelayed(delayedCloseRunnable, CAMERA_CLOSE_DELAY_MS);
        }
    }

    public float getZoomRatio() {
        return zoomRatio;
    }

    public void setZoomRatio(float ratio) {
        zoomRatio = Math.max(1f, ratio);
        if (cameraHandler != null) {
            cameraHandler.post(this::updateRepeatingRequest);
        }
    }

    public boolean isFilterModeEnabled() {
        return filterModeEnabled;
    }

    public void setFilterModeEnabled(boolean enabled) {
        filterModeEnabled = enabled;
    }

    private void maybeOpenCamera() {
        mainHandler.removeCallbacks(delayedCloseRunnable);
        if (!permissionGranted || listeners.isEmpty() || appContext == null) {
            return;
        }
        if (ContextCompat.checkSelfPermission(appContext, Manifest.permission.CAMERA)
                != android.content.pm.PackageManager.PERMISSION_GRANTED) {
            return;
        }
        if (cameraDevice != null) {
            return;
        }
        startBackgroundThread();
        openBackCamera();
    }

    private void startBackgroundThread() {
        if (cameraThread != null) {
            return;
        }
        cameraThread = new HandlerThread("CameraPreview");
        cameraThread.start();
        cameraHandler = new Handler(cameraThread.getLooper());
    }

    private void stopBackgroundThread() {
        if (cameraThread == null) {
            return;
        }
        cameraThread.quitSafely();
        try {
            cameraThread.join();
        } catch (InterruptedException ignored) {
            Thread.currentThread().interrupt();
        }
        cameraThread = null;
        cameraHandler = null;
    }

    private void openBackCamera() {
        if (appContext == null || cameraHandler == null) {
            return;
        }
        CameraManager cameraManager = appContext.getSystemService(CameraManager.class);
        if (cameraManager == null) {
            return;
        }
        try {
            cameraId = findBackCameraId(cameraManager);
            if (cameraId == null) {
                return;
            }
            CameraCharacteristics characteristics = cameraManager.getCameraCharacteristics(cameraId);
            previewRotationDegrees = computePreviewRotation(characteristics);
            Size previewSize = choosePreviewSize(characteristics);
            if (imageReader != null) {
                imageReader.close();
            }
            imageReader = ImageReader.newInstance(
                    previewSize.getWidth(),
                    previewSize.getHeight(),
                    ImageFormat.YUV_420_888,
                    2);
            imageReader.setOnImageAvailableListener(this::onImageAvailable, cameraHandler);
            cameraManager.openCamera(cameraId, cameraStateCallback, cameraHandler);
        } catch (CameraAccessException | SecurityException ignored) {
            closeCamera();
        }
    }

    @Nullable
    private String findBackCameraId(@NonNull CameraManager cameraManager) throws CameraAccessException {
        for (String id : cameraManager.getCameraIdList()) {
            CameraCharacteristics characteristics = cameraManager.getCameraCharacteristics(id);
            Integer facing = characteristics.get(CameraCharacteristics.LENS_FACING);
            if (facing != null && facing == CameraCharacteristics.LENS_FACING_BACK) {
                return id;
            }
        }
        return null;
    }

    @NonNull
    private Size choosePreviewSize(@NonNull CameraCharacteristics characteristics) {
        StreamConfigurationMap map = characteristics.get(CameraCharacteristics.SCALER_STREAM_CONFIGURATION_MAP);
        if (map == null) {
            return new Size(640, 480);
        }
        Size[] sizes = map.getOutputSizes(ImageFormat.YUV_420_888);
        if (sizes == null || sizes.length == 0) {
            return new Size(640, 480);
        }
        List<Size> candidates = new ArrayList<>();
        Collections.addAll(candidates, sizes);
        // 4:3 비율은 16:9보다 센서 영역을 넓게 쓰는 경우가 많아 시야가 넓어짐
        Size bestFourThree = null;
        long bestArea = 0;
        for (Size size : candidates) {
            int w = size.getWidth();
            int h = size.getHeight();
            boolean isFourThree = w * 3 == h * 4 || h * 3 == w * 4;
            if (!isFourThree) {
                continue;
            }
            long area = (long) w * h;
            if (area > bestArea && area <= 2073600L) {
                bestArea = area;
                bestFourThree = size;
            }
        }
        if (bestFourThree != null) {
            return bestFourThree;
        }
        Size best = candidates.get(0);
        long maxArea = 0;
        for (Size size : candidates) {
            long area = (long) size.getWidth() * size.getHeight();
            if (area > maxArea && area <= 2073600L) {
                maxArea = area;
                best = size;
            }
        }
        return best;
    }

    private int computePreviewRotation(@NonNull CameraCharacteristics characteristics) {
        Integer sensorOrientation = characteristics.get(CameraCharacteristics.SENSOR_ORIENTATION);
        if (sensorOrientation == null || hostActivity == null) {
            return 0;
        }
        int deviceRotation = hostActivity.getWindowManager().getDefaultDisplay().getRotation();
        int degrees;
        if (deviceRotation == Surface.ROTATION_90) {
            degrees = 90;
        } else if (deviceRotation == Surface.ROTATION_180) {
            degrees = 180;
        } else if (deviceRotation == Surface.ROTATION_270) {
            degrees = 270;
        } else {
            degrees = 0;
        }
        Integer lensFacing = characteristics.get(CameraCharacteristics.LENS_FACING);
        if (lensFacing != null && lensFacing == CameraCharacteristics.LENS_FACING_FRONT) {
            return (sensorOrientation + degrees) % 360;
        }
        return (sensorOrientation - degrees + 360) % 360;
    }

    private final CameraDevice.StateCallback cameraStateCallback = new CameraDevice.StateCallback() {
        @Override
        public void onOpened(@NonNull CameraDevice camera) {
            cameraDevice = camera;
            createCaptureSession();
        }

        @Override
        public void onDisconnected(@NonNull CameraDevice camera) {
            camera.close();
            cameraDevice = null;
        }

        @Override
        public void onError(@NonNull CameraDevice camera, int error) {
            camera.close();
            cameraDevice = null;
        }
    };

    private void createCaptureSession() {
        if (cameraDevice == null || imageReader == null || cameraHandler == null || appContext == null
                || cameraId == null) {
            return;
        }
        try {
            CameraManager cameraManager = appContext.getSystemService(CameraManager.class);
            if (cameraManager == null) {
                return;
            }
            CameraCharacteristics characteristics = cameraManager.getCameraCharacteristics(cameraId);
            Range<Integer> fpsRange = chooseFpsRange(
                    characteristics.get(CameraCharacteristics.CONTROL_AE_AVAILABLE_TARGET_FPS_RANGES));
            Surface surface = imageReader.getSurface();
            cameraDevice.createCaptureSession(
                    Collections.singletonList(surface),
                    new CameraCaptureSession.StateCallback() {
                        @Override
                        public void onConfigured(@NonNull CameraCaptureSession session) {
                            captureSession = session;
                            startRepeatingRequest(characteristics, fpsRange);
                        }

                        @Override
                        public void onConfigureFailed(@NonNull CameraCaptureSession session) {
                            closeCamera();
                        }
                    },
                    cameraHandler);
        } catch (CameraAccessException ignored) {
            closeCamera();
        }
    }

    private void startRepeatingRequest(
            @NonNull CameraCharacteristics characteristics,
            @Nullable Range<Integer> fpsRange) {
        updateRepeatingRequest(characteristics, fpsRange);
    }

    private void updateRepeatingRequest() {
        if (appContext == null || cameraId == null) {
            return;
        }
        CameraManager cameraManager = appContext.getSystemService(CameraManager.class);
        if (cameraManager == null) {
            return;
        }
        try {
            CameraCharacteristics characteristics = cameraManager.getCameraCharacteristics(cameraId);
            Range<Integer> fpsRange = chooseFpsRange(
                    characteristics.get(CameraCharacteristics.CONTROL_AE_AVAILABLE_TARGET_FPS_RANGES));
            updateRepeatingRequest(characteristics, fpsRange);
        } catch (CameraAccessException ignored) {
            closeCamera();
        }
    }

    private void updateRepeatingRequest(
            @NonNull CameraCharacteristics characteristics,
            @Nullable Range<Integer> fpsRange) {
        if (cameraDevice == null || captureSession == null || imageReader == null || cameraHandler == null) {
            return;
        }
        try {
            CaptureRequest.Builder builder = cameraDevice.createCaptureRequest(CameraDevice.TEMPLATE_PREVIEW);
            builder.addTarget(imageReader.getSurface());
            builder.set(CaptureRequest.CONTROL_MODE, CaptureRequest.CONTROL_MODE_AUTO);
            builder.set(CaptureRequest.CONTROL_AF_MODE,
                    CaptureRequest.CONTROL_AF_MODE_CONTINUOUS_PICTURE);
            if (fpsRange != null) {
                builder.set(CaptureRequest.CONTROL_AE_TARGET_FPS_RANGE, fpsRange);
            }
            applyZoom(builder, characteristics);
            captureSession.setRepeatingRequest(builder.build(), null, cameraHandler);
        } catch (CameraAccessException ignored) {
            closeCamera();
        }
    }

    private void applyZoom(
            @NonNull CaptureRequest.Builder builder,
            @NonNull CameraCharacteristics characteristics) {
        float clampedZoom = clampZoom(zoomRatio, characteristics);
        zoomRatio = clampedZoom;
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
            builder.set(CaptureRequest.CONTROL_ZOOM_RATIO, clampedZoom);
        } else {
            builder.set(CaptureRequest.SCALER_CROP_REGION, buildCropRegion(characteristics, clampedZoom));
        }
    }

    private float clampZoom(float requested, @NonNull CameraCharacteristics characteristics) {
        float maxZoom = getMaxDigitalZoom(characteristics);
        return Math.max(1f, Math.min(requested, maxZoom));
    }

    private float getMaxDigitalZoom(@NonNull CameraCharacteristics characteristics) {
        Float maxZoom = characteristics.get(CameraCharacteristics.SCALER_AVAILABLE_MAX_DIGITAL_ZOOM);
        return maxZoom != null && maxZoom >= 1f ? maxZoom : 1f;
    }

    @NonNull
    private static Rect buildCropRegion(
            @NonNull CameraCharacteristics characteristics,
            float zoom) {
        Rect activeArray = characteristics.get(CameraCharacteristics.SENSOR_INFO_ACTIVE_ARRAY_SIZE);
        if (activeArray == null) {
            return new Rect();
        }
        int cropWidth = (int) (activeArray.width() / zoom);
        int cropHeight = (int) (activeArray.height() / zoom);
        int left = activeArray.left + (activeArray.width() - cropWidth) / 2;
        int top = activeArray.top + (activeArray.height() - cropHeight) / 2;
        return new Rect(left, top, left + cropWidth, top + cropHeight);
    }

    @Nullable
    private static Range<Integer> chooseFpsRange(@Nullable Range<Integer>[] ranges) {
        if (ranges == null || ranges.length == 0) {
            return null;
        }
        Range<Integer> exactTen = null;
        Range<Integer> containsTen = null;
        Range<Integer> lowest = ranges[0];
        for (Range<Integer> range : ranges) {
            if (range.getLower() == 10 && range.getUpper() == 10) {
                exactTen = range;
            }
            if (range.getLower() <= 10 && range.getUpper() >= 10) {
                if (containsTen == null || range.getUpper() < containsTen.getUpper()) {
                    containsTen = range;
                }
            }
            if (range.getUpper() < lowest.getUpper()
                    || (range.getUpper().equals(lowest.getUpper()) && range.getLower() < lowest.getLower())) {
                lowest = range;
            }
        }
        if (exactTen != null) {
            return exactTen;
        }
        if (containsTen != null) {
            return containsTen;
        }
        return lowest;
    }

    private void onImageAvailable(@NonNull ImageReader reader) {
        Image image = reader.acquireLatestImage();
        if (image == null) {
            return;
        }
        long now = System.currentTimeMillis();
        if (now - lastFrameDeliveredMs < TARGET_FRAME_INTERVAL_MS) {
            image.close();
            return;
        }
        lastFrameDeliveredMs = now;
        Bitmap rawBitmap = CameraImageUtils.imageToSquareDisplayBitmap(image, previewRotationDegrees);
        image.close();
        if (rawBitmap == null || listeners.isEmpty()) {
            if (rawBitmap != null) {
                rawBitmap.recycle();
            }
            return;
        }
        processAndDispatchFrame(rawBitmap);
    }

    private void processAndDispatchFrame(@NonNull Bitmap rawBitmap) {
        AppSettings settings = AppSettings.getInstance();
        int targetColor = settings.getTargetColor();
        int tolerance = settings.getColorTolerance();
        int matchCount = ColorMaskProcessor.countMatchingPixels(rawBitmap, targetColor, tolerance);
        FloatAlarmManager.getInstance().updateFromMatchCount(matchCount);

        Bitmap displayBitmap;
        if (filterModeEnabled) {
            displayBitmap = ColorMaskProcessor.createMaskBitmap(rawBitmap, targetColor, tolerance);
        } else {
            displayBitmap = rawBitmap;
        }

        CameraFrame frame = new CameraFrame(rawBitmap, displayBitmap, matchCount);
        mainHandler.post(() -> dispatchFrame(frame));
    }

    private void dispatchFrame(@NonNull CameraFrame frame) {
        if (listeners.isEmpty()) {
            recycleFrame(frame);
            return;
        }
        for (CameraFrameListener listener : listeners) {
            listener.onCameraFrame(frame);
        }
    }

    private static void recycleFrame(@NonNull CameraFrame frame) {
        if (frame.getRawBitmap() != frame.getDisplayBitmap()) {
            frame.getDisplayBitmap().recycle();
        }
        frame.getRawBitmap().recycle();
    }

    private void closeCamera() {
        if (captureSession != null) {
            try {
                captureSession.stopRepeating();
                captureSession.close();
            } catch (Exception ignored) {
            }
            captureSession = null;
        }
        if (cameraDevice != null) {
            cameraDevice.close();
            cameraDevice = null;
        }
        if (imageReader != null) {
            imageReader.close();
            imageReader = null;
        }
        stopBackgroundThread();
        lastFrameDeliveredMs = 0L;
    }
}
