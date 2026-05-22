package com.fakeeyes.fishingfloatalert;

import android.graphics.Bitmap;
import android.graphics.drawable.BitmapDrawable;
import android.graphics.drawable.ColorDrawable;
import android.os.Bundle;
import android.view.LayoutInflater;
import android.view.MotionEvent;
import android.view.View;
import android.view.ViewGroup;
import android.widget.ImageView;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;

import com.google.android.material.button.MaterialButton;

public class ColorSettingsFragment extends FullscreenImageFragment {

    @Nullable
    private View colorPreviewSwatch;

    @Nullable
    private CrosshairOverlayView crosshairOverlay;

    @Nullable
    private ImageView fullscreenImage;

    private boolean captureMode;
    private int pickedColor;

    @Nullable
    private Bitmap frozenBitmap;

    @Nullable
    @Override
    public View onCreateView(@NonNull LayoutInflater inflater, @Nullable ViewGroup container,
            @Nullable Bundle savedInstanceState) {
        return inflater.inflate(R.layout.fragment_color_settings, container, false);
    }

    @Override
    public void onViewCreated(@NonNull View view, @Nullable Bundle savedInstanceState) {
        super.onViewCreated(view, savedInstanceState);
        colorPreviewSwatch = view.findViewById(R.id.color_preview_swatch);
        crosshairOverlay = view.findViewById(R.id.crosshair_overlay);
        fullscreenImage = getFullscreenImageView();
        pickedColor = AppSettings.getInstance().getTargetColor();
        updateColorSwatch(pickedColor);

        ZoomButtonHelper.setup(view);

        MaterialButton captureButton = view.findViewById(R.id.btn_capture);
        MaterialButton applyButton = view.findViewById(R.id.btn_apply);
        captureButton.setOnClickListener(v -> enterCaptureMode());
        applyButton.setOnClickListener(v -> applyPickedColor());

        View touchTarget = crosshairOverlay != null ? crosshairOverlay : fullscreenImage;
        if (touchTarget != null) {
            touchTarget.setOnTouchListener(this::onPreviewTouched);
        }
    }

    @Override
    public void onResume() {
        super.onResume();
        View view = getView();
        if (view != null) {
            ZoomButtonHelper.setup(view);
        }
    }

    @Override
    public void onDestroyView() {
        exitCaptureMode();
        super.onDestroyView();
    }

    @Override
    public void onCameraFrame(@NonNull CameraFrame frame) {
        if (captureMode) {
            return;
        }
        setFullscreenImage(frame.getRawBitmap());
    }

    @Override
    protected int getDefaultImageResId() {
        return R.drawable.bg_fragment_color_settings_placeholder;
    }

    private void enterCaptureMode() {
        if (captureMode) {
            return;
        }
        ImageView imageView = getFullscreenImageView();
        if (imageView == null) {
            return;
        }
        Bitmap source = getDisplayBitmap(imageView);
        if (source == null || source.isRecycled()) {
            return;
        }
        frozenBitmap = CameraImageUtils.copyBitmap(source);
        if (frozenBitmap == null) {
            return;
        }
        captureMode = true;
        setFullscreenImage(frozenBitmap);
        if (crosshairOverlay != null) {
            crosshairOverlay.hideCrosshair();
        }
    }

    private void applyPickedColor() {
        if (!captureMode) {
            return;
        }
        AppSettings.getInstance().setTargetColor(pickedColor);
        exitCaptureMode();
    }

    private void exitCaptureMode() {
        captureMode = false;
        detachFrozenBitmapFromImageView();
        releaseFrozenBitmap();
        if (crosshairOverlay != null) {
            crosshairOverlay.hideCrosshair();
        }
    }

    private boolean onPreviewTouched(@NonNull View view, @NonNull MotionEvent event) {
        if (!captureMode || frozenBitmap == null || frozenBitmap.isRecycled()) {
            return false;
        }
        int action = event.getActionMasked();
        if (action == MotionEvent.ACTION_DOWN || action == MotionEvent.ACTION_MOVE) {
            updateColorPickAt(view, event);
            return true;
        }
        if (action == MotionEvent.ACTION_UP || action == MotionEvent.ACTION_CANCEL) {
            return true;
        }
        return false;
    }

    private void updateColorPickAt(@NonNull View touchSource, @NonNull MotionEvent event) {
        ImageView imageView = getFullscreenImageView();
        if (imageView == null) {
            return;
        }
        float[] mapped = mapTouchToImageView(imageView, touchSource, event.getX(), event.getY());
        Integer color = ImageViewTouchMapper.getColorAtTouch(
                imageView,
                frozenBitmap,
                mapped[0],
                mapped[1]);
        if (color == null) {
            return;
        }
        pickedColor = color;
        updateColorSwatch(pickedColor);
        if (crosshairOverlay != null) {
            crosshairOverlay.setCrosshairPosition(mapped[0], mapped[1]);
        }
    }

    @Nullable
    private Bitmap getDisplayBitmap(@NonNull ImageView imageView) {
        if (imageView.getDrawable() instanceof BitmapDrawable) {
            BitmapDrawable bitmapDrawable = (BitmapDrawable) imageView.getDrawable();
            Bitmap bitmap = bitmapDrawable.getBitmap();
            if (bitmap != null && !bitmap.isRecycled()) {
                return bitmap;
            }
        }
        return null;
    }

    private void detachFrozenBitmapFromImageView() {
        ImageView imageView = getFullscreenImageView();
        if (imageView == null || frozenBitmap == null) {
            return;
        }
        if (imageView.getDrawable() instanceof BitmapDrawable) {
            Bitmap displayed = ((BitmapDrawable) imageView.getDrawable()).getBitmap();
            if (displayed == frozenBitmap) {
                imageView.setImageDrawable(null);
            }
        }
    }

    @NonNull
    private float[] mapTouchToImageView(
            @NonNull ImageView imageView,
            @NonNull View touchSource,
            float localX,
            float localY) {
        int[] imageLoc = new int[2];
        int[] touchLoc = new int[2];
        imageView.getLocationOnScreen(imageLoc);
        touchSource.getLocationOnScreen(touchLoc);
        float x = localX + (touchLoc[0] - imageLoc[0]);
        float y = localY + (touchLoc[1] - imageLoc[1]);
        return new float[]{x, y};
    }

    private void updateColorSwatch(int color) {
        if (colorPreviewSwatch != null) {
            colorPreviewSwatch.setBackground(new ColorDrawable(color));
        }
    }

    private void releaseFrozenBitmap() {
        if (frozenBitmap != null) {
            if (!frozenBitmap.isRecycled()) {
                frozenBitmap.recycle();
            }
            frozenBitmap = null;
        }
    }
}
