package com.fakeeyes.fishingfloatalert;

import android.os.Bundle;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.TextView;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;

import com.google.android.material.button.MaterialButtonToggleGroup;

public class CameraFragment extends FullscreenImageFragment {

    @Nullable
    private MaterialButtonToggleGroup previewModeToggle;

    @Nullable
    private TextView pixelCountBadge;

    @Nullable
    @Override
    public View onCreateView(@NonNull LayoutInflater inflater, @Nullable ViewGroup container,
            @Nullable Bundle savedInstanceState) {
        return inflater.inflate(R.layout.fragment_camera, container, false);
    }

    @Override
    public void onViewCreated(@NonNull View view, @Nullable Bundle savedInstanceState) {
        super.onViewCreated(view, savedInstanceState);
        previewModeToggle = view.findViewById(R.id.preview_mode_toggle);
        pixelCountBadge = view.findViewById(R.id.pixel_count_badge);
        ZoomButtonHelper.setup(view);
        setupPreviewModeToggle();
    }

    @Override
    public void onResume() {
        super.onResume();
        View view = getView();
        if (view != null) {
            ZoomButtonHelper.setup(view);
        }
        syncPreviewModeUi();
    }

    @Override
    public void onCameraFrame(@NonNull CameraFrame frame) {
        setFullscreenImage(frame.getDisplayBitmap());
        updatePixelCountBadge(frame.getMatchCount());
    }

    @Override
    protected int getDefaultImageResId() {
        return R.drawable.bg_fragment_camera_placeholder;
    }

    private void setupPreviewModeToggle() {
        if (previewModeToggle == null) {
            return;
        }
        previewModeToggle.addOnButtonCheckedListener((group, checkedId, isChecked) -> {
            if (!isChecked) {
                return;
            }
            boolean filterMode = checkedId == R.id.preview_mode_filter;
            CameraPreviewManager.getInstance().setFilterModeEnabled(filterMode);
            updatePixelCountVisibility(filterMode);
        });
    }

    private void syncPreviewModeUi() {
        boolean filterMode = CameraPreviewManager.getInstance().isFilterModeEnabled();
        if (previewModeToggle != null) {
            int buttonId = filterMode ? R.id.preview_mode_filter : R.id.preview_mode_camera;
            if (previewModeToggle.getCheckedButtonId() != buttonId) {
                previewModeToggle.check(buttonId);
            }
        }
        updatePixelCountVisibility(filterMode);
    }

    private void updatePixelCountVisibility(boolean filterMode) {
        if (pixelCountBadge == null) {
            return;
        }
        pixelCountBadge.setVisibility(filterMode ? View.VISIBLE : View.GONE);
    }

    private void updatePixelCountBadge(int matchCount) {
        if (pixelCountBadge == null
                || pixelCountBadge.getVisibility() != View.VISIBLE) {
            return;
        }
        pixelCountBadge.setText(getString(R.string.pixel_count_format, matchCount));
    }
}
