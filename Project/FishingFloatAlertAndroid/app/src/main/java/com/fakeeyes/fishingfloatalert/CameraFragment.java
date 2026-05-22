package com.fakeeyes.fishingfloatalert;

import android.os.Bundle;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.TextView;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;

public class CameraFragment extends FullscreenImageFragment {

    private static final float[] ZOOM_LEVELS = {1f, 2f, 3f, 5f, 10f};

    private static final int[] ZOOM_BUTTON_IDS = {
            R.id.zoom_btn_1,
            R.id.zoom_btn_2,
            R.id.zoom_btn_3,
            R.id.zoom_btn_5,
            R.id.zoom_btn_10
    };

    @Nullable
    @Override
    public View onCreateView(@NonNull LayoutInflater inflater, @Nullable ViewGroup container,
            @Nullable Bundle savedInstanceState) {
        return inflater.inflate(R.layout.fragment_camera, container, false);
    }

    @Override
    public void onViewCreated(@NonNull View view, @Nullable Bundle savedInstanceState) {
        super.onViewCreated(view, savedInstanceState);
        setupZoomButtons(view);
    }

    @Override
    protected int getDefaultImageResId() {
        return R.drawable.bg_fragment_camera_placeholder;
    }

    private void setupZoomButtons(@NonNull View root) {
        float currentZoom = CameraPreviewManager.getInstance().getZoomRatio();
        int selectedIndex = 0;
        for (int i = 0; i < ZOOM_LEVELS.length; i++) {
            TextView button = root.findViewById(ZOOM_BUTTON_IDS[i]);
            float zoom = ZOOM_LEVELS[i];
            if (Math.abs(currentZoom - zoom) < 0.01f) {
                selectedIndex = i;
            }
            button.setOnClickListener(v -> selectZoom(zoom, root));
        }
        updateZoomButtonStyles(root, selectedIndex);
    }

    private void selectZoom(float zoom, @NonNull View root) {
        CameraPreviewManager.getInstance().setZoomRatio(zoom);
        int selectedIndex = 0;
        for (int i = 0; i < ZOOM_LEVELS.length; i++) {
            if (Math.abs(ZOOM_LEVELS[i] - zoom) < 0.01f) {
                selectedIndex = i;
                break;
            }
        }
        updateZoomButtonStyles(root, selectedIndex);
    }

    private void updateZoomButtonStyles(@NonNull View root, int selectedIndex) {
        for (int i = 0; i < ZOOM_BUTTON_IDS.length; i++) {
            TextView button = root.findViewById(ZOOM_BUTTON_IDS[i]);
            boolean selected = i == selectedIndex;
            button.setBackgroundResource(selected
                    ? R.drawable.bg_zoom_button_selected
                    : R.drawable.bg_zoom_button);
        }
    }
}
