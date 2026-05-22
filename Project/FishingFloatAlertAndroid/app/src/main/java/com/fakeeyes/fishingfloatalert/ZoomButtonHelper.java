package com.fakeeyes.fishingfloatalert;

import android.view.View;
import android.widget.TextView;

import androidx.annotation.NonNull;

final class ZoomButtonHelper {

    static final float[] ZOOM_LEVELS = {1f, 2f, 3f, 5f, 10f};

    private static final int[] ZOOM_BUTTON_IDS = {
            R.id.zoom_btn_1,
            R.id.zoom_btn_2,
            R.id.zoom_btn_3,
            R.id.zoom_btn_5,
            R.id.zoom_btn_10
    };

    private ZoomButtonHelper() {
    }

    static void setup(@NonNull View root) {
        float currentZoom = CameraPreviewManager.getInstance().getZoomRatio();
        int selectedIndex = 0;
        for (int i = 0; i < ZOOM_LEVELS.length; i++) {
            TextView button = root.findViewById(ZOOM_BUTTON_IDS[i]);
            if (button == null) {
                continue;
            }
            float zoom = ZOOM_LEVELS[i];
            if (Math.abs(currentZoom - zoom) < 0.01f) {
                selectedIndex = i;
            }
            button.setOnClickListener(v -> selectZoom(zoom, root));
        }
        updateStyles(root, selectedIndex);
    }

    private static void selectZoom(float zoom, @NonNull View root) {
        CameraPreviewManager.getInstance().setZoomRatio(zoom);
        int selectedIndex = 0;
        for (int i = 0; i < ZOOM_LEVELS.length; i++) {
            if (Math.abs(ZOOM_LEVELS[i] - zoom) < 0.01f) {
                selectedIndex = i;
                break;
            }
        }
        updateStyles(root, selectedIndex);
    }

    private static void updateStyles(@NonNull View root, int selectedIndex) {
        for (int i = 0; i < ZOOM_BUTTON_IDS.length; i++) {
            TextView button = root.findViewById(ZOOM_BUTTON_IDS[i]);
            if (button == null) {
                continue;
            }
            boolean selected = i == selectedIndex;
            button.setBackgroundResource(selected
                    ? R.drawable.bg_zoom_button_selected
                    : R.drawable.bg_zoom_button);
        }
    }
}
