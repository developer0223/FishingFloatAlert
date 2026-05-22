package com.fakeeyes.fishingfloatalert;

import android.os.Bundle;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;

public class ColorSettingsFragment extends FullscreenImageFragment {

    @Nullable
    @Override
    public View onCreateView(@NonNull LayoutInflater inflater, @Nullable ViewGroup container,
            @Nullable Bundle savedInstanceState) {
        return inflater.inflate(R.layout.fragment_color_settings, container, false);
    }

    @Override
    protected int getDefaultImageResId() {
        return R.drawable.bg_fragment_color_settings_placeholder;
    }
}
