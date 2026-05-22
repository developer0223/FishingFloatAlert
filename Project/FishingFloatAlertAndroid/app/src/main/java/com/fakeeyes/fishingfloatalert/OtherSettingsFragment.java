package com.fakeeyes.fishingfloatalert;

import android.os.Bundle;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.TextView;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.fragment.app.Fragment;

import com.google.android.material.materialswitch.MaterialSwitch;
import com.google.android.material.slider.Slider;

public class OtherSettingsFragment extends Fragment {

    @Nullable
    private MaterialSwitch alarmSwitch;

    @Nullable
    private Slider colorToleranceSlider;

    @Nullable
    private Slider alarmMatchMaxSlider;

    @Nullable
    private TextView colorToleranceValueText;

    @Nullable
    private TextView alarmMatchMaxValueText;

    @Nullable
    @Override
    public View onCreateView(@NonNull LayoutInflater inflater, @Nullable ViewGroup container,
            @Nullable Bundle savedInstanceState) {
        return inflater.inflate(R.layout.fragment_other_settings, container, false);
    }

    @Override
    public void onViewCreated(@NonNull View view, @Nullable Bundle savedInstanceState) {
        super.onViewCreated(view, savedInstanceState);
        AppSettings settings = AppSettings.getInstance();

        alarmSwitch = view.findViewById(R.id.switch_alarm_enabled);
        colorToleranceSlider = view.findViewById(R.id.slider_color_tolerance);
        alarmMatchMaxSlider = view.findViewById(R.id.slider_alarm_match_max);
        colorToleranceValueText = view.findViewById(R.id.text_color_tolerance_value);
        alarmMatchMaxValueText = view.findViewById(R.id.text_alarm_match_max_value);

        alarmSwitch.setChecked(settings.isAlarmEnabled());
        alarmSwitch.setOnCheckedChangeListener((buttonView, isChecked) ->
                settings.setAlarmEnabled(isChecked));

        colorToleranceSlider.setValue(settings.getColorTolerance());
        updateColorToleranceLabel(settings.getColorTolerance());
        colorToleranceSlider.addOnChangeListener((slider, value, fromUser) -> {
            int tolerance = Math.round(value);
            settings.setColorTolerance(tolerance);
            updateColorToleranceLabel(tolerance);
        });

        alarmMatchMaxSlider.setValue(settings.getAlarmMatchMax());
        updateAlarmMatchMaxLabel(settings.getAlarmMatchMax());
        alarmMatchMaxSlider.addOnChangeListener((slider, value, fromUser) -> {
            int max = Math.round(value);
            settings.setAlarmMatchMax(max);
            updateAlarmMatchMaxLabel(max);
        });
    }

    private void updateColorToleranceLabel(int value) {
        if (colorToleranceValueText != null) {
            colorToleranceValueText.setText(
                    getString(R.string.settings_color_tolerance_value, value));
        }
    }

    private void updateAlarmMatchMaxLabel(int value) {
        if (alarmMatchMaxValueText != null) {
            alarmMatchMaxValueText.setText(
                    getString(R.string.settings_alarm_match_max_value, value));
        }
    }
}
