package com.fakeeyes.fishingfloatalert;

import android.content.Context;
import android.content.SharedPreferences;
import android.graphics.Color;

import androidx.annotation.NonNull;

public final class AppSettings {

    private static final String PREFS_NAME = "fishing_float_alert_settings";
    private static final String KEY_TARGET_COLOR = "target_color";
    private static final String KEY_COLOR_TOLERANCE = "color_tolerance";
    private static final String KEY_ALARM_ENABLED = "alarm_enabled";
    private static final String KEY_ALARM_MATCH_MAX = "alarm_match_max";

    private static final int DEFAULT_TARGET_COLOR = Color.RED;
    private static final int DEFAULT_COLOR_TOLERANCE = 12;
    private static final int DEFAULT_ALARM_MATCH_MAX = 5;

    private static volatile AppSettings instance;

    @NonNull
    private final SharedPreferences preferences;

    private AppSettings(@NonNull Context context) {
        preferences = context.getApplicationContext()
                .getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE);
    }

    public static void init(@NonNull Context context) {
        if (instance == null) {
            synchronized (AppSettings.class) {
                if (instance == null) {
                    instance = new AppSettings(context);
                }
            }
        }
    }

    @NonNull
    public static AppSettings getInstance() {
        if (instance == null) {
            throw new IllegalStateException("AppSettings.init() must be called first");
        }
        return instance;
    }

    public int getTargetColor() {
        return preferences.getInt(KEY_TARGET_COLOR, DEFAULT_TARGET_COLOR);
    }

    public void setTargetColor(int color) {
        preferences.edit().putInt(KEY_TARGET_COLOR, color).apply();
    }

    public int getColorTolerance() {
        return preferences.getInt(KEY_COLOR_TOLERANCE, DEFAULT_COLOR_TOLERANCE);
    }

    public void setColorTolerance(int tolerance) {
        int clamped = Math.max(1, Math.min(50, tolerance));
        preferences.edit().putInt(KEY_COLOR_TOLERANCE, clamped).apply();
    }

    public boolean isAlarmEnabled() {
        return preferences.getBoolean(KEY_ALARM_ENABLED, true);
    }

    public void setAlarmEnabled(boolean enabled) {
        preferences.edit().putBoolean(KEY_ALARM_ENABLED, enabled).apply();
    }

    public int getAlarmMatchMax() {
        return preferences.getInt(KEY_ALARM_MATCH_MAX, DEFAULT_ALARM_MATCH_MAX);
    }

    public void setAlarmMatchMax(int max) {
        int clamped = Math.max(1, Math.min(50, max));
        preferences.edit().putInt(KEY_ALARM_MATCH_MAX, clamped).apply();
    }
}
