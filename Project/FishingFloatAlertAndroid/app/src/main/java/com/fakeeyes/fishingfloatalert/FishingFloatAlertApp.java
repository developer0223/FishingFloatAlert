package com.fakeeyes.fishingfloatalert;

import android.app.Application;

public class FishingFloatAlertApp extends Application {

    @Override
    public void onCreate() {
        super.onCreate();
        AppSettings.init(this);
        ColorMaskProcessor.ensureOpenCvLoaded();
        FloatAlarmManager.getInstance();
    }
}
