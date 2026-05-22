package com.fakeeyes.fishingfloatalert;

import android.graphics.Bitmap;

import androidx.annotation.NonNull;

public interface CameraFrameListener {

    void onCameraFrame(@NonNull Bitmap frame);
}
