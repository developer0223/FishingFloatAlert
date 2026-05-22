package com.fakeeyes.fishingfloatalert;

import android.graphics.Bitmap;

import androidx.annotation.NonNull;

public final class CameraFrame {

    @NonNull
    private final Bitmap rawBitmap;

    @NonNull
    private final Bitmap displayBitmap;

    private final int matchCount;

    public CameraFrame(
            @NonNull Bitmap rawBitmap,
            @NonNull Bitmap displayBitmap,
            int matchCount) {
        this.rawBitmap = rawBitmap;
        this.displayBitmap = displayBitmap;
        this.matchCount = matchCount;
    }

    @NonNull
    public Bitmap getRawBitmap() {
        return rawBitmap;
    }

    @NonNull
    public Bitmap getDisplayBitmap() {
        return displayBitmap;
    }

    public int getMatchCount() {
        return matchCount;
    }
}
