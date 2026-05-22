package com.fakeeyes.fishingfloatalert;

import android.graphics.Bitmap;
import android.graphics.Matrix;
import android.graphics.RectF;
import android.widget.ImageView;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;

final class ImageViewTouchMapper {

    private ImageViewTouchMapper() {
    }

    static boolean mapTouchToBitmapPixel(
            @NonNull ImageView imageView,
            @NonNull Bitmap bitmap,
            float touchX,
            float touchY,
            @NonNull int[] outPixel) {
        if (bitmap.getWidth() < 1 || bitmap.getHeight() < 1) {
            return false;
        }
        Matrix matrix = new Matrix(imageView.getImageMatrix());
        RectF imageRect = new RectF(0f, 0f, bitmap.getWidth(), bitmap.getHeight());
        matrix.mapRect(imageRect);
        float paddingLeft = imageView.getPaddingLeft();
        float paddingTop = imageView.getPaddingTop();
        float localX = touchX - paddingLeft;
        float localY = touchY - paddingTop;
        if (!imageRect.contains(localX, localY)) {
            return false;
        }
        float nx = (localX - imageRect.left) / imageRect.width();
        float ny = (localY - imageRect.top) / imageRect.height();
        outPixel[0] = Math.max(0, Math.min(bitmap.getWidth() - 1, Math.round(nx * (bitmap.getWidth() - 1))));
        outPixel[1] = Math.max(0, Math.min(bitmap.getHeight() - 1, Math.round(ny * (bitmap.getHeight() - 1))));
        return true;
    }

    @Nullable
    static Integer getColorAtTouch(
            @NonNull ImageView imageView,
            @NonNull Bitmap bitmap,
            float touchX,
            float touchY) {
        int[] pixel = new int[2];
        if (!mapTouchToBitmapPixel(imageView, bitmap, touchX, touchY, pixel)) {
            return null;
        }
        return ColorMaskProcessor.getColorAt(bitmap, pixel[0], pixel[1]);
    }
}
