package com.fakeeyes.fishingfloatalert;

import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.graphics.ImageFormat;
import android.graphics.Matrix;
import android.graphics.Rect;
import android.graphics.YuvImage;
import android.media.Image;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;

import java.io.ByteArrayOutputStream;
import java.nio.ByteBuffer;

final class CameraImageUtils {

    private CameraImageUtils() {
    }

    @Nullable
    static Bitmap imageToDisplayBitmap(@NonNull Image image, int rotationDegrees) {
        Bitmap bitmap = yuv420888ToBitmap(image);
        if (bitmap == null) {
            return null;
        }
        if (rotationDegrees == 0) {
            return bitmap;
        }
        Bitmap rotated = rotateBitmap(bitmap, rotationDegrees);
        if (rotated != bitmap) {
            bitmap.recycle();
        }
        return rotated;
    }

    @Nullable
    private static Bitmap yuv420888ToBitmap(@NonNull Image image) {
        if (image.getFormat() != ImageFormat.YUV_420_888) {
            return null;
        }
        int width = image.getWidth();
        int height = image.getHeight();
        byte[] nv21 = yuv420888ToNv21(image);
        YuvImage yuvImage = new YuvImage(nv21, ImageFormat.NV21, width, height, null);
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        if (!yuvImage.compressToJpeg(new Rect(0, 0, width, height), 90, outputStream)) {
            return null;
        }
        byte[] jpegBytes = outputStream.toByteArray();
        return BitmapFactory.decodeByteArray(jpegBytes, 0, jpegBytes.length);
    }

    @NonNull
    private static byte[] yuv420888ToNv21(@NonNull Image image) {
        int width = image.getWidth();
        int height = image.getHeight();
        int ySize = width * height;
        int uvSize = width * height / 4;
        byte[] nv21 = new byte[ySize + uvSize * 2];

        Image.Plane yPlane = image.getPlanes()[0];
        ByteBuffer yBuffer = yPlane.getBuffer();
        int yRowStride = yPlane.getRowStride();
        int yPixelStride = yPlane.getPixelStride();
        int pos = 0;
        if (yRowStride == width && yPixelStride == 1) {
            yBuffer.get(nv21, 0, ySize);
            pos = ySize;
        } else {
            for (int row = 0; row < height; row++) {
                for (int col = 0; col < width; col++) {
                    nv21[pos++] = yBuffer.get(row * yRowStride + col * yPixelStride);
                }
            }
        }

        Image.Plane uPlane = image.getPlanes()[1];
        Image.Plane vPlane = image.getPlanes()[2];
        ByteBuffer uBuffer = uPlane.getBuffer();
        ByteBuffer vBuffer = vPlane.getBuffer();
        int chromaRowStride = uPlane.getRowStride();
        int chromaPixelStride = uPlane.getPixelStride();
        int chromaHeight = height / 2;
        int chromaWidth = width / 2;

        if (chromaPixelStride == 2) {
            for (int row = 0; row < chromaHeight; row++) {
                for (int col = 0; col < chromaWidth; col++) {
                    int offset = row * chromaRowStride + col * chromaPixelStride;
                    nv21[pos++] = vBuffer.get(offset);
                    nv21[pos++] = uBuffer.get(offset);
                }
            }
        } else {
            for (int row = 0; row < chromaHeight; row++) {
                for (int col = 0; col < chromaWidth; col++) {
                    int offset = row * chromaRowStride + col * chromaPixelStride;
                    nv21[pos++] = vBuffer.get(offset);
                    nv21[pos++] = uBuffer.get(offset);
                }
            }
        }
        return nv21;
    }

    @NonNull
    private static Bitmap rotateBitmap(@NonNull Bitmap source, int degrees) {
        Matrix matrix = new Matrix();
        matrix.postRotate(degrees);
        return Bitmap.createBitmap(source, 0, 0, source.getWidth(), source.getHeight(), matrix, true);
    }

    @NonNull
    static Bitmap cropCenterSquare(@NonNull Bitmap source) {
        int width = source.getWidth();
        int height = source.getHeight();
        int size = Math.min(width, height);
        if (width == height) {
            return source;
        }
        int x = (width - size) / 2;
        int y = (height - size) / 2;
        Bitmap cropped = Bitmap.createBitmap(source, x, y, size, size);
        source.recycle();
        return cropped;
    }
}
