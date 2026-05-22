package com.fakeeyes.fishingfloatalert;

import android.graphics.Bitmap;
import android.graphics.Color;

import androidx.annotation.NonNull;

import org.opencv.core.Core;
import org.opencv.core.CvType;
import org.opencv.core.Mat;
import org.opencv.core.Scalar;
import org.opencv.core.Size;
import org.opencv.imgproc.Imgproc;

final class ColorMaskProcessor {

    private static final int MAX_PROCESS_EDGE = 640;
    private static volatile boolean openCvLoaded;

    private ColorMaskProcessor() {
    }

    static void ensureOpenCvLoaded() {
        if (openCvLoaded) {
            return;
        }
        synchronized (ColorMaskProcessor.class) {
            if (openCvLoaded) {
                return;
            }
            try {
                System.loadLibrary("opencv_java4");
                openCvLoaded = true;
            } catch (UnsatisfiedLinkError ignored) {
                openCvLoaded = false;
            }
        }
    }

    static int countMatchingPixels(
            @NonNull Bitmap source,
            int targetColor,
            int tolerance) {
        if (ensureOpenCvReady()) {
            Mat mask = createMaskMat(source, targetColor, tolerance);
            try {
                return Core.countNonZero(mask);
            } finally {
                mask.release();
            }
        }
        return countMatchingPixelsJava(source, targetColor, tolerance);
    }

    @NonNull
    static Bitmap createMaskBitmap(
            @NonNull Bitmap source,
            int targetColor,
            int tolerance) {
        if (ensureOpenCvReady()) {
            Mat mask = createMaskMat(source, targetColor, tolerance);
            try {
                return maskMatToBitmap(mask, source.getWidth(), source.getHeight());
            } finally {
                mask.release();
            }
        }
        return createMaskBitmapJava(source, targetColor, tolerance);
    }

    private static boolean ensureOpenCvReady() {
        ensureOpenCvLoaded();
        return openCvLoaded;
    }

    @NonNull
    private static Mat createMaskMat(
            @NonNull Bitmap source,
            int targetColor,
            int tolerance) {
        Bitmap working = source;
        Bitmap scaled = null;
        int width = source.getWidth();
        int height = source.getHeight();
        int maxEdge = Math.max(width, height);
        float scaleFactor = 1f;
        if (maxEdge > MAX_PROCESS_EDGE) {
            scaleFactor = MAX_PROCESS_EDGE / (float) maxEdge;
            int newW = Math.max(1, Math.round(width * scaleFactor));
            int newH = Math.max(1, Math.round(height * scaleFactor));
            scaled = Bitmap.createScaledBitmap(source, newW, newH, true);
            working = scaled;
        }

        Mat rgb = bitmapToRgbMat(working);
        if (scaled != null) {
            scaled.recycle();
        }

        int tol = Math.max(0, tolerance);
        int r = Color.red(targetColor);
        int g = Color.green(targetColor);
        int b = Color.blue(targetColor);
        Scalar lower = new Scalar(
                Math.max(0, r - tol),
                Math.max(0, g - tol),
                Math.max(0, b - tol));
        Scalar upper = new Scalar(
                Math.min(255, r + tol),
                Math.min(255, g + tol),
                Math.min(255, b + tol));

        Mat smallMask = new Mat();
        Core.inRange(rgb, lower, upper, smallMask);
        rgb.release();

        if (scaleFactor >= 0.999f) {
            return smallMask;
        }
        Mat fullMask = new Mat();
        Imgproc.resize(smallMask, fullMask, new Size(width, height));
        smallMask.release();
        return fullMask;
    }

    @NonNull
    private static Mat bitmapToRgbMat(@NonNull Bitmap bitmap) {
        Bitmap argb = bitmap.copy(Bitmap.Config.ARGB_8888, false);
        int width = argb.getWidth();
        int height = argb.getHeight();
        int[] pixels = new int[width * height];
        argb.getPixels(pixels, 0, width, 0, 0, width, height);
        if (argb != bitmap) {
            argb.recycle();
        }

        byte[] data = new byte[width * height * 3];
        int index = 0;
        for (int pixel : pixels) {
            data[index++] = (byte) Color.red(pixel);
            data[index++] = (byte) Color.green(pixel);
            data[index++] = (byte) Color.blue(pixel);
        }

        Mat rgb = new Mat(height, width, CvType.CV_8UC3);
        rgb.put(0, 0, data);
        return rgb;
    }

    @NonNull
    private static Bitmap maskMatToBitmap(@NonNull Mat mask, int width, int height) {
        Bitmap output = Bitmap.createBitmap(width, height, Bitmap.Config.ARGB_8888);
        int[] pixels = new int[width * height];
        byte[] maskBytes = new byte[width * height];
        mask.get(0, 0, maskBytes);
        for (int i = 0; i < maskBytes.length; i++) {
            pixels[i] = (maskBytes[i] & 0xFF) != 0 ? Color.WHITE : Color.BLACK;
        }
        output.setPixels(pixels, 0, width, 0, 0, width, height);
        return output;
    }

    private static int countMatchingPixelsJava(
            @NonNull Bitmap source,
            int targetColor,
            int tolerance) {
        int tr = Color.red(targetColor);
        int tg = Color.green(targetColor);
        int tb = Color.blue(targetColor);
        int tol = Math.max(0, tolerance);
        int width = source.getWidth();
        int height = source.getHeight();
        int[] row = new int[width];
        int count = 0;
        for (int y = 0; y < height; y++) {
            source.getPixels(row, 0, width, 0, y, width, 1);
            for (int x = 0; x < width; x++) {
                int pixel = row[x];
                if (matchesColor(pixel, tr, tg, tb, tol)) {
                    count++;
                }
            }
        }
        return count;
    }

    @NonNull
    private static Bitmap createMaskBitmapJava(
            @NonNull Bitmap source,
            int targetColor,
            int tolerance) {
        int tr = Color.red(targetColor);
        int tg = Color.green(targetColor);
        int tb = Color.blue(targetColor);
        int tol = Math.max(0, tolerance);
        int width = source.getWidth();
        int height = source.getHeight();
        int[] pixels = new int[width * height];
        source.getPixels(pixels, 0, width, 0, 0, width, height);
        for (int i = 0; i < pixels.length; i++) {
            pixels[i] = matchesColor(pixels[i], tr, tg, tb, tol) ? Color.WHITE : Color.BLACK;
        }
        Bitmap output = Bitmap.createBitmap(width, height, Bitmap.Config.ARGB_8888);
        output.setPixels(pixels, 0, width, 0, 0, width, height);
        return output;
    }

    private static boolean matchesColor(int pixel, int tr, int tg, int tb, int tol) {
        return Math.abs(Color.red(pixel) - tr) <= tol
                && Math.abs(Color.green(pixel) - tg) <= tol
                && Math.abs(Color.blue(pixel) - tb) <= tol;
    }

    static int getColorAt(@NonNull Bitmap bitmap, int x, int y) {
        int clampedX = Math.max(0, Math.min(bitmap.getWidth() - 1, x));
        int clampedY = Math.max(0, Math.min(bitmap.getHeight() - 1, y));
        return bitmap.getPixel(clampedX, clampedY);
    }
}
