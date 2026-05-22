package com.fakeeyes.fishingfloatalert;

import android.content.Context;
import android.graphics.Canvas;
import android.graphics.Paint;
import android.util.AttributeSet;
import android.view.View;

import androidx.annotation.Nullable;

public class CrosshairOverlayView extends View {

    private static final float CROSSHAIR_ARM_DP = 28f;

    private final Paint paint = new Paint(Paint.ANTI_ALIAS_FLAG);
    private float centerX = -1f;
    private float centerY = -1f;
    private float armLengthPx;

    public CrosshairOverlayView(Context context) {
        super(context);
        init();
    }

    public CrosshairOverlayView(Context context, @Nullable AttributeSet attrs) {
        super(context, attrs);
        init();
    }

    public CrosshairOverlayView(Context context, @Nullable AttributeSet attrs, int defStyleAttr) {
        super(context, attrs, defStyleAttr);
        init();
    }

    private void init() {
        paint.setColor(0xFFFFFFFF);
        paint.setStrokeWidth(3f);
        paint.setStyle(Paint.Style.STROKE);
        setWillNotDraw(false);
        armLengthPx = CROSSHAIR_ARM_DP * getResources().getDisplayMetrics().density;
    }

    public void setCrosshairPosition(float x, float y) {
        centerX = x;
        centerY = y;
        invalidate();
    }

    public void hideCrosshair() {
        centerX = -1f;
        centerY = -1f;
        invalidate();
    }

    public boolean isCrosshairVisible() {
        return centerX >= 0f && centerY >= 0f;
    }

    @Override
    protected void onDraw(Canvas canvas) {
        super.onDraw(canvas);
        if (!isCrosshairVisible()) {
            return;
        }
        canvas.drawLine(centerX - armLengthPx, centerY, centerX + armLengthPx, centerY, paint);
        canvas.drawLine(centerX, centerY - armLengthPx, centerX, centerY + armLengthPx, paint);
    }
}
