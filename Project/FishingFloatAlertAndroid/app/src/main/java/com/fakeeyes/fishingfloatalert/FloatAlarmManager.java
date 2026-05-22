package com.fakeeyes.fishingfloatalert;

import android.media.AudioManager;
import android.media.ToneGenerator;
import android.os.Handler;
import android.os.HandlerThread;

import androidx.annotation.NonNull;

public final class FloatAlarmManager {

    private static final int BEEP_FREQUENCY_HZ = 1320;
    private static final int BEEP_DURATION_MS = 150;
    private static final long BEEP_INTERVAL_MS = 220L;

    private static volatile FloatAlarmManager instance;

    private final HandlerThread workerThread = new HandlerThread("FloatAlarm");
    private Handler workerHandler;
    private volatile boolean shutdown;
    private volatile boolean sustainRequested;

    private FloatAlarmManager() {
        workerThread.start();
        workerHandler = new Handler(workerThread.getLooper());
        workerHandler.post(this::beepLoop);
    }

    @NonNull
    public static FloatAlarmManager getInstance() {
        if (instance == null) {
            synchronized (FloatAlarmManager.class) {
                if (instance == null) {
                    instance = new FloatAlarmManager();
                }
            }
        }
        return instance;
    }

    public void updateFromMatchCount(int matchCount) {
        AppSettings settings = AppSettings.getInstance();
        boolean shouldSustain = settings.isAlarmEnabled()
                && matchCount <= settings.getAlarmMatchMax();
        sustainRequested = shouldSustain;
    }

    public void stop() {
        sustainRequested = false;
    }

    public void shutdown() {
        shutdown = true;
        sustainRequested = false;
        workerThread.quitSafely();
    }

    private void beepLoop() {
        if (shutdown) {
            return;
        }
        if (sustainRequested) {
            playBeep();
            workerHandler.postDelayed(this::beepLoop, BEEP_INTERVAL_MS);
        } else {
            workerHandler.postDelayed(this::beepLoop, BEEP_INTERVAL_MS);
        }
    }

    private void playBeep() {
        try {
            ToneGenerator toneGenerator = new ToneGenerator(
                    AudioManager.STREAM_MUSIC,
                    (int) (ToneGenerator.MAX_VOLUME * 0.85f));
            toneGenerator.startTone(ToneGenerator.TONE_PROP_BEEP, BEEP_DURATION_MS);
        } catch (RuntimeException ignored) {
        }
    }
}
