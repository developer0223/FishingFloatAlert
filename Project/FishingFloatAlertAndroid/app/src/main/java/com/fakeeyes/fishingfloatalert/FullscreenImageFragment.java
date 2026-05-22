package com.fakeeyes.fishingfloatalert;

import android.graphics.Bitmap;
import android.os.Bundle;
import android.view.View;
import android.widget.ImageView;

import androidx.annotation.DrawableRes;
import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.fragment.app.Fragment;

abstract class FullscreenImageFragment extends Fragment {

    @Nullable
    private ImageView fullscreenImage;

    @Override
    public void onViewCreated(@NonNull View view, @Nullable Bundle savedInstanceState) {
        super.onViewCreated(view, savedInstanceState);
        fullscreenImage = view.findViewById(R.id.fullscreen_image);
        int defaultImageResId = getDefaultImageResId();
        if (fullscreenImage != null && defaultImageResId != 0) {
            fullscreenImage.setImageResource(defaultImageResId);
        }
    }

    @DrawableRes
    protected abstract int getDefaultImageResId();

    public void setFullscreenImage(@DrawableRes int drawableResId) {
        if (fullscreenImage != null) {
            fullscreenImage.setImageResource(drawableResId);
        }
    }

    public void setFullscreenImage(@Nullable Bitmap bitmap) {
        if (fullscreenImage != null) {
            fullscreenImage.setImageBitmap(bitmap);
        }
    }

    @Nullable
    public ImageView getFullscreenImageView() {
        return fullscreenImage;
    }
}
