package com.NAtools.util;

import java.util.Base64;

public class ImageUtil {
    public static String toBase64String(byte[] imageData) {
        return Base64.getEncoder().encodeToString(imageData);
    }
}