package com.NAtools.htmlparsing;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class RTFAdjuster {

    // Maximum width of the page in twips
    private static final int MAX_PAGE_WIDTH_TWIPS = 12240;

    public static String adjustRTFContent(String rtfContent) {
        // Adjust cellx values
        rtfContent = adjustCellxValues(rtfContent);

        // Adjust image dimensions
        rtfContent = adjustImageDimensions(rtfContent);

        return rtfContent;
    }

    private static String adjustCellxValues(String rtfContent) {
        Pattern cellxPattern = Pattern.compile("\\\\cellx(\\d+)");
        Matcher matcher = cellxPattern.matcher(rtfContent);
        StringBuffer adjustedContent = new StringBuffer();

        while (matcher.find()) {
            int cellxValue = Integer.parseInt(matcher.group(1));

            // Ensure the cellx value does not exceed the page width
            if (cellxValue > MAX_PAGE_WIDTH_TWIPS) {
                cellxValue = MAX_PAGE_WIDTH_TWIPS;
            }

            matcher.appendReplacement(adjustedContent, "\\\\cellx" + cellxValue);
        }
        matcher.appendTail(adjustedContent);

        return adjustedContent.toString();
    }

    private static String adjustImageDimensions(String rtfContent) {
        Pattern pictPattern = Pattern.compile("\\\\pict\\wmetafile8\\\\picwgoal(\\d+)\\\\pichgoal(\\d+)");
        Matcher matcher = pictPattern.matcher(rtfContent);
        StringBuffer adjustedContent = new StringBuffer();

        while (matcher.find()) {
            int picwgoal = Integer.parseInt(matcher.group(1));
            int pichgoal = Integer.parseInt(matcher.group(2));

            // Adjust image width if it exceeds page width
            if (picwgoal > MAX_PAGE_WIDTH_TWIPS) {
                double aspectRatio = (double) pichgoal / picwgoal;
                picwgoal = MAX_PAGE_WIDTH_TWIPS;
                pichgoal = (int) (picwgoal * aspectRatio);
            }

            matcher.appendReplacement(adjustedContent, "\\\\pict\\wmetafile8\\\\picwgoal" + picwgoal + "\\\\pichgoal" + pichgoal);
        }
        matcher.appendTail(adjustedContent);

        return adjustedContent.toString();
    }

    public static void main(String[] args) {
        String rtfContent = "Your RTF content here";
        String adjustedRTF = adjustRTFContent(rtfContent);
        System.out.println(adjustedRTF);
    }
}

