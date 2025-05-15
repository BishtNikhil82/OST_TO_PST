package com.NAtools.htmlparsing;

import org.jsoup.nodes.Element;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import org.jsoup.Jsoup;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

public class HtmlToRtfConverter {
    private static final Logger logger = Logger.getLogger(HtmlToRtfConverter.class.getName());

    // Inner class to represent text parts
    public static class TextPart {
        String text;
        boolean isTable;
        boolean isRow;
        boolean isCell;
        boolean isBold;
        boolean isItalic;
        boolean isUnderline;
        int cellWidth;
        boolean isImage;
        int imageWidth;
        int imageHeight;
        String imageData;

        TextPart(String text, boolean isTable, boolean isRow, boolean isCell, boolean isBold, boolean isItalic, boolean isUnderline, int cellWidth, boolean isImage, int imageWidth, int imageHeight, String imageData) {
            this.text = text;
            this.isTable = isTable;
            this.isRow = isRow;
            this.isCell = isCell;
            this.isBold = isBold;
            this.isItalic = isItalic;
            this.isUnderline = isUnderline;
            this.cellWidth = cellWidth;
            this.isImage = isImage;
            this.imageWidth = imageWidth;
            this.imageHeight = imageHeight;
            this.imageData = imageData;
        }
    }

    public static void extractTextPartsFromElement(Element element, List<TextPart> textParts) {
        boolean isBold = element.tagName().equals("b") || element.tagName().equals("strong");
        boolean isItalic = element.tagName().equals("i") || element.tagName().equals("em");
        boolean isUnderline = element.tagName().equals("u");

        if (element.tagName().equals("table")) {
            textParts.add(new TextPart(null, true, false, false, false, false, false, 0, false, 0, 0, null));  // Start of table
        } else if (element.tagName().equals("tr")) {
            textParts.add(new TextPart(null, false, true, false, false, false, false, 0, false, 0, 0, null));  // Start of row
        } else if (element.tagName().equals("td") || element.tagName().equals("th")) {
            int cellWidth = 5000;  // Default width in twips (RTF units)
            if (element.hasAttr("width")) {
                try {
                    cellWidth = Integer.parseInt(element.attr("width")) * 20;  // Convert to RTF units
                } catch (NumberFormatException e) {
                    logger.warning("Invalid width value: " + element.attr("width"));
                }
            } else if (element.hasAttr("style")) {
                String style = element.attr("style");
                if (style.contains("width:")) {
                    try {
                        String widthValue = style.split("width:")[1].split(";")[0].replaceAll("[^0-9]", "");
                        cellWidth = Integer.parseInt(widthValue) * 20;  // Convert to RTF units
                    } catch (NumberFormatException | ArrayIndexOutOfBoundsException e) {
                        logger.warning("Invalid style width value: " + style);
                    }
                }
            }

            // Ensure the cell width does not exceed the page width
            int maxCellWidth = 8000;  // Reduced to fit within the page
            cellWidth = Math.min(cellWidth, maxCellWidth);

            textParts.add(new TextPart(element.text(), false, false, true, isBold, isItalic, isUnderline, cellWidth, false, 0, 0, null));  // Add cell content
        } else if (element.tagName().equals("img")) {
            try {
                String fakeImageHexData = createFakeImage(300, 200);  // Create a small fake image as placeholder
                textParts.add(new TextPart(fakeImageHexData, false, false, false, false, false, false, 0, true, 300, 200, fakeImageHexData));  // Embed image
            } catch (Exception e) {
                logger.warning("Failed to create fake image: " + e.getMessage());
            }
        } else if (element.hasText() && !element.tagName().equals("script") && !element.tagName().equals("style")) {
            textParts.add(new TextPart(element.ownText(), false, false, false, isBold, isItalic, isUnderline, 0, false, 0, 0, null));
        }

        for (Element child : element.children()) {
            extractTextPartsFromElement(child, textParts);  // Recursively process child elements
        }
    }

    public static List<TextPart> extractTextAndLinksFromHtml(String htmlContent) {
        List<TextPart> textParts = new ArrayList<>();
        Document document = Jsoup.parse(htmlContent);
        Elements bodyElements = document.body().children();  // Get only direct children of the body

        for (Element element : bodyElements) {
            extractTextPartsFromElement(element, textParts);
        }

        return textParts;
    }

    // Convert text parts to RTF format
    public static String convertToRtf(List<TextPart> textParts) {
        StringBuilder rtfContent = new StringBuilder("{\\rtf1\\ansi\\deff0");

        boolean inTable = false;
        boolean inRow = false;
        int currentCellIndex = 0;

        for (TextPart part : textParts) {
            if (part.isTable) {
                if (inTable) {
                    if (inRow) {
                        rtfContent.append("\\row "); // Close the last row
                        inRow = false;
                    }
                    rtfContent.append("\\pard \\par "); // Close the table
                }
                rtfContent.append("\\trowd\\trgaph108\\trleft-108 ");
                inTable = true;
                inRow = false;
                currentCellIndex = 0;
            } else if (part.isRow) {
                if (inRow) {
                    rtfContent.append("\\row "); // Close the previous row
                }
                rtfContent.append("\\trowd\\trgaph108\\trleft-108 ");
                inRow = true;
                currentCellIndex = 0;
            } else if (part.isCell) {
                currentCellIndex += part.cellWidth;  // Increment current cell width
                rtfContent.append("\\cellx").append(currentCellIndex);
                if (part.isBold) rtfContent.append("\\b ");
                if (part.isItalic) rtfContent.append("\\i ");
                if (part.isUnderline) rtfContent.append("\\ul ");

                rtfContent.append(part.text);

                if (part.isUnderline) rtfContent.append("\\ulnone ");
                if (part.isItalic) rtfContent.append("\\i0 ");
                if (part.isBold) rtfContent.append("\\b0 ");

                rtfContent.append("\\cell ");
            } else if (part.isImage) {
                rtfContent.append("\\pard\\qc ");
                rtfContent.append("{\\pict\\wmetafile8\\picwgoal").append(part.imageWidth).append("\\pichgoal").append(part.imageHeight).append(" ");
                rtfContent.append(part.imageData);  // Embed the fake image data
                rtfContent.append("}\\par ");
            } else {
                if (inRow) {
                    rtfContent.append("\\row "); // Close the row before starting non-table content
                    inRow = false;
                }
                if (part.isBold) rtfContent.append("\\b ");
                if (part.isItalic) rtfContent.append("\\i ");
                if (part.isUnderline) rtfContent.append("\\ul ");

                rtfContent.append(part.text);

                if (part.isUnderline) rtfContent.append("\\ulnone ");
                if (part.isItalic) rtfContent.append("\\i0 ");
                if (part.isBold) rtfContent.append("\\b0 ");

                rtfContent.append("\\par "); // Paragraph break
            }
        }

        if (inRow) {
            rtfContent.append("\\row "); // Close the last row if still open
        }
        if (inTable) {
            rtfContent.append("\\pard \\par "); // Close the table
        }

        rtfContent.append("}");

        return rtfContent.toString();
    }

    // New method to extract HTML-like content from RTF and process it
    public static String convertRtfWithHtmlToRtf(String rtfContent) {
        // Step 1: Extract HTML-like content from the RTF
        String htmlContent = extractHtmlFromRtf(rtfContent);

        // Step 2: Parse the HTML content into text parts
        List<TextPart> textParts = extractTextAndLinksFromHtml(htmlContent);

        // Step 3: Convert the parsed text parts into RTF format
        return convertToRtf(textParts);
    }

    // Helper method to extract HTML-like content from RTF
    public static String extractHtmlFromRtf(String rtfContent) {
        // Simple regex to extract content between tags like < and >
        return rtfContent.replaceAll("\\{\\\\[^}]+\\}", "").replaceAll("\\\\[^ ]+", "").replaceAll("[{}]", "").trim();
    }

    // Helper method to create a fake image and return its hex data
    public static String createFakeImage(int width, int height) throws IOException {
        BufferedImage fakeImage = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2d = fakeImage.createGraphics();
        g2d.setColor(Color.LIGHT_GRAY);
        g2d.fillRect(0, 0, width, height);
        g2d.setColor(Color.BLACK);
        g2d.drawString("Fake Image", width / 4, height / 2);
        g2d.dispose();

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(fakeImage, "png", baos); // Save as PNG format
        byte[] bytes = baos.toByteArray();

        StringBuilder hexString = new StringBuilder();
        for (byte b : bytes) {
            hexString.append(String.format("%02X", b));
        }

        return hexString.toString();
    }
}
