package com.NAtools.htmlparsing;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.util.ArrayList;
import java.util.List;

public class HtmlTextAndLinkExtractor {

    public static class TextPart {
        public String text;
        public String link;
        public boolean isListItem;
        public boolean isBold;
        public boolean isItalic;
        public boolean isUnderline;

        TextPart(String text, String link, boolean isListItem, boolean isBold, boolean isItalic, boolean isUnderline) {
            this.text = text;
            this.link = link;
            this.isListItem = isListItem;
            this.isBold = isBold;
            this.isItalic = isItalic;
            this.isUnderline = isUnderline;
        }
    }
    private static void extractTextPartsFromElement(Element element, List<TextPart> textParts) {
        boolean isListItem = element.tagName().equals("li");
        boolean isBold = element.tagName().equals("b") || element.tagName().equals("strong");
        boolean isItalic = element.tagName().equals("i") || element.tagName().equals("em");
        boolean isUnderline = element.tagName().equals("u");

        if (element.tagName().equals("a")) {
            String linkText = element.text();
            String linkHref = element.attr("href");
            textParts.add(new TextPart(linkText, linkHref, isListItem, isBold, isItalic, isUnderline));
        } else if (element.tagName().equals("img")) {
            String imageSrc = element.attr("src");
            // Skip embedded images (data URIs or CID references)
            if (!imageSrc.startsWith("data:") && !imageSrc.startsWith("cid:")) {
               // textParts.add(new TextPart("[Image: " + imageSrc + "]", null, isListItem, isBold, isItalic, isUnderline));
            }
        } else if (element.hasText() && !element.tagName().equals("script") && !element.tagName().equals("style")) {
            textParts.add(new TextPart(element.ownText(), null, isListItem, isBold, isItalic, isUnderline));  // Use ownText() to avoid duplication
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

    void parseElementToRtf(Element element, StringBuilder rtfContent) {
        String tagName = element.tagName();

        switch (tagName) {
            case "p":
                rtfContent.append("\\par ");
                rtfContent.append(element.text());
                rtfContent.append("\\par ");
                break;
            case "b":
                rtfContent.append("\\b ");
                rtfContent.append(element.text());
                rtfContent.append("\\b0 ");
                break;
            case "i":
                rtfContent.append("\\i ");
                rtfContent.append(element.text());
                rtfContent.append("\\i0 ");
                break;
            case "u":
                rtfContent.append("\\ul ");
                rtfContent.append(element.text());
                rtfContent.append("\\ulnone ");
                break;
            case "a":
                String linkHref = element.attr("href");
                rtfContent.append("{\\field{\\*\\fldinst HYPERLINK \"");
                rtfContent.append(linkHref);
                rtfContent.append("\"}{\\fldrslt ");
                rtfContent.append(element.text());
                rtfContent.append("}}");
                break;
            // Add more cases as needed
            default:
                rtfContent.append(element.text());
                break;
        }

        for (Element child : element.children()) {
            parseElementToRtf(child, rtfContent);
        }
    }



}