package com.NAtools.Convertor;
import com.NAtools.config.LogManagerConfig;
import com.NAtools.htmlparsing.HtmlTextAndLinkExtractor;
import com.NAtools.htmlparsing.HtmlToRtfConverter;
import com.NAtools.htmlparsing.RTFAdjuster;
import com.aspose.email.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import javax.imageio.ImageIO;
import javax.swing.text.rtf.RTFEditorKit;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.List;
import java.util.logging.Logger;
import javax.swing.text.BadLocationException;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.safety.Safelist;
import org.jsoup.select.Elements;

import static com.NAtools.htmlparsing.RTFAdjuster.adjustRTFContent;

public class OSTToRTF {
    private static final Logger logger = Logger.getLogger(OSTToRTF.class.getName());
    static {
        LogManagerConfig.configureLogger(logger);
    }
    private final PersonalStorage sourcePst;
    private final String outputDirectory;
    private final boolean enableDuplicateCheck;
    private final Set<String> setDuplicacy = new HashSet<>();
    private final Set<String> setDupliccal = new HashSet<>();
    private final Set<String> setDuplictask = new HashSet<>();
    private final Set<String> setDupliccontact = new HashSet<>();
    private final String ostFileName;

    public OSTToRTF(PersonalStorage sourcePst, String outputDirectory, String sourcePstPath, boolean enableDuplicateCheck) {
        this.sourcePst = sourcePst;
        this.outputDirectory = outputDirectory;
        this.enableDuplicateCheck = enableDuplicateCheck;

        // Extract the OST file name without the extension
        this.ostFileName = new File(sourcePstPath).getName().replaceFirst("[.][^.]+$", "");
    }

    public void convertToRTF() {
        try {
            FolderInfo rootFolder = sourcePst.getRootFolder();
            processFolder(rootFolder, outputDirectory);
            logger.info("Conversion to RTF format completed.");
        } catch (Exception e) {
            logger.severe("Error during RTF conversion: " + e.getMessage());
        } finally {
            sourcePst.dispose();
        }
    }

    private void processFolder(FolderInfo folder, String parentPath) {
        try {
            String folderName = cleanFolderName(folder.getDisplayName());
            if (folderName.isEmpty() || folderName.equals("Unnamed")) {
                folderName = ostFileName;
                logger.info("Unnamed folder found. Generated name: " + folderName);
            }

            String fullPath = parentPath.isEmpty() ? folderName : parentPath + "/" + folderName;
            File directory = new File(fullPath);
            if (!directory.exists()) {
                directory.mkdirs();  // Create directories if they don't exist
            }

            FolderInfoCollection subFolders = folder.getSubFolders();
            for (FolderInfo subFolder : subFolders) {
                processFolder(subFolder, fullPath);
            }

            processMessages(folder, fullPath);

        } catch (Exception e) {
            logger.severe("Error processing folder: " + e.getMessage());
        }
    }

    private void processMessages(FolderInfo folder, String folderPath) {
        try {
            MessageInfoCollection messages = folder.getContents();

            for (int i = 0; i < messages.size(); i++) {
                try {
                    MessageInfo messageInfo = messages.get_Item(i);
                    // Create MailConversionOptions object
                    MapiConversionOptions mapiOptions = MapiConversionOptions.getUnicodeFormat();
                    MailConversionOptions options = new MailConversionOptions();

                    // Extract the message as a MailMessage
                    MapiMessage message1 = sourcePst.extractMessage(messageInfo);
                    MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
                    d.setPreserveEmbeddedMessageFormat(true);
                    MailConversionOptions de = new MailConversionOptions();
                    MailMessage mess1 = message1.toMailMessage(de);
                    MapiMessage message = MapiMessage.fromMailMessage(mess1, d);
                    switch (message.getMessageClass()) {
                        case "IPM.Note":
                            //saveMessage(message, folderPath);
                            saveMessage(message,folderPath);
                            break;
                        case "IPM.Contact":
                            saveContact(message, folderPath);
                            break;
                        case "IPM.Appointment":
                        case "IPM.Schedule.Meeting.Request":
                            saveCalendarAsICS(message, folderPath);
                            break;
                        case "IPM.Task":
                            saveTask(message, folderPath);
                            break;
                        default:
                            saveGenericMessage( message, folderPath);
                            break;
                    }

                } catch (Exception e) {
                    logger.warning("Error processing message: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            logger.severe("Error processing messages: " + e.getMessage());
        }
    }

    public String resolveSender(MapiMessage message) {
        String senderAddress = message.getSenderEmailAddress();
        String senderDisplayName = message.getSenderName(); // This gets the display name if available

        // Check if the sender address is a DN (distinguished name)
        if (senderAddress != null && senderAddress.startsWith("/O=FIRST ORGANIZATION/OU=EXCHANGE ADMINISTRATIVE GROUP")) {
            // If senderDisplayName is null or empty, use a fallback like "Unknown Sender"
            if (senderDisplayName == null || senderDisplayName.isEmpty()) {
                senderDisplayName = "Unknown Sender";
            }
            // Replace DN with the display name and actual email address from MapiMessage
            return senderDisplayName ;
        } else {
            // If it's not a DN, return the standard display name and email format
            if (senderDisplayName != null && !senderDisplayName.isEmpty()) {
                return senderDisplayName ;
            } else {
                return senderAddress; // If no display name, just return the email
            }
        }
    }

    private String convertPlainTextToHtml(String plainText) {
        if (plainText == null || plainText.isEmpty()) {
            return "";
        }

        // Convert URLs into clickable links
        String htmlContent = plainText.replaceAll("(?i)(https?://\\S+)", "<a href=\"$1\">$1</a>");

        // Wrap the plain text in basic HTML tags
        htmlContent = "<html><body style=\"font-family: Arial, sans-serif;\">" + htmlContent + "</body></html>";

        return htmlContent;
    }
    private String extractRecipients(MapiRecipientCollection recipients) {
        StringBuilder recipientList = new StringBuilder();
        for (MapiRecipient recipient : recipients) {
            if (recipientList.length() > 0) {
                recipientList.append(", ");
            }
            recipientList.append(recipient.getEmailAddress());
        }
        return recipientList.toString();
    }
    private String cleanHtmlComments(String htmlContent) {
        return htmlContent.replaceAll("<!--.*?-->", ""); // Basic example, can be adjusted for more complex cases
    }
    private Map<String, String> saveAndMapCidAttachments(MapiMessage message, File attachmentDir) {
        File directory = new File(attachmentDir.getAbsolutePath() + File.separator + cleanFileName(message.getSubject()).trim());
        if (directory.exists()) {
            directory = new File(attachmentDir.getAbsolutePath() + File.separator + cleanFileName(message.getSubject()).trim() + "_" + System.currentTimeMillis());
        }
        directory.mkdir();
        int idx =0 ;
        Map<String, String> cidToImageMap = new HashMap<>();
        for (MapiAttachment attachment : message.getAttachments()) {
            try {
                String cid = attachment.getDisplayName().trim();
                String cidNameWithoutExtension = cid.substring(0, cid.lastIndexOf('.'));
                String filePath = directory.getAbsolutePath() + File.separator + cleanFileName(attachment.getDisplayName()).trim();
                attachment.save(filePath);
                cidToImageMap.put(cidNameWithoutExtension, filePath);
                // ++idx;
            } catch (Exception e) {
                attachment.save(directory.getAbsolutePath() + File.separator + cleanFileName(attachment.getLongFileName()).trim());
            }
        }
        return cidToImageMap;
    }
    private String convertPlainTextToRtf(String plainText) {
        if (plainText == null || plainText.isEmpty()) {
            return "";
        }

        // Convert URLs into clickable links in RTF
        String rtfContent = plainText.replaceAll("(?i)(https?://\\S+)", "{\\\\field{\\\\*\\\\fldinst HYPERLINK \"$1\"}{\\\\fldrslt $1}}");

        // Wrap the plain text in basic RTF formatting
        rtfContent = "{\\rtf1\\ansi\\deff0 " + rtfContent.replace("\n", "\\par ") + "}";

        return rtfContent;
    }

    private String convertCssToRtf(String css) {
        // Parse CSS and map to RTF equivalents
        StringBuilder rtfCss = new StringBuilder();

        if (css.contains("font-weight:bold")) {
            rtfCss.append("\\b ");
        }
        if (css.contains("font-style:italic")) {
            rtfCss.append("\\i ");
        }
        // Add more mappings as needed...

        return rtfCss.toString();
    }
    private String getRtfColorCode(String color) {
        if (color.startsWith("#")) {
            // Convert the color hex code to R, G, B values
            int r = Integer.parseInt(color.substring(1, 3), 16);
            int g = Integer.parseInt(color.substring(3, 5), 16);
            int b = Integer.parseInt(color.substring(5, 7), 16);

            // Return the RTF color definition
            return "\\red" + r + "\\green" + g + "\\blue" + b + ";";
        }
        // You can add more cases to handle named colors if needed
        // For example: if (color.equalsIgnoreCase("red")) return "\\red255\\green0\\blue0;";

        return "";
    }

    public static List<String> extractTextPartsFromHtml(String htmlContent) {
        List<String> textParts = new ArrayList<>();

        Document document = Jsoup.parse(htmlContent);
        Elements elements = document.body().getAllElements();

        for (Element element : elements) {
            if (element.tagName().equals("p") || element.tagName().equals("br") || element.tagName().equals("a")) {
                textParts.add(element.text());
            }
        }

        return textParts;
    }
    public static String createRtfFromTextParts(List<HtmlTextAndLinkExtractor.TextPart> textParts) {
        StringBuilder rtfContent = new StringBuilder();
        rtfContent.append("{\\rtf1\\ansi\\deff0");
        rtfContent.append("{\\fonttbl{\\f0 Arial;}{\\f1 Segoe UI;}}");
        rtfContent.append("{\\colortbl ;\\red0\\green0\\blue0;\\red255\\green0\\blue0;\\red0\\green0\\blue255;}");

        for (HtmlTextAndLinkExtractor.TextPart part : textParts) {
            if (part.link != null) {
                // Add hyperlink with associated text
                rtfContent.append("{\\field{\\*\\fldinst HYPERLINK \"")
                        .append(part.link)
                        .append("\"}{\\fldrslt \\ul\\cf3 ")
                        .append(part.text)
                        .append("\\ulnone\\cf0}}");
            } else {
                // Add plain text
                rtfContent.append(part.text);
            }

            // Append a paragraph break after each part
            rtfContent.append("\\par ");
        }

        rtfContent.append("}");
        return rtfContent.toString();
    }


    public void saveMessage(MapiMessage message, String outputDirectory) {
        Logger logger = Logger.getLogger("RTFLogger");
        logger.info("Starting to process the message: " + message.getSubject());
        try {
            String subject = message.getSubject() != null ? message.getSubject() : "Unnamed";
            String fileName = subject.replaceAll("[^a-zA-Z0-9\\.\\-]", "_") + ".rtf";
            if (fileName.contains("Welcome")){
                logger.info("check *******************");
            }
            // Extract the body content
            String bodyContent = message.getBodyHtml();
            bodyContent = cleanHtmlComments(bodyContent);
            logger.info("Body Content Length: " + bodyContent.length());
            logger.info("Body Content Preview: " + (bodyContent.length() > 50 ? bodyContent.substring(0, 50) : bodyContent));

//            HtmlTextAndLinkExtractor extractor = new HtmlTextAndLinkExtractor();
//
//            List<HtmlTextAndLinkExtractor.TextPart> textParts = extractor.extractTextAndLinksFromHtml(bodyContent);
//
//            String rtfContent = createRtfFromTextParts(textParts);

            //String rtfContent = convertHtmlToRtf(bodyContent);


            HtmlToRtfConverter cvt = new HtmlToRtfConverter();
          //  List<HtmlToRtfConverter.TextPart> textParts = cvt.extractTextAndLinksFromHtml(bodyContent);
            String rtfContent = cvt.convertRtfWithHtmlToRtf(bodyContent);



            // Save RTF content to a file
            //rtfContent = cleanHtmlContent(rtfContent);
//            String fakeImageHexData = createFakeImage(600, 450);
//
//             rtfContent = "{\\rtf1\\ansi\\deff0\\par \n" +
//                    "\\trowd\\trgaph108\\trleft-108 \\cellx5000\\cell Hi Welcome to your Outlook!\\cell \\row \n" +
//                    "\\trowd\\trgaph108\\trleft-108 \\cellx5000 Get the free Outlook mobile app\\cell \\row \n" +
//                    "\\trowd\\trgaph108\\trleft-108 \\cellx5000 Use Microsoft 365 for free\\cell \\row \n" +
//                    "\\trowd\\trgaph108\\trleft-108 \\cellx5000\\pard\\qc {\\pict\\pngblip\\picwgoal300\\pichgoal200 " + fakeImageHexData + "}\\cell \\row \n" +  // Insert image
//                    "}";


            File outputFile = new File(outputDirectory, fileName);
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                fos.write(rtfContent.getBytes(StandardCharsets.ISO_8859_1)); // Use ISO-8859-1 for RTF
                logger.info("Saved email as RTF: " + fileName);
            } catch (IOException e) {
                logger.warning("Error saving RTF file: " + e.getMessage());
            }


        } catch (Exception e) {
            logger.warning("Error processing message for RTF conversion: " + e.getMessage());
        }
    }

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

    public String convertHtmlToRtf(String htmlContent) {
        // Remove style and script tags
        htmlContent = htmlContent.replaceAll("(?s)<style.*?</style>", "");
        htmlContent = htmlContent.replaceAll("(?s)<script.*?</script>", "");

        // Initialize the RTF content
        StringBuilder rtfContent = new StringBuilder();
        rtfContent.append("{\\rtf1\\ansi\\deff0");
        rtfContent.append("{\\fonttbl{\\f0 Arial;}{\\f1 Segoe UI;}}");
        rtfContent.append("{\\colortbl ;\\red0\\green0\\blue0;\\red255\\green0\\blue0;\\red0\\green0\\blue255;}");

        // Replace non-breaking spaces and other unnecessary characters
        htmlContent = htmlContent.replaceAll("&nbsp;", " ");
        htmlContent = htmlContent.replaceAll("[\\u200B\\u200C\\u200D\\u2060]", "");  // Remove non-printable characters

        // Replace line breaks and paragraphs properly
        htmlContent = htmlContent.replaceAll("<br\\s*/?>", "\\\\par ");
        htmlContent = htmlContent.replaceAll("</p>", "\\\\par ");
        htmlContent = htmlContent.replaceAll("<p>", "\\\\par ");

        // Handle hyperlinks
        htmlContent = htmlContent.replaceAll(
                "<a\\s+href\\s*=\\s*\"([^\"]*)\">(.*?)</a>",
                "{\\\\field{\\\\*\\\\fldinst HYPERLINK \"$1\"}{\\\\fldrslt \\\\ul\\\\cf3 $2\\\\ulnone\\\\cf0}}"
        );

        // Handle basic text formatting (bold, italic, underline)
        htmlContent = htmlContent.replaceAll("<b>(.*?)</b>", "\\\\b $1\\\\b0 ");
        htmlContent = htmlContent.replaceAll("<strong>(.*?)</strong>", "\\\\b $1\\\\b0 ");
        htmlContent = htmlContent.replaceAll("<i>(.*?)</i>", "\\\\i $1\\\\i0 ");
        htmlContent = htmlContent.replaceAll("<em>(.*?)</em>", "\\\\i $1\\\\i0 ");
        htmlContent = htmlContent.replaceAll("<u>(.*?)</u>", "\\\\ul $1\\\\ulnone ");

        // Remove remaining HTML tags
        htmlContent = htmlContent.replaceAll("<[^>]+>", "");

        // Clean up multiple spaces and line breaks
        htmlContent = htmlContent.replaceAll("\\s{2,}", " ");
        htmlContent = htmlContent.replaceAll("(\\\\par\\s*)+", "\\\\par ");
        htmlContent = htmlContent.replaceAll("(\\\\~\\s*)+", " ");
        htmlContent = htmlContent.replaceAll("\\\\par\\s*\\\\par", "\\\\par");  // Remove consecutive \par

        // Append the processed content to the RTF document
        rtfContent.append(htmlContent.trim());
        rtfContent.append("}");

        return rtfContent.toString();
    }







    private void parseElement(Element element, StringBuilder rtf, MapiMessage message) {
        for (Node node : element.childNodes()) {
            if (node instanceof Element) {
                Element childElement = (Element) node;
                switch (childElement.tagName()) {
                    case "b":
                    case "strong":
                        rtf.append("\\b ");
                        parseElement(childElement, rtf, message);
                        rtf.append("\\b0 ");
                        break;
                    case "i":
                    case "em":
                        rtf.append("\\i ");
                        parseElement(childElement, rtf, message);
                        rtf.append("\\i0 ");
                        break;
                    case "u":
                        rtf.append("\\ul ");
                        parseElement(childElement, rtf, message);
                        rtf.append("\\ulnone ");
                        break;
                    case "a":
                        rtf.append("\\ul \\cf3 ");
                        rtf.append(childElement.text());
                        rtf.append("\\ulnone \\cf0 ");
                        break;
                    case "span":
                        if (childElement.attr("style").contains("color:red")) {
                            rtf.append("\\cf2 ");
                        }
                        parseElement(childElement, rtf, message);
                        rtf.append("\\cf0 ");
                        break;
                    case "p":
                        parseElement(childElement, rtf, message);
                        rtf.append("\\par ");
                        break;
                    case "br":
                        rtf.append("\\line ");
                        break;
                    case "img":
                        String altText = childElement.attr("alt").isEmpty() ? "[Image]" : childElement.attr("alt");
                        rtf.append(altText);
                        break;
                    case "table":
                        rtf.append("\\trowd\\trautofit1 ");
                        parseElement(childElement, rtf, message);
                        rtf.append("\\row ");
                        break;
                    case "tr":
                        parseElement(childElement, rtf, message);
                        rtf.append("\\row ");
                        break;
                    case "td":
                        rtf.append("\\intbl\\cellx5000 ");
                        parseElement(childElement, rtf, message);
                        rtf.append("\\cell ");
                        break;
                    default:
                        parseElement(childElement, rtf, message);
                        break;
                }
            } else {
                rtf.append(node.toString().replace("\n", "\\par ").replace("\t", "\\tab "));
            }
        }
    }


    private String cleanHtmlContent(String rtfContent) {
        // Placeholder for any additional RTF cleanup
        return rtfContent;
    }






    private int calculateCellWidth(Element cell) {
        // Logic to calculate cell width based on content or other factors
        return 5000; // Placeholder value
    }










//    private void saveMessage(MapiMessage message, String folderPath) {
//        try {
//            logger.info("Starting to process the message: " + message.getSubject());
//
//            // Determine the base file name
//            String baseFileName = cleanFileName(message.getSubject());
//            if (baseFileName.isEmpty() || baseFileName.equals("Unnamed")) {
//                baseFileName = "Unnamed";
//            }
//            String fileName = baseFileName + ".rtf";
//            File file = new File(folderPath, fileName);
//            int count = 1;
//
//            while (file.exists()) {
//                fileName = baseFileName + "_" + count + ".rtf";
//                file = new File(folderPath, fileName);
//                count++;
//            }
//
//            // Check the body type and get the appropriate body content
//            String bodyContent;
//            if (message.getBodyType() == BodyContentType.Html) {
//                bodyContent = message.getBodyHtml(); // Use getBodyHtml() if available
//            } else if (message.getBodyType() == BodyContentType.Rtf) {
//                bodyContent = message.getBodyRtf(); // Use getBodyRtf() if available
//            } else {
//                bodyContent = message.getBody(); // Fallback to plain text
//            }
//
//            if (bodyContent == null || bodyContent.isEmpty()) {
//                logger.warning("Email body is empty or null for: " + fileName);
//            } else {
//                logger.info("Body Content Length: " + bodyContent.length());
//                logger.info("Body Content Preview: " + bodyContent.substring(0, Math.min(100, bodyContent.length())));
//
//                // If the content is HTML, clean it using Jsoup
//                if (message.getBodyType() == BodyContentType.Html) {
//                    bodyContent = Jsoup.clean(bodyContent, "", Safelist.none());
//                }
//
//                try (FileOutputStream out = new FileOutputStream(file)) {
//                    RTFEditorKit rtfEditorKit = new RTFEditorKit();
//                    javax.swing.text.Document doc = rtfEditorKit.createDefaultDocument();
//                    bodyContent = cleanText(bodyContent);
//                    // Insert the cleaned text into the RTF document
//                    doc.insertString(0, bodyContent, null);
//
//                    // Write the RTF document to file
//                    rtfEditorKit.write(out, doc, 0, doc.getLength());
//                    logger.info("Saved email as RTF: " + fileName);
//                } catch (IOException | BadLocationException e) {
//                    logger.warning("Error saving email as RTF: " + e.getMessage());
//                }
//                if (!message.getAttachments().isEmpty()) {
//                    File attachmentDir = new File(folderPath + "/Attachments");
//                    if (!attachmentDir.exists()) {
//                        attachmentDir.mkdirs();
//                    }
//                    saveContactAttachment(message, attachmentDir);
//                }
//            }
//        } catch (Exception e) {
//            logger.warning("Error processing message for RTF conversion: " + e.getMessage());
//        }
//    }

    private void saveContactAttachment(MapiMessage message, File attachmentDir) {
        File directory = new File(attachmentDir.getAbsolutePath() + File.separator + cleanFileName(message.getSubject()).trim());
        if (directory.exists()) {
            directory = new File(attachmentDir.getAbsolutePath() + File.separator + cleanFileName(message.getSubject()).trim() + "_" + System.currentTimeMillis());
        }
        directory.mkdir();

        for (MapiAttachment attachment : message.getAttachments()) {
            try {
                attachment.save(directory.getAbsolutePath() + File.separator + cleanFileName(attachment.getDisplayName()).trim());
            } catch (Exception e) {
                attachment.save(directory.getAbsolutePath() + File.separator + cleanFileName(attachment.getLongFileName()).trim());
            }
        }
    }

    private void saveContact(MapiMessage message, String folderPath) {
        try {
            logger.info("Starting to process the message: " + message.getSubject());

            // Determine the base file name
            String baseFileName = cleanFileName(message.getSubject());
            if (baseFileName.isEmpty() || baseFileName.equals("Unnamed")) {
                baseFileName = "Unnamed";
            }
            String fileName = baseFileName + ".rtf";
            File file = new File(folderPath, fileName);
            int count = 1;

            while (file.exists()) {
                fileName = baseFileName + "_" + count + ".rtf";
                file = new File(folderPath, fileName);
                count++;
            }

            // Check the body type and get the appropriate body content
            String bodyContent;
            if (message.getBodyType() == BodyContentType.Html) {
                bodyContent = message.getBodyHtml();
            } else if (message.getBodyType() == BodyContentType.Rtf) {
                bodyContent = message.getBodyRtf();
            } else {
                bodyContent = message.getBody();
            }

            if (bodyContent == null || bodyContent.isEmpty()) {
                logger.warning("Email body is empty or null for: " + fileName);
            } else {
                logger.info("Body Content Length: " + bodyContent.length());
                logger.info("Body Content Preview: " + bodyContent.substring(0, Math.min(100, bodyContent.length())));

                // If the content is HTML, clean it using Jsoup
                if (message.getBodyType() == BodyContentType.Html) {
                    bodyContent = Jsoup.clean(bodyContent, "", Safelist.none());
                    // Remove common unwanted characters and replace HTML entities
                }

                try (FileOutputStream out = new FileOutputStream(file)) {
                    RTFEditorKit rtfEditorKit = new RTFEditorKit();
                    javax.swing.text.Document doc = rtfEditorKit.createDefaultDocument();
                    bodyContent = cleanText(bodyContent);
                    // Insert the cleaned plain text into the RTF document
                    doc.insertString(0, bodyContent, null);

                    // Write the RTF document to file
                    rtfEditorKit.write(out, doc, 0, doc.getLength());
                    logger.info("Saved email as RTF: " + fileName);
                } catch (IOException | BadLocationException e) {
                    logger.warning("Error saving email as RTF: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            logger.warning("Error processing message for RTF conversion: " + e.getMessage());
        }
    }


    private void saveCalendar(MapiMessage message, String folderPath) {
        try {
            MapiCalendar calendar = (MapiCalendar) MapiMessage.fromMailMessage(String.valueOf(message)).toMapiMessageItem();
            if (!isDuplicateCalendar(calendar)) {
                String baseFileName = cleanFileName(calendar.getSubject());
                if (baseFileName.isEmpty() || baseFileName.equals("Unnamed")) {
                    System.out.println("Unnamed calendar event found");
                    baseFileName = "Unnamed";
                }
                String fileName = baseFileName + ".rtf";
                File file = new File(folderPath, fileName);
                int count = 1;

                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".rtf";
                    file = new File(folderPath, fileName);
                    count++;
                }

                // Use Apache POI to create an RTF document for calendar
                try (FileOutputStream out = new FileOutputStream(file)) {
                    HWPFDocument doc = new HWPFDocument(new POIFSFileSystem());
                    Range range = doc.getRange();
                    range.insertAfter("Subject: " + calendar.getSubject() + "\n");
                    range.insertAfter("Location: " + calendar.getLocation() + "\n");
                    range.insertAfter("Start: " + calendar.getStartDate() + "\n");
                    range.insertAfter("End: " + calendar.getEndDate() + "\n");

                    doc.write(out);
                    logger.info("Saved calendar as RTF: " + fileName);
                } catch (IOException e) {
                    logger.warning("Error saving calendar as RTF: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            logger.warning("Error processing calendar for RTF conversion: " + e.getMessage());
        }
    }
    private void saveCalendarAsICS(MapiMessage message, String folderPath) {
        try {
            MapiCalendar calendar = (MapiCalendar) message.toMapiMessageItem();
            if (!isDuplicateCalendar(calendar) || !enableDuplicateCheck) {
                // Save as .msg
                String baseFileName = cleanFileName(calendar.getSubject());
                int count = 1;
                // Save as .ics
                String icsFileName = baseFileName + ".ics";
                File icsFile = new File(folderPath, icsFileName);
                count = 1;

                while (icsFile.exists()) {
                    icsFileName = baseFileName + "_" + count + ".ics";
                    icsFile = new File(folderPath, icsFileName);
                    count++;
                }
                // Save the calendar as ICS format
                calendar.save(icsFile.getAbsolutePath(), AppointmentSaveFormat.Ics);
                logger.info("Saved calendar as ICS: " + icsFileName);
            }
        } catch (Exception e) {
            logger.warning("Error saving calendar: " + e.getMessage());
        }
    }

    private void saveTask(MapiMessage message, String folderPath) {
        try {
            MapiTask task = (MapiTask) MapiMessage.fromMailMessage(String.valueOf(message)).toMapiMessageItem();
            if (!isDuplicateTask(task)) {
                String baseFileName = cleanFileName(task.getSubject());
                if (baseFileName.isEmpty() || baseFileName.equals("Unnamed")) {
                    System.out.println("Unnamed task found");
                    baseFileName = "Unnamed";
                }
                String fileName = baseFileName + ".rtf";
                File file = new File(folderPath, fileName);
                int count = 1;

                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".rtf";
                    file = new File(folderPath, fileName);
                    count++;
                }

                // Use Apache POI to create an RTF document for task
                try (FileOutputStream out = new FileOutputStream(file)) {
                    HWPFDocument doc = new HWPFDocument(new POIFSFileSystem());
                    Range range = doc.getRange();
                    range.insertAfter("Subject: " + task.getSubject() + "\n");
                    range.insertAfter("Start: " + task.getStartDate() + "\n");
                    range.insertAfter("Due: " + task.getDueDate() + "\n");
                    range.insertAfter("Status: " + task.getStatus() + "\n");
                    range.insertAfter(task.getBody());

                    doc.write(out);
                    logger.info("Saved task as RTF: " + fileName);
                } catch (IOException e) {
                    logger.warning("Error saving task as RTF: " + e.getMessage());
                }
            }
        } catch (Exception e) {
            logger.warning("Error processing task for RTF conversion: " + e.getMessage());
        }
    }


    private void saveGenericMessage(MapiMessage message, String folderPath) {
        try {
            logger.info("Starting to process the message: " + message.getSubject());
            // Determine the base file name
            String baseFileName = cleanFileName(message.getSubject());
            if (baseFileName.isEmpty() || baseFileName.equals("Unnamed")) {
                baseFileName = "Unnamed";
            }
            String fileName = baseFileName + ".rtf";
            File file = new File(folderPath, fileName);
            int count = 1;

            while (file.exists()) {
                fileName = baseFileName + "_" + count + ".rtf";
                file = new File(folderPath, fileName);
                count++;
            }
                // Parse the body content using Jsoup to remove HTML tags
                String plainText = Jsoup.parse(message.getBody()).text();
                logger.info("Body Content Length: " + plainText.length());
                logger.info("Body Content Preview: " + plainText.substring(0, Math.min(100, plainText.length())));
                try (FileOutputStream out = new FileOutputStream(file)) {
                    RTFEditorKit rtfEditorKit = new RTFEditorKit();
                    javax.swing.text.Document doc = rtfEditorKit.createDefaultDocument();

                    // Insert the cleaned plain text into the RTF document
                    doc.insertString(0, plainText, null);

                    // Write the RTF document to file
                    rtfEditorKit.write(out, doc, 0, doc.getLength());
                    logger.info("Saved email as RTF: " + fileName);
                } catch (IOException | BadLocationException e) {
                    logger.warning("Error saving email as RTF: " + e.getMessage());
                }

        } catch (Exception e) {
            logger.warning("Error processing message for RTF conversion: " + e.getMessage());
        }
    }

    private boolean isDuplicateCalendar(MapiCalendar calendar) {
        String key = calendar.getLocation() + calendar.getStartDate() + calendar.getEndDate();
        if (setDupliccal.contains(key)) {
            return true;
        } else {
            setDupliccal.add(key);
            return false;
        }
    }

    private boolean isDuplicateContact(MapiContact contact) {
        String key = contact.getNameInfo().getDisplayName() + contact.getPersonalInfo().getNotes();
        if (setDupliccontact.contains(key)) {
            return true;
        } else {
            setDupliccontact.add(key);
            return false;
        }
    }

    private boolean isDuplicateTask(MapiTask task) {
        String key = task.getSubject() + task.getBody();
        if (setDuplictask.contains(key)) {
            return true;
        } else {
            setDuplictask.add(key);
            return false;
        }
    }

    private boolean isDuplicateMessage(MailMessage message) {
        String key = message.getSubject() + message.getBody();
        if (setDuplicacy.contains(key)) {
            return true;
        } else {
            setDuplicacy.add(key);
            return false;
        }
    }
    private String cleanText(String text) {
        if (text == null) {
            return "";
        }

        // Remove &nbsp; and other specific HTML entities
        text = text.replaceAll("&nbsp;", " ").replaceAll("&[a-zA-Z0-9#]+;", "");

        // Remove multiple spaces or other non-visible characters
        text = text.replaceAll("\\s+", " ").trim();

        // Remove remaining non-printable characters
        text = text.replaceAll("[^\\p{Print}]", "");

        // Specific removal of any lingering unwanted characters
        text = text.replaceAll("(?i)\\s*evaluation only. created with aspose.email for java. copyright\\s+\\d{4}-\\d{4} aspose pty ltd.view eula online\\s*", "");

        return text;
    }

    private String cleanFileName(String fileName) {
        return (fileName != null ? fileName : "Unnamed").replaceAll("[\\\\/:*?\"<>|]", "_").trim();
    }
    private String cleanFolderName(String folderName) {
        if (folderName == null || folderName.trim().isEmpty()) {
            return "Unnamed";
        }
        folderName = folderName.replace(",", "").replace(".", "");
        folderName = folderName.replaceAll("[\\[\\]]", "");
        return folderName.trim();
    }

    private String generateFolderNameFromPath(String path) {
        return path.replace("/", "_").replace("\\", "_").replace(":", "_");
    }
}