package com.NAtools.Convertor;


import com.NAtools.config.LogManagerConfig;
import com.NAtools.util.ImageUtil;
import com.aspose.email.*;
import com.openhtmltopdf.extend.FSStream;
import com.openhtmltopdf.extend.FSStreamFactory;
import com.openhtmltopdf.pdfboxout.PdfBoxRenderer;
import com.sun.xml.internal.ws.transport.http.ResourceLoader;
import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.jsoup.helper.W3CDom;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.w3c.dom.DOMException;
import org.xml.sax.InputSource; // Corrected import
import org.xml.sax.SAXException; // Corrected import
import org.xml.sax.SAXParseException; // Corrected import
import org.apache.pdfbox.pdmodel.*;
        import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.safety.Safelist;
import com.openhtmltopdf.pdfboxout.PdfRendererBuilder;

import javax.swing.text.rtf.RTFEditorKit;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.*;
        import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


public class OSTOPDF_Test {
    private static final Logger logger = Logger.getLogger(OSTToPDF.class.getName());
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

    public OSTOPDF_Test(PersonalStorage sourcePst, String outputDirectory, String sourcePstPath, boolean enableDuplicateCheck) {
        this.sourcePst = sourcePst;
        this.outputDirectory = outputDirectory;
        this.enableDuplicateCheck = enableDuplicateCheck;
        this.ostFileName = new File(sourcePstPath).getName().replaceFirst("[.][^.]+$", "");
    }

    public void convertToPDF() {
        try {
            FolderInfo rootFolder = sourcePst.getRootFolder();
            processFolder(rootFolder, outputDirectory);
            logger.info("Conversion to PDF format completed.");
        } catch (Exception e) {
            logger.severe("Error during PDF conversion: " + e.getMessage());
        } finally {
            sourcePst.dispose();
        }
    }

    public void processFolder(FolderInfo folder, String parentPath) {
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
                    MapiMessage message = sourcePst.extractMessage(messageInfo);
                    switch (message.getMessageClass()) {
                        case "IPM.Note":
                            saveMessageAsPDF(message, folderPath);
                            break;
                        case "IPM.Contact":
                            saveContactAsPDF(message, folderPath);
                            break;
                        case "IPM.Appointment":
                        case "IPM.Schedule.Meeting.Request":
                            saveCalendarAsICS(message, folderPath);
                            break;
                        case "IPM.Task":
                            saveTaskAsPDF(message, folderPath);
                            break;
                        default:
                            saveMessageAsPDF(message, folderPath);
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

    public void saveMessageAsPDF(MapiMessage message, String folderPath) {
        OutputStream os = null;
        try {
            String subject = message.getSubject();
            if (subject == null || subject.isEmpty()) {
                subject = "Unnamed"; // Fallback if subject is null or empty
            }
            logger.info("Starting to process the message: " + subject);

            if(subject.contains("forget")){
                logger.info("Debugging injected images");
            }
            // Prepare the email headers
            StringBuilder headers = new StringBuilder();
            headers.append("<div style=\"font-family: Arial, sans-serif; font-size: 12px; margin-bottom: 20px;\">")
                    .append("<strong>From:</strong> ").append(resolveSender(message)).append("<br/>")
                    .append("<strong>Sent:</strong> ").append(message.getDeliveryTime().toString()).append("<br/>")
                    .append("<strong>To:</strong> ").append(extractRecipients(message.getRecipients())).append("<br/>")
                    .append("<strong>Subject:</strong> ").append(message.getSubject()).append("<br/>");

            // Append attachments info if present
            if (!message.getAttachments().isEmpty()) {
                headers.append("<strong>Attachments:</strong> ");
                for (MapiAttachment attachment : message.getAttachments()) {
                    headers.append(attachment.getDisplayName()).append(", ");
                }
                headers.setLength(headers.length() - 2); // Remove trailing comma
                headers.append("<br/>");
            }
            headers.append("</div><hr/>"); // Close header div and add a horizontal line

            // Determine the message format and convert to HTML if necessary
            String htmlContent = null;
            if (message.getBodyType() == BodyContentType.Html) {
                htmlContent = message.getBodyHtml();
            } else if (message.getBodyType() == BodyContentType.PlainText || message.getBodyType() == BodyContentType.Rtf) {
               // htmlContent = "<html><body><pre>" + message.getBody() + "</pre></body></html>"; // Wrap plain text in HTML
                htmlContent = convertPlainTextToHtml(message.getBody());
            }

            if (htmlContent == null ) { // It's ok to have empty body
                logger.warning("Message has no content to process.");
                return;
            }
            logger.info("Extracted content for message: " + subject + ". Content length: " + htmlContent.length());

            htmlContent = cleanHtmlComments(htmlContent);
            // Extract CID images and replace references in HTML
            Map<String, String> cidToImageMap = new HashMap<>();
            if (!message.getAttachments().isEmpty()) {
                File attachmentDir = new File(folderPath + "/Attachments");
                if (!attachmentDir.exists()) {
                    attachmentDir.mkdirs();
                }
                cidToImageMap = saveAndMapCidAttachments(message, attachmentDir);
            }

            for (Map.Entry<String, String> entry : cidToImageMap.entrySet()) {
                String cid = entry.getKey(); // e.g., "image001.png"
                String filePath = entry.getValue(); // e.g., "C:/images/image001.png"
                String uriPath = new File(filePath).toURI().toString();
                htmlContent = htmlContent.replace("cid:" + cid, uriPath);
            }

            // Parse HTML content using Jsoup to safely inject headers
            Document jsoupDoc = Jsoup.parse(htmlContent);
            jsoupDoc.outputSettings()
                    .escapeMode(org.jsoup.nodes.Entities.EscapeMode.xhtml)
                    .syntax(Document.OutputSettings.Syntax.xml); // Ensure XHTML compliance

            // Inject headers into the body
            Element body = jsoupDoc.body();
            body.prepend(headers.toString());  // Prepend headers to the body content

            String cleanedHtml = jsoupDoc.outerHtml();
            logger.info("Cleaned HTML entities for message: " + subject + ". Cleaned content length: " + cleanedHtml.length());

            // Replace common HTML entities with their numeric equivalents
            cleanedHtml = cleanedHtml.replaceAll("&nbsp;", "&#160;");

            if (subject.contains("Welcome")){
                logger.info("********* check ***********");
            }
            // Convert cleaned HTML to W3C DOM Document
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            InputSource is = new InputSource(new StringReader(cleanedHtml));
            org.w3c.dom.Document w3cDoc = builder.parse(is); // org.w3c.dom.Document

            if (w3cDoc == null) {
                logger.warning("w3cDoc is null after parsing HTML content.");
                return;
            }

            // Determine the file name and path
            String baseFileName = cleanFileName(subject);
            if (baseFileName.isEmpty()) {
                baseFileName = "Unnamed";
            }
            String fileName = baseFileName + ".pdf";
            File file = new File(folderPath, fileName);
            int count = 1;

            while (file.exists()) {
                fileName = baseFileName + "_" + count + ".pdf";
                file = new File(folderPath, fileName);
                count++;
            }

            // Open output stream for PDF
            os = Files.newOutputStream(file.toPath());
            if (os == null) {
                logger.warning("OutputStream is null.");
                return;
            }

            // Use OpenHTMLtoPDF to create the PDF
            PdfRendererBuilder pdfBuilder = new PdfRendererBuilder();
            pdfBuilder.useFastMode(); // Optional: enables faster rendering
            pdfBuilder.withW3cDocument(w3cDoc, null);
            pdfBuilder.toStream(os);

            // Additional logging before running the PDF build process
            logger.info("Starting PDF generation for message: " + subject);

            pdfBuilder.run();
            logger.info("Saved email as PDF: " + fileName);

        } catch (SAXException | IOException | ParserConfigurationException e) {
            logger.warning("Error parsing or saving message as PDF: " + e.getMessage());
        } catch (Exception e) {
            logger.warning("General Error saving message as PDF: " + e.getMessage());
            e.printStackTrace();  // Print stack trace for deeper insight
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {
                    logger.warning("Error closing OutputStream: " + e.getMessage());
                }
            }
        }
    }



    public String embedCIDImagesUsingJsoup(String htmlContent, Map<String, byte[]> cidToImageMap) {
        try {
            // Parse the HTML content into a Document
            Document document = Jsoup.parse(htmlContent);

            // Select all img elements with src starting with "cid:"
            Elements imgTags = document.select("img[src^=cid:]");

            // Iterate over each img tag and replace the CID reference with Base64 data
            for (Element img : imgTags) {
                String src = img.attr("src");
                String cid = src.substring(4);  // Extract CID without "cid:"

                // Retrieve the image data from the map
                byte[] imageData = cidToImageMap.get(cid);

                if (imageData != null) {
                    // Convert image data to Base64
                    String base64Image = "data:image/png;base64," + java.util.Base64.getEncoder().encodeToString(imageData);

                    // Set the src attribute to the Base64 string
                    img.attr("src", base64Image);
                }
            }

            // Return the modified HTML as a string
            return document.outerHtml();

        } catch (Exception e) {
            e.printStackTrace();
            return htmlContent; // Return original content in case of error
        }
    }


    private Map<String, byte[]> saveAndMapCidAttachments(MapiMessage message) {
        Map<String, byte[]> cidToImageMap = new HashMap<>();

        for (MapiAttachment attachment : message.getAttachments()) {
            try {
                String cid = attachment.getItemId();
                if (cid != null && !cid.isEmpty()) {
                    byte[] imageData = attachment.getBinaryData();
                    cidToImageMap.put(cid, imageData);
                }
            } catch (Exception e) {
                logger.warning("Failed to save attachment or map CID: " + e.getMessage());
            }
        }

        return cidToImageMap;
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


//    private void downloadAndReplaceResources(String htmlContent, File saveDirectory,MapiMessage message) {
//        // Regex to find all image tags with src attributes
//        //Pattern pattern = Pattern.compile("<img[^>]+src=\"(http[^\"]+)\"", Pattern.CASE_INSENSITIVE);
//        Pattern pattern = Pattern.compile("<a[^>]+href=\"(https?://[^\"]+)\"", Pattern.CASE_INSENSITIVE);
//
//        Matcher matcher = pattern.matcher(htmlContent);
//
//        while (matcher.find()) {
//            String url = matcher.group(1);
//            try {
//                // Download the resource
//                String fileName = Paths.get(new URL(url).getPath()).getFileName().toString() ;
//                File savedFile = new File(saveDirectory.getAbsolutePath(), fileName);
//                try (InputStream in = new URL(url).openStream()) {
//                    Files.copy(in, savedFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
//                }
//
//                // Update the HTML content with the local file path
//                String localPath = savedFile.toURI().toString();
//                htmlContent = htmlContent.replace(url, localPath);
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//        }
//    }
private void downloadAndReplaceResources(String htmlContent, File saveDirectory, MapiMessage message) {
    // Create a subdirectory using message ID and subject to avoid name clashes
    String messageId = message.getInternetMessageId();
    String subject = cleanFileName(message.getSubject()).trim();
    File messageDirectory = new File(saveDirectory.getAbsolutePath(),  subject);

    if (messageDirectory.exists()) {
        messageDirectory = new File(saveDirectory.getAbsolutePath() + File.separator + cleanFileName(message.getSubject()).trim() + "_" + System.currentTimeMillis());
    }
    messageDirectory.mkdirs();

    // Regex to find all href attributes in anchor tags
    Pattern pattern = Pattern.compile("<img[^>]+src=\"(http[^\"]+)\"", Pattern.CASE_INSENSITIVE);
    Matcher matcher = pattern.matcher(htmlContent);

    while (matcher.find()) {
        String url = matcher.group(1);
        try {
            // Download the resource
            String fileName = Paths.get(new URL(url).getPath()).getFileName().toString();
            File savedFile = new File(messageDirectory.getAbsolutePath(), fileName);
            try (InputStream in = new URL(url).openStream()) {
                Files.copy(in, savedFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
            }

            // Update the HTML content with the local file path
            String localPath = savedFile.toURI().toString();
            htmlContent = htmlContent.replace(url, localPath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}


    private String extractHtmlContent(MapiMessage message) {
        try {
            logger.info("Extracting content from message: " + message.getSubject());
            if (message.getBodyType() == BodyContentType.Html) {
                return message.getBodyHtml();
            } else if (message.getBodyType() == BodyContentType.Rtf) {
                return message.getBodyRtf(); // Conversion function to be implemented
            } else {
                return message.getBody();
            }
        } catch (Exception e) {
            logger.warning("Error extracting content from message: " + e.getMessage());
            return "";
        }
    }

    private String cleanHtmlComments(String htmlContent) {
        return htmlContent.replaceAll("<!--.*?-->", ""); // Basic example, can be adjusted for more complex cases
    }

    // Method to extract CID images from the MapiMessage
    private Map<String, byte[]> extractCidToImageMap(MapiMessage message) {
        Map<String, byte[]> cidToImageMap = new HashMap<>();
        int attachmentCounter = 1;  // Counter to generate unique keys

        for (MapiAttachment attachment : message.getAttachments()) {
            try {
                // Generate a unique key for each attachment
                String generatedContentId = "generated-cid-" + attachmentCounter;
                byte[] imageData = attachment.getBinaryData();

                // Store the generated key and image data in the map
                cidToImageMap.put(generatedContentId, imageData);
                attachmentCounter++;
            } catch (Exception e) {
                logger.warning("Error extracting or processing attachment: " + e.getMessage());
            }
        }

         return cidToImageMap;
    }

    private String embedCIDImagesAsBase64(String htmlContent, Map<String, byte[]> cidToImageMap) {
        for (Map.Entry<String, byte[]> entry : cidToImageMap.entrySet()) {
            String cid = entry.getKey();
            byte[] imageBytes = entry.getValue();
            String base64Image = Base64.getEncoder().encodeToString(imageBytes);
            String imgTag = "\"data:image/png;base64, " + base64Image + "\"";

            htmlContent = htmlContent.replace("cid:" , imgTag);
        }
        return htmlContent;
    }



    private String cleanFileName(String fileName) {
        return (fileName != null ? fileName : "Unnamed").replaceAll("[\\\\/:*?\"<>|]", "_").trim();
    }

    // Utility method for word wrapping text
    private String wordWrap(String text, PDFont font, float fontSize, float width) throws IOException {
        StringBuilder result = new StringBuilder();
        String[] words = text.split(" ");
        float spaceWidth = font.getStringWidth(" ") / 1000 * fontSize;

        float lineWidth = 0;

        for (String word : words) {
            float wordWidth = font.getStringWidth(word) / 1000 * fontSize;

            if (lineWidth + wordWidth > width) {
                result.append("\n");
                lineWidth = 0;
            }

            result.append(word).append(" ");
            lineWidth += wordWidth + spaceWidth;
        }

        return result.toString().trim();
    }


    // Utility method to extract recipients from MapiRecipientCollection
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


    private String[] formatBodyText(String text, int maxLineLength) {
        // Split the text into lines of up to `maxLineLength` characters
        String[] words = text.split(" ");
        StringBuilder line = new StringBuilder();
        StringBuilder result = new StringBuilder();

        for (String word : words) {
            if (line.length() + word.length() > maxLineLength) {
                result.append(line).append("\n");
                line = new StringBuilder(word).append(" ");
            } else {
                line.append(word).append(" ");
            }
        }
        result.append(line);  // Append the last line
        return result.toString().split("\n");
    }







    private String cleanText(String text) {
        if (text == null) {
            return "";
        }

        // Replace carriage returns with newlines or spaces
        text = text.replaceAll("\r\n", "\n").replaceAll("\r", "\n");

        // Remove any unsupported characters
        text = text.replaceAll("[^\\p{Print}]", "");

        return text;
    }


    private String[] wrapText(String text, int maxCharsPerLine) {
        String[] words = text.split(" ");
        StringBuilder line = new StringBuilder();
        List<String> lines = new ArrayList<>();

        for (String word : words) {
            if (line.length() + word.length() > maxCharsPerLine) {
                lines.add(line.toString());
                line = new StringBuilder();
            }
            if (line.length() > 0) {
                line.append(" ");
            }
            line.append(word);
        }
        lines.add(line.toString()); // Add the last line

        return lines.toArray(new String[0]);
    }



    private void saveContactAsPDF(MapiMessage message, String folderPath) {
        try {
            MapiContact contact = (MapiContact) message.toMapiMessageItem();
            if (!isDuplicateContact(contact) || !enableDuplicateCheck) {
                String baseFileName = cleanFileName(contact.getNameInfo().getDisplayName());
                String fileName = baseFileName + ".pdf";
                File file = new File(folderPath, fileName);
                int count = 1;

                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".pdf";
                    file = new File(folderPath, fileName);
                    count++;
                }

                try (PDDocument document = new PDDocument()) {
                    PDPage page = new PDPage(PDRectangle.A4);
                    document.addPage(page);
                    try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
                        contentStream.setFont(PDType1Font.HELVETICA, 12);

                        contentStream.beginText();
                        contentStream.newLineAtOffset(25, 750);
                        contentStream.showText("Contact Name: " + contact.getNameInfo().getDisplayName());
                        contentStream.newLine();
                        contentStream.showText("Email: " + contact.getElectronicAddresses().getEmail1());
                        contentStream.newLine();
                        contentStream.showText("Phone: " + contact.getTelephones().getBusinessTelephoneNumber());
                        contentStream.endText();
                    }

                    document.save(file);
                    logger.info("Saved contact as PDF: " + fileName);
                }
            }
        } catch (IOException e) {
            logger.warning("Error saving contact as PDF: " + e.getMessage());
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

    private void saveCalendarAsPDF(MapiMessage message, String folderPath) {
        try {
            MapiCalendar calendar = (MapiCalendar) message.toMapiMessageItem();
            if (!isDuplicateCalendar(calendar) || !enableDuplicateCheck) {
                String baseFileName = cleanFileName(calendar.getSubject());
                String fileName = baseFileName + ".pdf";
                File file = new File(folderPath, fileName);
                int count = 1;

                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".pdf";
                    file = new File(folderPath, fileName);
                    count++;
                }

                try (PDDocument document = new PDDocument()) {
                    PDPage page = new PDPage(PDRectangle.A4);
                    document.addPage(page);
                    try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
                        contentStream.setFont(PDType1Font.HELVETICA, 12);

                        contentStream.beginText();
                        contentStream.newLineAtOffset(25, 750);
                        contentStream.showText("Subject: " + calendar.getSubject());
                        contentStream.newLine();
                        contentStream.showText("Location: " + calendar.getLocation());
                        contentStream.newLine();
                        contentStream.showText("Start: " + calendar.getStartDate());
                        contentStream.newLine();
                        contentStream.showText("End: " + calendar.getEndDate());
                        contentStream.endText();
                    }

                    document.save(file);
                    logger.info("Saved calendar as PDF: " + fileName);
                }
            }
        } catch (IOException e) {
            logger.warning("Error saving calendar as PDF: " + e.getMessage());
        }
    }
    private void saveTaskAsPDF (MapiMessage message, String folderPath){
        try {
            MapiTask task = (MapiTask) message.toMapiMessageItem();
            if (!isDuplicateTask(task) || !enableDuplicateCheck) {
                String baseFileName = cleanFileName(task.getSubject());
                if (baseFileName.isEmpty() || baseFileName.equals("Unnamed")) {
                    baseFileName = "Unnamed";
                }
                String fileName = baseFileName + ".pdf";
                File file = new File(folderPath, fileName);
                int count = 1;

                while (file.exists()) {
                    fileName = baseFileName + "_" + count + ".pdf";
                    file = new File(folderPath, fileName);
                    count++;
                }

                try (PDDocument document = new PDDocument()) {
                    PDPage page = new PDPage(PDRectangle.A4);
                    document.addPage(page);
                    try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
                        contentStream.setFont(PDType1Font.HELVETICA, 12);

                        contentStream.beginText();
                        contentStream.newLineAtOffset(25, 750);
                        contentStream.showText("Subject: " + task.getSubject());
                        contentStream.newLine();
                        contentStream.showText("Start: " + task.getStartDate());
                        contentStream.newLine();
                        contentStream.showText("Due: " + task.getDueDate());
                        contentStream.newLine();
                        contentStream.showText("Status: " + task.getStatus());
                        contentStream.newLine();
                        contentStream.showText("Body:");
                        contentStream.newLine();
                        contentStream.showText(task.getBody());
                        contentStream.endText();
                    }

                    document.save(file);
                    logger.info("Saved task as PDF: " + fileName);
                }
            }
        } catch (IOException e) {
            logger.warning("Error saving task as PDF: " + e.getMessage());
        }
    }

    private boolean isDuplicateCalendar (MapiCalendar calendar){
        String key = calendar.getLocation() + calendar.getStartDate() + calendar.getEndDate();
        if (setDupliccal.contains(key)) {
            return true;
        } else {
            setDupliccal.add(key);
            return false;
        }
    }

    private boolean isDuplicateContact (MapiContact contact){
        String key = contact.getNameInfo().getDisplayName() + contact.getPersonalInfo().getNotes();
        if (setDupliccontact.contains(key)) {
            return true;
        } else {
            setDupliccontact.add(key);
            return false;
        }
    }

    private boolean isDuplicateTask (MapiTask task){
        String key = task.getSubject() + task.getBody();
        if (setDuplictask.contains(key)) {
            return true;
        } else {
            setDuplictask.add(key);
            return false;
        }
    }


    private String cleanFolderName (String folderName){
        if (folderName == null || folderName.trim().isEmpty()) {
            return "Unnamed";
        }
        folderName = folderName.replace(",", "").replace(".", "");
        folderName = folderName.replaceAll("[\\[\\]]", "");
        return folderName.trim();
    }

    private String getBodyContent (MapiMessage message){
        if (message.getBodyType() == BodyContentType.Html) {
            return Jsoup.clean(message.getBodyHtml(), "", Safelist.none());
        } else if (message.getBodyType() == BodyContentType.Rtf) {
            return message.getBodyRtf();
        } else {
            return message.getBody();
        }
    }

    private class CIDFSStreamFactory implements FSStreamFactory {

        private final Map<String, byte[]> cidToImageMap;

        public CIDFSStreamFactory(Map<String, byte[]> cidToImageMap) {
            this.cidToImageMap = cidToImageMap;
        }

        @Override
        public FSStream getUrl(String url) {
            return new FSStream() {
                @Override
                public InputStream getStream() {
                    String cid = url.substring("cid:".length());
                    byte[] imageData = cidToImageMap.get(cid);
                    if (imageData != null) {
                        return new ByteArrayInputStream(imageData);
                    }
                    throw new IllegalArgumentException("CID not found: " + cid);
                }

                @Override
                public Reader getReader() {
                    return new InputStreamReader(getStream(), StandardCharsets.UTF_8);
                }
            };
        }
    }
}