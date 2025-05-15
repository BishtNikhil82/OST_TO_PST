package com.NAtools.Convertor;
import com.NAtools.config.LogManagerConfig;
import com.NAtools.util.ImageUtil;
import com.aspose.email.*;
import org.apache.pdfbox.pdmodel.font.PDFont;
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

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.logging.Logger;


public class OSTToPDF {
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

    public OSTToPDF(PersonalStorage sourcePst, String outputDirectory, String sourcePstPath, boolean enableDuplicateCheck) {
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
                            saveCalendarAsPDF(message, folderPath);
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



    public void saveMessageAsPDF(MapiMessage message, String folderPath) {
        OutputStream os = null;
        try {
            logger.info("Starting to process the message: " + message.getSubject());

            // Extract the HTML content from the email
            String htmlContent = extractHtmlContent(message);
            logger.info("Extracted HTML content for message: " + message.getSubject());

            // Clean up the HTML content to avoid parsing issues with comments
            htmlContent = cleanHtmlComments(htmlContent);
            logger.info("Cleaned HTML comments for message: " + message.getSubject());

            // Initialize DocumentBuilder
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();

            // Parse and clean the HTML content using Jsoup to ensure it is well-formed XHTML.
            Document jsoupDoc = Jsoup.parse(htmlContent);
            jsoupDoc.outputSettings()
                    .escapeMode(org.jsoup.nodes.Entities.EscapeMode.xhtml)
                    .syntax(Document.OutputSettings.Syntax.xml); // Ensure XHTML compliance

            String cleanedHtml = jsoupDoc.outerHtml();
            logger.info("Cleaned HTML entities for message: " + message.getSubject());

            // Replace common HTML entities with their numeric equivalents
            cleanedHtml = cleanedHtml.replaceAll("&nbsp;", "&#160;");

            // Convert cleaned HTML to W3C DOM Document
            InputSource is = new InputSource(new StringReader(cleanedHtml));
            org.w3c.dom.Document w3cDoc = builder.parse(is); // org.w3c.dom.Document

            if (w3cDoc == null) {
                logger.warning("w3cDoc is null after parsing HTML content.");
                return;
            }

            // Determine the file name and path
            String baseFileName = cleanFileName(message.getSubject());
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
            logger.info("Starting PDF generation for message: " + message.getSubject());

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




    private String extractHtmlContent(MapiMessage message) {
        try {
            // Handle different email formats
            if (message.getBodyType() == BodyContentType.Html) {
                return message.getBodyHtml();
            } else if (message.getBodyType() == BodyContentType.Rtf) {
                return message.getBodyRtf();
            } else {
                // Handle multipart/related with alternative parts
                if (message.getMessageClass().equals("IPM.Note") || message.getMessageClass().equals("IPM.Schedule.Meeting.Request")) {
                    MapiProperty prop = message.getProperties().get_Item(MapiPropertyTag.PR_HTML);
                    if (prop != null) {
                        try {
                            return new String(prop.getData(), "Windows-1252"); // Use appropriate encoding
                        } catch (UnsupportedEncodingException e) {
                            logger.warning("Unsupported encoding: " + e.getMessage());
                        }
                    }
                }
            }
        } catch (Exception e) {
            logger.warning("Error extracting HTML content: " + e.getMessage());
        }
        return ""; // Fallback
    }


    private String cleanHtmlComments(String html) {
        // Remove problematic comments that contain "--"
        return html.replaceAll("<!--(.*?)-->", ""); // Remove all HTML comments
    }

    private MapiAttachment findAttachmentByCid(MapiMessage message, String cid) {
        for (int i = 0; i < message.getAttachments().size(); i++) {
            MapiAttachment attachment = message.getAttachments().get_Item(i);
            String attachmentContentId = extractContentId(attachment);
            if (attachmentContentId != null && attachmentContentId.equals(cid)) {
                return attachment;
            }
        }
        return null;
    }

    private String extractContentId(MapiAttachment attachment) {
        // Here, we try to extract the Content-ID from the name or headers of the attachment
        String contentId = null;

        // Option 1: Extract from attachment name
        if (attachment.getLongFileName() != null && attachment.getLongFileName().contains("@")) {
            contentId = attachment.getLongFileName().replaceFirst("^<|>$", ""); // Strip < and > if present
        }

        // Option 2: If the Content-ID is stored in custom properties, fetch it accordingly
        if (contentId == null && attachment.getProperties() != null) {
            contentId = attachment.getProperties().get_Item(0x3712).getString(); // 0x3712 is PR_ATTACHMENT_CONTENT_ID
        }

        return contentId;
    }

    private String saveImageToTempFile(MapiAttachment attachment) {
        try {
            Path tempFile = Files.createTempFile(null, ".png");
            Files.write(tempFile, attachment.getBinaryData());
            return tempFile.toUri().toString();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private String generateHtmlContent(MapiMessage message) {
        StringBuilder html = new StringBuilder();
        html.append("<html><body>");

        // From
        html.append("<b>From:</b> ").append(message.getSenderEmailAddress()).append("<br>");

        // Sent
        html.append("<b>Sent:</b> ").append(message.getDeliveryTime().toString()).append("<br>");

        // To
        html.append("<b>To:</b> ").append(extractRecipients(message.getRecipients())).append("<br>");

        // Subject
        html.append("<b>Subject:</b> ").append(message.getSubject()).append("<br><br>");

        // Body content
        String bodyContent = message.getBody();
        if (message.getBodyType() == BodyContentType.Html) {
            bodyContent = Jsoup.parse(bodyContent).text(); // Convert HTML to plain text
        }

        html.append("<p>").append(bodyContent.replaceAll("\n", "<br>")).append("</p>");
        html.append("</body></html>");

        return html.toString();
    }


    private String convertMessageToHTML(MapiMessage message) {
        StringBuilder html = new StringBuilder();

        html.append("<html><body>");
        html.append("<p><strong>From:</strong> ").append(message.getSenderEmailAddress()).append("</p>");
        html.append("<p><strong>Sent:</strong> ").append(message.getDeliveryTime().toString()).append("</p>");
        html.append("<p><strong>To:</strong> ").append(extractRecipients(message.getRecipients())).append("</p>");
        html.append("<p><strong>Subject:</strong> ").append(message.getSubject()).append("</p>");
        html.append("<hr>");

        String bodyContent = message.getBody();
        if (message.getBodyType() == BodyContentType.Html) {
            bodyContent = Jsoup.parse(bodyContent).html();
        } else {
            bodyContent = bodyContent.replace("\n", "<br>");
        }

        html.append(bodyContent);
        html.append("</body></html>");

        return html.toString();
    }




    private String createHtmlContentFromMessage(MapiMessage message) {
        StringBuilder html = new StringBuilder();
        html.append("<html><head>");
        html.append("<style>")
                .append("body { font-family: 'Arial', sans-serif; }")
                .append("h1 { color: navy; }")
                .append("</style>")
                .append("</head><body>");

        html.append("<p><strong>From:</strong> ").append(message.getSenderEmailAddress()).append("</p>");
        html.append("<p><strong>Sent:</strong> ").append(message.getDeliveryTime().toString()).append("</p>");
        html.append("<p><strong>To:</strong> ").append(extractRecipients(message.getRecipients())).append("</p>");
        html.append("<p><strong>Subject:</strong> ").append(message.getSubject()).append("</p>");
        html.append("<hr/>");

        String bodyContent = message.getBody();
        if (message.getBodyType() == BodyContentType.Html) {
            html.append(bodyContent); // Directly use HTML content
        } else {
            html.append("<pre>").append(bodyContent).append("</pre>"); // Preserve plain text formatting
        }

        html.append("</body></html>");
        return html.toString();
    }

    // Clean and preprocess the HTML content
    private String cleanHtml(String htmlContent) {
        Document doc = Jsoup.parse(htmlContent);
        doc.outputSettings().prettyPrint(true);
        return Jsoup.clean(doc.html(), Safelist.relaxed());
    }

    // Generate the email content in HTML format
    private String generateEmailHtml(MapiMessage message, String bodyContent) {
        StringBuilder html = new StringBuilder();
        html.append("<html><body>");
        html.append("<strong>From:</strong> ").append(message.getSenderEmailAddress()).append("<br/>");
        html.append("<strong>Sent:</strong> ").append(message.getDeliveryTime()).append("<br/>");
        html.append("<strong>To:</strong> ").append(extractRecipients(message.getRecipients())).append("<br/>");
        html.append("<strong>Subject:</strong> ").append(message.getSubject()).append("<br/><br/>");
        html.append(bodyContent);
        html.append("</body></html>");
        return html.toString();
    }

    private String filterUnsupportedCharacters(String text) {
        // If needed, implement a method to filter unsupported characters
        return text;
    }




    private String filterUnsupportedCharacters(String text, PDFont font) throws IOException {
        StringBuilder filteredText = new StringBuilder();
        for (char c : text.toCharArray()) {
            if (font.getStringWidth(String.valueOf(c)) > 0) {
                filteredText.append(c);
            } else {
                filteredText.append('?'); // Placeholder for unsupported characters
            }
        }
        return filteredText.toString();
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

        private String cleanFileName (String fileName){
            return (fileName != null ? fileName : "Unnamed").replaceAll("[\\\\/:*?\"<>|]", "_").trim();
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
    }