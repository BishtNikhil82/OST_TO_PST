package com.NAtools.Convertor;
import com.NAtools.config.LogManagerConfig;
import com.NAtools.util.CustomUserAgent;
import com.aspose.email.*;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.convert.out.html.AbstractHtmlExporter;
import org.docx4j.convert.out.html.HtmlExporterNG2;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.org.xhtmlrenderer.docx.DocxRenderer;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.xml.sax.InputSource;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.net.MalformedURLException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.logging.Logger;




public class OSTToDOCX {
    private static final Logger logger = Logger.getLogger(OSTToDOCX.class.getName());

    static {
        LogManagerConfig.configureLogger(logger);
    }

    private final PersonalStorage sourcePst;
    private final String outputDirectory;
    private final String sourcePstPath;
    private final boolean enableDuplicateCheck;
    private final Set<String> setDuplicacy = new HashSet<>();
    private final Set<String> setDupliccal = new HashSet<>();
    private final Set<String> setDuplictask = new HashSet<>();
    private final Set<String> setDupliccontact = new HashSet<>();
    private final String ostFileName;
    private final String defaultContentIdPath = "file:/D:/Gmail_OST_BKUP/default_path/default.png";

    public OSTToDOCX(PersonalStorage sourcePst, String outputDirectory, String sourcePstPath, boolean enableDuplicateCheck) {
        this.sourcePst = sourcePst;
        this.outputDirectory = outputDirectory;
        this.sourcePstPath = sourcePstPath;
        this.enableDuplicateCheck = enableDuplicateCheck;

        // Extract the OST file name without the extension
        this.ostFileName = new File(sourcePstPath).getName().replaceFirst("[.][^.]+$", "");
    }

    public void convertToDocx() {
        try {
            FolderInfo rootFolder = sourcePst.getRootFolder();
            processFolder(rootFolder, outputDirectory);
            logger.info("Conversion to DOCX format completed.");
        } catch (Exception e) {
            logger.severe("Error during DOCX conversion: " + e.getMessage());
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
                    MapiMessage message = sourcePst.extractMessage(messageInfo);

                    switch (message.getMessageClass()) {
                        case "IPM.Note":
                            saveMessageAsDOCX(message, folderPath);
                            break;
                        case "IPM.Contact":
                            //saveContactAsDocx(message, folderPath);
                            break;
                        case "IPM.Appointment":
                        case "IPM.Schedule.Meeting.Request":
                            //saveCalendarAsDocx(message, folderPath);
                            break;
                        case "IPM.Task":
                            //saveTaskAsDocx(message, folderPath);
                            break;
                        default:
                            //saveGenericMessageAsDocx(message, folderPath);
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

    public void saveMessageAsDOCX(MapiMessage message, String folderPath) {
        try {
            String subject = message.getSubject();
            if (subject == null || subject.isEmpty()) {
                subject = "Unnamed"; // Fallback if subject is null or empty
            }
            logger.info("Starting to process the message: " + subject);

            // Prepare the email headers
            StringBuilder headers = new StringBuilder();
            headers.append("<div style=\"font-family: Arial, sans-serif; font-size: 12px; margin-bottom: 20px;\">")
                    .append("<strong>From:</strong> ").append(resolveSender(message)).append("<br/>")
                    .append("<strong>Sent:</strong> ").append(message.getDeliveryTime().toString()).append("<br/>")
                    .append("<strong>To:</strong> ").append(extractRecipients(message.getRecipients())).append("<br/>")
                    .append("<strong>Subject:</strong> ").append(message.getSubject()).append("<br/>");

            if (!message.getAttachments().isEmpty()) {
                headers.append("<strong>Attachments:</strong> ");
                for (MapiAttachment attachment : message.getAttachments()) {
                    headers.append(attachment.getDisplayName()).append(", ");
                }
                headers.setLength(headers.length() - 2);
                headers.append("<br/>");
            }
            headers.append("</div><hr/>");

            // Convert the message body to HTML if necessary
            String htmlContent;
            if (message.getBodyType() == BodyContentType.Html) {
                htmlContent = message.getBodyHtml();
            } else if (message.getBodyType() == BodyContentType.PlainText || message.getBodyType() == BodyContentType.Rtf) {
                htmlContent = convertPlainTextToHtml(message.getBody());
            } else {
                logger.warning("Message has no content to process.");
                return;
            }

            htmlContent = cleanHtmlComments(htmlContent);
            logger.info("Extracted content for message: " + subject + ". Content length: " + htmlContent.length());

            // Process embedded images and attachments
            Map<String, String> cidToImageMap = new HashMap<>();
            File attachmentDir = new File(folderPath + "/Images");
            if (!attachmentDir.exists()) {
                attachmentDir.mkdirs();
            }

            cidToImageMap = saveAndMapCidAttachments(message, attachmentDir);

            // Replace CID references with properly encoded file URIs or use default if not found
            for (Map.Entry<String, String> entry : cidToImageMap.entrySet()) {
                String cid = entry.getKey();
                String filePath = entry.getValue();
                try {
                    if (filePath != null && !filePath.isEmpty()) {
                        String encodedFilePath = URLEncoder.encode(new File(filePath).getName(), StandardCharsets.UTF_8.toString());
                        String encodedUriPath = new File(filePath).getParentFile().toURI().toURL().toString() + encodedFilePath;
                        htmlContent = htmlContent.replace("cid:" + cid, encodedUriPath);
                    } else {
                        htmlContent = htmlContent.replace("cid:" + cid, defaultContentIdPath);
                    }
                } catch (Exception e) {
                    logger.warning("Error processing URL for image path: " + filePath + ". Using default path.");
                    htmlContent = htmlContent.replace("cid:" + cid, defaultContentIdPath);
                }
            }

            // Replace any remaining `cid:` references with the default path
            htmlContent = htmlContent.replaceAll("cid:[\\w-]+", defaultContentIdPath);

            // Convert to Document using Jsoup
            Document jsoupDoc = Jsoup.parse(htmlContent);
            jsoupDoc.outputSettings()
                    .escapeMode(org.jsoup.nodes.Entities.EscapeMode.xhtml)
                    .syntax(Document.OutputSettings.Syntax.xml);

            Element body = jsoupDoc.body();
            body.prepend(headers.toString());

            // Clean and simplify HTML content
            htmlContent = cleanHtmlContent(jsoupDoc);

            htmlContent = htmlContent.replaceAll("&nbsp;", "&#160;");

            // Save DOCX file
            String baseFileName = cleanFileName(subject);
            if (baseFileName.isEmpty()) {
                baseFileName = "Unnamed";
            }
            String fileName = baseFileName + ".docx";
            File file = new File(folderPath, fileName);
            int count = 1;

            while (file.exists()) {
                fileName = baseFileName + "_" + count + ".docx";
                file = new File(folderPath, fileName);
                count++;
            }

            // Initialize WordprocessingMLPackage and XHTMLImporter
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            XHTMLImporterImpl xhtmlImporter = new XHTMLImporterImpl(wordMLPackage);

            // Convert HTML content using Reader
            try (Reader htmlReader = new StringReader(htmlContent)) {
                List<Object> docxObjects = xhtmlImporter.convert(htmlReader, null);
                mainDocumentPart.getContent().addAll(docxObjects);
                wordMLPackage.save(file);
                logger.info("Saved email as DOCX: " + fileName);
            }
        } catch (Exception e) {
            logger.warning("General Error saving message as DOCX: " + e.getMessage());
            e.printStackTrace();
        }
    }









    // Function to clean HTML content
    private String cleanHtmlContent(Document doc) {
        // Remove Microsoft-specific tags and attributes
        doc.select("[style*='mso-']").remove(); // Remove all Microsoft-specific styles
        doc.select("o\\:*").remove(); // Remove Office-specific elements
        doc.select("v\\:*").remove(); // Remove VML-specific elements
        doc.select("xml").remove(); // Remove XML declarations
        doc.select("meta[http-equiv='X-UA-Compatible']").remove(); // Remove X-UA-Compatible meta tags

        // Remove script and style tags for security and cleanliness
        doc.select("script").remove();
        doc.select("style").remove();

        // Remove any unwanted inline styles that might interfere with DOCX rendering
        doc.select("[style]").removeAttr("style");

        // Remove unwanted tags such as form elements, iframes, and objects
        doc.select("iframe").remove();
        doc.select("embed").remove();
        doc.select("object").remove();
        doc.select("form").remove();
        doc.select("input").remove();
        doc.select("textarea").remove();
        doc.select("button").remove();
        doc.select("select").remove();
        doc.select("option").remove();

        // Remove empty elements like empty divs, spans, paragraphs, etc.
        doc.select("div:empty").remove();
        doc.select("span:empty").remove();
        doc.select("p:empty").remove();
        doc.select("strong:empty").remove();
        doc.select("em:empty").remove();

        // Remove unnecessary attributes
        doc.select("*").removeAttr("id");
        doc.select("*").removeAttr("class");
        doc.select("*").removeAttr("onclick");
        doc.select("*").removeAttr("onload");
        doc.select("*").removeAttr("onerror");

        // Remove comments
//        doc.select("*").forEach(element -> {
//            element.childNodes().removeIf(node -> node.nodeName().equals("#comment"));
//        });

        // Normalize and clean up
        doc.outputSettings().escapeMode(org.jsoup.nodes.Entities.EscapeMode.xhtml);
        doc.outputSettings().syntax(Document.OutputSettings.Syntax.xml);
        doc.outputSettings().prettyPrint(true); // Optional: prettify the output for readability

        return doc.html();
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

    // Other save methods for Contacts, Calendar, Task, and Generic messages (similar to the MSG methods) ...
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
    // Methods to handle duplicates...
    private boolean isDuplicateMessage(MapiMessage message) {
        String key = message.getSubject() + message.getBody();
        if (setDuplicacy.contains(key)) {
            return true;
        } else {
            setDuplicacy.add(key);
            return false;
        }
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
    private String cleanHtmlComments(String htmlContent) {
        return htmlContent.replaceAll("<!--.*?-->", ""); // Basic example, can be adjusted for more complex cases
    }
}
