package com.NAtools.util;

import org.docx4j.org.xhtmlrenderer.swing.NaiveUserAgent;
import org.docx4j.org.xhtmlrenderer.resource.ImageResource;
import org.docx4j.org.xhtmlrenderer.extend.FSImage;
import org.docx4j.org.xhtmlrenderer.pdf.ITextFSImage;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.net.URL;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import com.lowagie.text.Image;
import java.awt.Graphics2D;
import java.awt.Color;

public class CustomUserAgent extends NaiveUserAgent {
    private static final Logger logger = Logger.getLogger(CustomUserAgent.class.getName());

    // Class-level variables for default content ID and path
    private static String defaultContentIdPath;
    private static ImageResource defaultImageResource;
    private static final String DEFAULT_CONTENT_ID = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAHElEQVQoU2Nk+M+AFzEwMDAwUjA0NEASABoGAQct+gABAAAAAElFTkSuQmCC";

    // Static initializer block to create default content ID path and image resource
    static {
        initializeDefaultContentId();
    }

    public static String getDefaultContentIdPath() {
        return defaultContentIdPath;
    }
    public static ImageResource getDefaulImageResource() {
        return defaultImageResource;
    }

    @Override
    public ImageResource getImageResource(String uri) {
        String resolvedUrl = null;
        try {
            resolvedUrl = resolveURI(uri);
            logger.info(" ********** Calling getImageResource 1 *******");
            if (resolvedUrl != null && !resolvedUrl.equals(defaultContentIdPath)) {
                BufferedImage image = loadImage(resolvedUrl);
                if (image != null) {
                    logger.info(" ********** Calling getImageResource 2 *******");
                    FSImage fsImage = convertToFSImage(image);
                    return new ImageResource(uri, fsImage);
                } else {
                    logger.warning("ImageIO.read() returned null for URI: " + resolvedUrl + ". Using default placeholder image.");
                }
            }
        } catch (Exception e) {
            logger.warning("Failed to resolve or load image from URI: " + uri + " - " + e.getMessage());
        }

        // Always return the default image resource if the image is not found or URI cannot be resolved
        logger.info(" ********** Calling getImageResource 3  with resolve url = *******"+ resolvedUrl);
        logger.info(" ********** Calling getImageResource 3  Defaulturi = *******" + defaultImageResource.getImageUri());
        return defaultImageResource;

    }

    @Override
    public String resolveURI(String uri) {
        if (uri == null || uri.isEmpty()) {
            logger.warning("Received a null or empty URI. Using default content ID path.");
            return defaultContentIdPath; // Use default content ID path if URI is null or empty
        }

        try {
            URL url = new URL(uri);
            return url.toString();
        } catch (Exception e) {
            if (uri.startsWith("cid:")) {
                logger.warning("CID URI detected: " + uri + ". Handling CID content appropriately.");
                return handleCidUri(uri); // Handle CID URIs if needed
            } else {
                logger.warning("Invalid or malformed URI: " + uri + ". Using default content ID path.");
                return defaultContentIdPath; // Return default content ID path for invalid URIs
            }
        }
    }

    private BufferedImage loadImage(String url) {
        try {
            return ImageIO.read(new URL(url));
        } catch (Exception e) {
            logger.warning("Failed to load image from URL: " + url + " - " + e.getMessage());
            return null;
        }
    }

    private FSImage convertToFSImage(BufferedImage image) {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            ImageIO.write(image, "png", baos);
            byte[] imageBytes = baos.toByteArray();
            Image itextImage = Image.getInstance(imageBytes);
            return new ITextFSImage(itextImage);
        } catch (Exception e) {
            logger.warning("Failed to convert BufferedImage to FSImage - " + e.getMessage());
            return null;
        }
    }

    private String handleCidUri(String cidUri) {
        // Check if the URI is in the expected CID format
        if (cidUri == null || !cidUri.startsWith("cid:")) {
            logger.warning("Invalid CID URI format: " + cidUri);
            return defaultContentIdPath; // Return the default Content-ID path if the format is invalid
        }

        // Extract the Content-ID from the URI
        String contentId = cidUri.substring(4); // Remove "cid:" prefix

        try {
            // Assuming you have a predefined directory where attachments are stored
            String attachmentsDir = "path/to/attachments/directory"; // Change this to your actual attachments directory

            // Construct the local file path using the Content-ID
            File attachmentFile = new File(attachmentsDir, contentId);

            // Check if the file exists at the resolved path
            if (attachmentFile.exists()) {
                logger.info("Resolved CID URI to local file path: " + attachmentFile.getAbsolutePath());
                return attachmentFile.toURI().toURL().toString(); // Convert file path to URL
            } else {
                logger.warning("Attachment file not found for CID: " + cidUri);
            }
        } catch (Exception e) {
            logger.warning("Error while handling CID URI: " + cidUri + ". " + e.getMessage());
        }

        // Return the default content ID path if the file is not found or an error occurs
        return defaultContentIdPath;
    }

    // Static method to initialize the default content ID path and image resource
    private static void initializeDefaultContentId() {
        try {
            // Generate a default placeholder image (e.g., 10x10 light gray image)
            BufferedImage placeholderImage = new BufferedImage(10, 10, BufferedImage.TYPE_INT_RGB);
            Graphics2D g2d = placeholderImage.createGraphics();
            g2d.setColor(Color.LIGHT_GRAY);
            g2d.fillRect(0, 0, 10, 10);
            g2d.dispose();

            // Save the placeholder image to a temporary file
            File tempFile = File.createTempFile("default_image", ".png");
            ImageIO.write(placeholderImage, "png", tempFile);

            logger.info("Created default placeholder image at: " + tempFile.getAbsolutePath());
            defaultContentIdPath = tempFile.toURI().toURL().toString(); // Assign the URL of the placeholder image

            // Create the default ImageResource
            FSImage fsImage = convertToFSImageStatic(placeholderImage);
            defaultImageResource = new ImageResource(defaultContentIdPath, fsImage);
        } catch (Exception e) {
            logger.warning("Failed to create a default placeholder image: " + e.getMessage());
            defaultContentIdPath = DEFAULT_CONTENT_ID; // Assign a small base64-encoded transparent PNG as a last resort
            defaultImageResource = new ImageResource(defaultContentIdPath, null); // Default image resource as null
        }
    }

    // Static method to convert BufferedImage to FSImage
    private static FSImage convertToFSImageStatic(BufferedImage image) {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            ImageIO.write(image, "png", baos);
            byte[] imageBytes = baos.toByteArray();
            Image itextImage = Image.getInstance(imageBytes);
            return new ITextFSImage(itextImage);
        } catch (Exception e) {
            logger.warning("Failed to convert BufferedImage to FSImage - " + e.getMessage());
            return null;
        }
    }
}
