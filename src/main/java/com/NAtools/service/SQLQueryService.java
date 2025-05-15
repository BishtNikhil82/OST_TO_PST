package com.NAtools.service;

import com.NAtools.model.Attachment;
import com.NAtools.model.Folder;
import com.NAtools.model.Message;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

public class SQLQueryService {

    private static final Logger logger = Logger.getLogger(SQLQueryService.class.getName());

    public List<Folder> getRootFolders(Connection connection) {
        List<Folder> folders = new ArrayList<>();
        String query = "SELECT id, name FROM Folders WHERE parent_id IS NULL";
        try (PreparedStatement stmt = connection.prepareStatement(query);
             ResultSet rs = stmt.executeQuery()) {

            while (rs.next()) {
                int id = rs.getInt("id");
                String name = rs.getString("name");
                if (name != null) {
                    Folder folder = new Folder();
                    folder.setId(id);
                    folder.setName(name);
                    folders.add(folder);
                    logger.info("Fetched root folder: " + name);
                } else {
                    logger.warning("Folder name is null for ID: " + id);
                }
            }
        } catch (SQLException e) {
            logger.severe("Error fetching root folders: " + e.getMessage());
            e.printStackTrace();
        }
        return folders;
    }

    public List<Folder> getSubFolders(Connection connection, int parentId) {
        List<Folder> folders = new ArrayList<>();
        String query = "SELECT id, name FROM Folders WHERE parent_id = ?";
        try (PreparedStatement stmt = connection.prepareStatement(query)) {
            stmt.setInt(1, parentId);
            try (ResultSet rs = stmt.executeQuery()) {

                while (rs.next()) {
                    int id = rs.getInt("id");
                    String name = rs.getString("name");
                    if (name != null) {
                        Folder folder = new Folder();
                        folder.setId(id);
                        folder.setName(name);
                        folders.add(folder);
                        logger.info("Fetched subfolder: " + name);
                    } else {
                        logger.warning("Folder name is null for ID: " + id);
                    }
                }
            }
        } catch (SQLException e) {
            logger.severe("Error fetching subfolders: " + e.getMessage());
            e.printStackTrace();
        }
        return folders;
    }

    public List<Message> getMessagesForFolder(Connection connection, int folderId) {
        List<Message> messages = new ArrayList<>();
        String query = "SELECT * FROM Messages WHERE folder_id = ?";
        try (PreparedStatement stmt = connection.prepareStatement(query)) {
            stmt.setInt(1, folderId);
            try (ResultSet rs = stmt.executeQuery()) {

                while (rs.next()) {
                    int id = rs.getInt("id");
                    String subject = rs.getString("subject");
                    String body = rs.getString("body");
                    String body_format = rs.getString("body_format");
                    String senderEmail = rs.getString("sender_email");
                    String recipients = rs.getString("recipients");
                    String receivedDate = rs.getString("received_date");

                    // Use default values if any field is null
                    if (subject == null) subject = "(No Subject)";
                    if (body == null) body = "(No Body)";
                    if (senderEmail == null) senderEmail = "(No Sender)";
                    if (recipients == null) recipients = "(No Recipients)";
                    if (receivedDate == null) receivedDate = "1970-01-01 00:00:00"; // Default to Unix epoch

                    Message message = new Message();
                    message.setId(id);
                    message.setSubject(subject);
                    message.setBody(body);
                    message.setBodyFormat(body_format);
                    message.setSenderEmail(senderEmail);
                    message.setRecipients(recipients);
                    message.setReceivedDate(receivedDate);

                    messages.add(message);
                    logger.info("Fetched message with subject: " + subject);
                }
            }
        } catch (SQLException e) {
            logger.severe("Error fetching messages for folder: " + e.getMessage());
            e.printStackTrace();
        }
        return messages;
    }
    public List<Attachment> getAttachmentsForMessage(Connection connection,int messageId) {
        List<Attachment> attachments = new ArrayList<>();
        String query = "SELECT * FROM Attachments WHERE message_id = ?";
        try (PreparedStatement stmt = connection.prepareStatement(query)) {
            stmt.setInt(1, messageId);
            try (ResultSet rs = stmt.executeQuery()) {

                while (rs.next()) {
                    int id = rs.getInt("id");
                    String fileName = rs.getString("file_name");
                    String filePath = rs.getString("file_path");

                    Attachment attachment = new Attachment();
                    attachment.setId(id);
                    attachment.setFileName(fileName);
                    //attachment.setFilePath(filePath);
                    attachments.add(attachment);
                    logger.info("Fetched attachment with file name: " + fileName + " for message ID: " + messageId);
                }
            }
        } catch (SQLException e) {
            logger.severe("Error fetching attachments for message: " + e.getMessage());
            e.printStackTrace();
        }
        return attachments;
    }

}
