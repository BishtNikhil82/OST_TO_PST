package com.NAtools.service;

import com.NAtools.model.Attachment;
import com.NAtools.model.EmailPreview;


import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public class EmailPreviewService {

    private final Connection conn;

    public EmailPreviewService(Connection conn) {
        this.conn = conn;
    }

    public EmailPreview getEmailPreview(int emailId, String folderName) {
        EmailPreview emailPreview = new EmailPreview();
        List<com.NAtools.model.Attachment> attachments = new ArrayList<>();
        boolean hasAttachments = false;

        try {
            // Query to get email details and attachments, filtered by folder name
            String query = "SELECT m.*, a.file_name as attachment_name, length(a.file_data) as attachment_size, f.name as folder_name " +
                    "FROM Messages m " +
                    "JOIN Folders f ON m.folder_id = f.id " +
                    "LEFT JOIN Attachments a ON m.id = a.message_id " +
                    "WHERE m.id = ? AND f.name = ?";
            PreparedStatement stmt = conn.prepareStatement(query);
            stmt.setInt(1, emailId);
            stmt.setString(2, folderName);

            ResultSet rs = stmt.executeQuery();

            while (rs.next()) {
                emailPreview.setEmailId(emailId);
                emailPreview.setFrom(rs.getString("sender_email"));
                emailPreview.setTo("");  // Replace with correct column name if available
                emailPreview.setSubject(rs.getString("subject"));
                emailPreview.setDate(rs.getString("received_date"));
                emailPreview.setFolderName(rs.getString("folder_name"));

                String attachmentName = rs.getString("attachment_name");
                if (attachmentName != null) {
                    Attachment attachment = new Attachment();
                    attachment.setName(attachmentName);
                    attachment.setSize(rs.getLong("attachment_size"));
                    attachments.add(attachment);
                    hasAttachments = true;
                }
            }

            emailPreview.setAttachments(attachments);
            emailPreview.setHasAttachments(hasAttachments);

        } catch (SQLException e) {
            e.printStackTrace();
        }

        return emailPreview;
    }
}
