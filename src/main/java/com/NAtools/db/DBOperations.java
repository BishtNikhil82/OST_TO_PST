package com.NAtools.db;

import com.NAtools.model.*;

import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public class DBOperations {

    // Method to insert a folder using an external connection
    public int insertFolder(Connection conn, String name, Integer parentId) throws SQLException {
        String sql = "INSERT INTO Folders(name, parent_id) VALUES(?, ?)";
        int generatedId = -1;

        try (PreparedStatement pstmt = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            pstmt.setString(1, name);
            if (parentId != null) {
                pstmt.setInt(2, parentId);
            } else {
                pstmt.setNull(2, java.sql.Types.INTEGER);
            }
            pstmt.executeUpdate();

            // Retrieve the generated ID
            ResultSet rs = pstmt.getGeneratedKeys();
            if (rs.next()) {
                generatedId = rs.getInt(1);
            }
        }
        return generatedId;
    }

    // Method to update a folder using an external connection
    public void updateFolder(Connection conn, int id, String name, Integer parentId) throws SQLException {
        String sql = "UPDATE Folders SET name = ?, parent_id = ? WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setString(1, name);
            if (parentId != null) {
                pstmt.setInt(2, parentId);
            } else {
                pstmt.setNull(2, java.sql.Types.INTEGER);
            }
            pstmt.setInt(3, id);
            pstmt.executeUpdate();
        }
    }

    // Method to delete a folder using an external connection
    public void deleteFolder(Connection conn, int id) throws SQLException {
        String sql = "DELETE FROM Folders WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            pstmt.executeUpdate();
        }
    }

    // Method to get a folder by ID using an external connection
    public Folder getFolderById(Connection conn, int id) throws SQLException {
        String sql = "SELECT * FROM Folders WHERE id = ?";
        Folder folder = null;
        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            ResultSet rs = pstmt.executeQuery();
            if (rs.next()) {
                folder = new Folder();
                folder.setId(rs.getInt("id"));
                folder.setName(rs.getString("name"));
                folder.setParentId(rs.getInt("parent_id"));
            }
        }
        return folder;
    }

    // Method to get all folders using an external connection
    public List<Folder> getAllFolders(Connection conn) throws SQLException {
        List<Folder> folders = new ArrayList<>();
        String sql = "SELECT * FROM Folders";

        try (Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {

            while (rs.next()) {
                Folder folder = new Folder();
                folder.setId(rs.getInt("id"));
                folder.setName(rs.getString("name"));
                folder.setParentId(rs.getInt("parent_id"));
                folders.add(folder);
            }
        }
        return folders;
    }

    // Method to insert a message using an external connection
// Method to insert a message using an external connection and return the generated message ID
// Updated insertMessage method
    public int insertMessage(Connection conn, int folderId, String subject, String senderEmail, String body, String receivedDate, String recipients, String bodyFormat) throws SQLException {
        String sql = "INSERT INTO Messages(folder_id, subject, sender_email, body, body_format, received_date, recipients) VALUES(?, ?, ?, ?, ?, ?, ?)";
        int generatedId = -1;

        try (PreparedStatement pstmt = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            pstmt.setInt(1, folderId);
            pstmt.setString(2, subject);
            pstmt.setString(3, senderEmail);
            pstmt.setString(4, body);
            pstmt.setString(5, bodyFormat);
            pstmt.setString(6, receivedDate);
            pstmt.setString(7, recipients);

            pstmt.executeUpdate();

            // Retrieve the generated ID
            ResultSet rs = pstmt.getGeneratedKeys();
            if (rs.next()) {
                generatedId = rs.getInt(1);
            }
        }
        return generatedId;
    }



    // Method to update a message using an external connection
    public void updateMessage(Connection conn, int id, int folderId, String subject, String senderEmail, String body, Timestamp receivedDate) throws SQLException {
        String sql = "UPDATE Messages SET folder_id = ?, subject = ?, sender_email = ?, body = ?, received_date = ? WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, folderId);
            pstmt.setString(2, subject);
            pstmt.setString(3, senderEmail);
            pstmt.setString(4, body);
            pstmt.setTimestamp(5, receivedDate);
            pstmt.setInt(6, id);
            pstmt.executeUpdate();
        }
    }

    // Method to delete a message using an external connection
    public void deleteMessage(Connection conn, int id) throws SQLException {
        String sql = "DELETE FROM Messages WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            pstmt.executeUpdate();
        }
    }

    // Method to get a message by ID using an external connection
    public Message getMessageById(Connection conn, int id) throws SQLException {
        String sql = "SELECT * FROM Messages WHERE id = ?";
        Message message = null;
        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            ResultSet rs = pstmt.executeQuery();
            if (rs.next()) {
                message = new Message();
                message.setId(rs.getInt("id"));
                message.setFolderId(rs.getInt("folder_id"));
                message.setSubject(rs.getString("subject"));
                message.setSenderEmail(rs.getString("sender_email"));
                message.setBody(rs.getString("body"));
                message.setReceivedDate("received_date");
                message.setRecipients(rs.getString("recipients"));// Use getTimestamp for DATETIME fields
            }
        }
        return message;
    }

    // Method to insert a calendar event using an external connection
    public void insertCalendarEvent(Connection conn, int folderId, String location, Timestamp startDate, Timestamp endDate, String description) throws SQLException {
        String sql = "INSERT INTO CalendarEvents(folder_id, location, start_date, end_date, description) VALUES(?, ?, ?, ?, ?)";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, folderId);
            pstmt.setString(2, location);
            pstmt.setTimestamp(3, startDate); // Use Timestamp for DATETIME fields
            pstmt.setTimestamp(4, endDate);   // Use Timestamp for DATETIME fields
            pstmt.setString(5, description);
            pstmt.executeUpdate();
        }
    }

    // Method to update a calendar event using an external connection
    public void updateCalendarEvent(Connection conn, int id, int folderId, String location, Timestamp startDate, Timestamp endDate, String description) throws SQLException {
        String sql = "UPDATE CalendarEvents SET folder_id = ?, location = ?, start_date = ?, end_date = ?, description = ? WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, folderId);
            pstmt.setString(2, location);
            pstmt.setTimestamp(3, startDate); // Use Timestamp for DATETIME fields
            pstmt.setTimestamp(4, endDate);   // Use Timestamp for DATETIME fields
            pstmt.setString(5, description);
            pstmt.setInt(6, id);
            pstmt.executeUpdate();
        }
    }

    // Method to delete a calendar event using an external connection
    public void deleteCalendarEvent(Connection conn, int id) throws SQLException {
        String sql = "DELETE FROM CalendarEvents WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            pstmt.executeUpdate();
        }
    }

    // Method to get a calendar event by ID using an external connection
    public CalendarEvent getCalendarEventById(Connection conn, int id) throws SQLException {
        String sql = "SELECT * FROM CalendarEvents WHERE id = ?";
        CalendarEvent event = null;
        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            ResultSet rs = pstmt.executeQuery();
            if (rs.next()) {
                event = new CalendarEvent();
                event.setId(rs.getInt("id"));
                event.setFolderId(rs.getInt("folder_id"));
                event.setLocation(rs.getString("location"));
                event.setStartDate(rs.getTimestamp("start_date")); // Use getTimestamp for DATETIME fields
                event.setEndDate(rs.getTimestamp("end_date"));     // Use getTimestamp for DATETIME fields
                event.setDescription(rs.getString("description"));
            }
        }
        return event;
    }

    // Method to insert a contact using an external connection
    public void insertContact(Connection conn, int folderId, String displayName, String notes, String email) throws SQLException {
        String sql = "INSERT INTO Contacts(folder_id, display_name, notes, email) VALUES(?, ?, ?, ?)";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, folderId);
            pstmt.setString(2, displayName);
            pstmt.setString(3, notes);
            pstmt.setString(4, email);
            pstmt.executeUpdate();
        }
    }

    // Method to update a contact using an external connection
    public void updateContact(Connection conn, int id, int folderId, String displayName, String notes, String email) throws SQLException {
        String sql = "UPDATE Contacts SET folder_id = ?, display_name = ?, notes = ?, email = ? WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, folderId);
            pstmt.setString(2, displayName);
            pstmt.setString(3, notes);
            pstmt.setString(4, email);
            pstmt.setInt(5, id);
            pstmt.executeUpdate();
        }
    }

    // Method to delete a contact using an external connection
    public void deleteContact(Connection conn, int id) throws SQLException {
        String sql = "DELETE FROM Contacts WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            pstmt.executeUpdate();
        }
    }

    // Method to get a contact by ID using an external connection
    public Contact getContactById(Connection conn, int id) throws SQLException {
        String sql = "SELECT * FROM Contacts WHERE id = ?";
        Contact contact = null;
        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            ResultSet rs = pstmt.executeQuery();
            if (rs.next()) {
                contact = new Contact();
                contact.setId(rs.getInt("id"));
                contact.setFolderId(rs.getInt("folder_id"));
                contact.setDisplayName(rs.getString("display_name"));
                contact.setNotes(rs.getString("notes"));
                contact.setEmail(rs.getString("email"));
            }
        }
        return contact;
    }

    // Method to insert a task using an external connection
    public void insertTask(Connection conn, int folderId, String subject, String body, Timestamp dueDate) throws SQLException {
        String sql = "INSERT INTO Tasks(folder_id, subject, body, due_date) VALUES(?, ?, ?, ?)";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, folderId);
            pstmt.setString(2, subject);
            pstmt.setString(3, body);
            pstmt.setTimestamp(4, dueDate); // Use Timestamp for DATETIME fields
            pstmt.executeUpdate();
        }
    }

    // Method to update a task using an external connection
    public void updateTask(Connection conn, int id, int folderId, String subject, String body, Timestamp dueDate) throws SQLException {
        String sql = "UPDATE Tasks SET folder_id = ?, subject = ?, body = ?, due_date = ? WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, folderId);
            pstmt.setString(2, subject);
            pstmt.setString(3, body);
            pstmt.setTimestamp(4, dueDate); // Use Timestamp for DATETIME fields
            pstmt.setInt(5, id);
            pstmt.executeUpdate();
        }
    }

    // Method to delete a task using an external connection
    public void deleteTask(Connection conn, int id) throws SQLException {
        String sql = "DELETE FROM Tasks WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            pstmt.executeUpdate();
        }
    }

    // Method to get a task by ID using an external connection
    public Task getTaskById(Connection conn, int id) throws SQLException {
        String sql = "SELECT * FROM Tasks WHERE id = ?";
        Task task = null;
        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            ResultSet rs = pstmt.executeQuery();
            if (rs.next()) {
                task = new Task();
                task.setId(rs.getInt("id"));
                task.setFolderId(rs.getInt("folder_id"));
                task.setSubject(rs.getString("subject"));
                task.setBody(rs.getString("body"));
                task.setDueDate(rs.getTimestamp("due_date")); // Using Timestamp for dueDate
            }
        }
        return task;
    }


    // Method to insert an attachment using an external connection
    public void insertAttachment(Connection conn, int messageId, String fileName, byte[] fileData) throws SQLException {
        String sql = "INSERT INTO Attachments(message_id, file_name, file_data) VALUES(?, ?, ?)";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, messageId);
            pstmt.setString(2, fileName);
            pstmt.setBytes(3, fileData);
            pstmt.executeUpdate();
        }
    }

    // Method to update an attachment using an external connection
    public void updateAttachment(Connection conn, int id, int messageId, String fileName, byte[] fileData) throws SQLException {
        String sql = "UPDATE Attachments SET message_id = ?, file_name = ?, file_data = ? WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, messageId);
            pstmt.setString(2, fileName);
            pstmt.setBytes(3, fileData);
            pstmt.setInt(4, id);
            pstmt.executeUpdate();
        }
    }

    // Method to delete an attachment using an external connection
    public void deleteAttachment(Connection conn, int id) throws SQLException {
        String sql = "DELETE FROM Attachments WHERE id = ?";

        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            pstmt.executeUpdate();
        }
    }

    // Method to get an attachment by ID using an external connection
    public Attachment getAttachmentById(Connection conn, int id) throws SQLException {
        String sql = "SELECT * FROM Attachments WHERE id = ?";
        Attachment attachment = null;
        try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setInt(1, id);
            try (ResultSet rs = pstmt.executeQuery()) {
                if (rs.next()) {
                    attachment = new Attachment();
                    attachment.setId(rs.getInt("id"));
                    attachment.setMessageId(rs.getInt("message_id"));
                    attachment.setFileName(rs.getString("file_name"));
                    attachment.setFileData(rs.getBytes("file_data"));
                }
            }
        }
        return attachment;
    }
}
