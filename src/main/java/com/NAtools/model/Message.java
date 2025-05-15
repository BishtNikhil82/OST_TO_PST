package com.NAtools.model;

import java.util.Date;

public class Message {
    private int id;
    private int folderId;
    private String subject;
    private String senderEmail;
    private String body;
    private String receivedDate; // Use Date type for receivedDate
    private String recipients;
    private String bodyFormat;
    // Getters and setters
    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public int getFolderId() {
        return folderId;
    }

    public void setFolderId(int folderId) {
        this.folderId = folderId;
    }

    public String getSubject() {
        return subject;
    }

    public void setSubject(String subject) {
        this.subject = subject;
    }

    public String getSenderEmail() {
        return senderEmail;
    }

    public void setSenderEmail(String senderEmail) {
        this.senderEmail = senderEmail;
    }

    public String getBody() {
        return body;
    }

    public void setBody(String body) {
        this.body = body;
    }
    public String getBodyFormat() { // Getter for body format
        return bodyFormat;
    }

    public void setBodyFormat(String bodyFormat) { // Setter for body format
        this.bodyFormat = bodyFormat;
    }

    public String getReceivedDate() {
        return receivedDate;
    }

    public void setReceivedDate(String receivedDate) {
        this.receivedDate = receivedDate;
    }
    public String getRecipients() {
        return recipients;
    }

    public void setRecipients(String recipients) {
        this.recipients = recipients;
    }
}
