package com.NAtools.model;

public class Attachment {
    private int id;
    private int messageId;
    private String fileName;
    private String filePath;
    private byte[] fileData;
    private String name;
    private long size;
    // Getters and Setters
    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public int getMessageId() {
        return messageId;
    }

    public void setMessageId(int messageId) {
        this.messageId = messageId;
    }

    public String getFileName() {
        return fileName;
    }

    public void  setFilePath(String fileName) {
        this.fileName = fileName;
    }

    public void setFileName(String filePath) {
        this.filePath = filePath;
    }
    public byte[] getFileData() {
        return fileData;
    }

    public void setFileData(byte[] fileData) {
        this.fileData = fileData;
    }
    // Getters and Setters
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public long getSize() {
        return size;
    }

    public void setSize(long size) {
        this.size = size;
    }
}
