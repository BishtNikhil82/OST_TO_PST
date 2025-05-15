package com.NAtools.db;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;

public class QueryBuilder {
    private final Connection conn;
    private final StringBuilder query;
    private final List<Object> parameters;
    private boolean whereClauseAdded = false;

    public QueryBuilder(Connection conn) {
        this.conn = conn;
        this.query = new StringBuilder();
        this.parameters = new ArrayList<>();
    }

    public QueryBuilder selectEmails() {
        query.append("SELECT m.* FROM Messages m ");
        return this;
    }

    public QueryBuilder joinFolders() {
        query.append("JOIN Folders f ON m.folder_id = f.id ");
        return this;
    }

    public QueryBuilder filterByFolder(String folderName) {
        addWhereClauseIfNeeded();
        query.append("f.name = ? ");
        parameters.add(folderName);
        return this;
    }

    public QueryBuilder filterByDateRange(String startDate, String endDate) {
        addWhereClauseIfNeeded();
        query.append("m.received_date BETWEEN ? AND ? ");
        parameters.add(startDate);
        parameters.add(endDate);
        return this;
    }

    public QueryBuilder filterBySender(String senderEmail) {
        addWhereClauseIfNeeded();
        query.append("m.sender_email = ? ");
        parameters.add(senderEmail);
        return this;
    }

    public QueryBuilder filterBySubject(String keyword) {
        addWhereClauseIfNeeded();
        query.append("m.subject LIKE ? ");
        parameters.add("%" + keyword + "%");
        return this;
    }

    public QueryBuilder orderByDateDesc() {
        query.append("ORDER BY m.received_date DESC ");
        return this;
    }

    public ResultSet executeQuery() {
        ResultSet rs = null;
        try {

            PreparedStatement stmt = conn.prepareStatement(query.toString());
            for (int i = 0; i < parameters.size(); i++) {
                stmt.setObject(i + 1, parameters.get(i));
            }
            System.out.println("Executing query: " + query.toString());
            for (int i = 0; i < parameters.size(); i++) {
                System.out.println("Parameter " + (i + 1) + ": " + parameters.get(i));
            }
            rs = stmt.executeQuery();
        } catch (SQLException e) {
            e.printStackTrace();
            System.err.println("Error executing query: " + e.getMessage());
        }
        return rs;
    }


    private void addWhereClauseIfNeeded() {
        if (!whereClauseAdded) {
            query.append("WHERE ");
            whereClauseAdded = true;
        } else {
            query.append("AND ");
        }
    }
}
