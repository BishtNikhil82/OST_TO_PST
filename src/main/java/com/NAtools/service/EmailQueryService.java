package com.NAtools.service;


import com.NAtools.db.QueryBuilder;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public class EmailQueryService {
    private final Connection conn;

    public EmailQueryService(Connection conn) {
        this.conn = conn;
    }
    //Convert use Date Utils formatDateToISO to convert StartDate and EndDate in ISO format
    public List<String> getEmailsByDateRangeAndFolder(String folderName, String startDate, String endDate) {
        QueryBuilder qb = new QueryBuilder(conn);
        ResultSet rs = null;
        List<String> emails = new ArrayList<>();
        try {
            rs = qb.selectEmails()
                    .joinFolders()
                    .filterByFolder(folderName)
                    .filterByDateRange(startDate, endDate)
                    .orderByDateDesc()
                    .executeQuery();

            while (rs.next()) {
                // Collect email subjects for example
                emails.add(rs.getString("subject"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            closeResultSet(rs);
        }
        return emails;
    }

    public List<String> getEmailsBySender(String senderEmail) {
        QueryBuilder qb = new QueryBuilder(conn);
        ResultSet rs = null;
        List<String> emails = new ArrayList<>();
        try {
            rs = qb.selectEmails()
                    .filterBySender(senderEmail)
                    .executeQuery();

            while (rs.next()) {
                emails.add(rs.getString("subject"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            closeResultSet(rs);
        }
        return emails;
    }

    public List<String> getEmailsBySubjectKeyword(String keyword) {
        QueryBuilder qb = new QueryBuilder(conn);
        ResultSet rs = null;
        List<String> emails = new ArrayList<>();
        try {
            rs = qb.selectEmails()
                    .filterBySubject(keyword)
                    .executeQuery();

            while (rs.next()) {
                emails.add(rs.getString("subject"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            closeResultSet(rs);
        }
        return emails;
    }

    private void closeResultSet(ResultSet rs) {
        if (rs != null) {
            try {
                rs.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }
}
