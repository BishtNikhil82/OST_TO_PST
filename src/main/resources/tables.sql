
-- Create new tables
CREATE TABLE Folders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    parent_id INTEGER,
    FOREIGN KEY(parent_id) REFERENCES Folders(id)
);

CREATE TABLE Messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_id INTEGER,
    subject TEXT,
    sender_email TEXT,
    recipients TEXT,
    body TEXT,
    body_format TEXT,
    received_date TEXT,  -- Updated to DATETIME
    FOREIGN KEY(folder_id) REFERENCES Folders(id)
);

CREATE TABLE CalendarEvents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_id INTEGER,
    location TEXT,
    start_date DATETIME,  -- Updated to DATETIME
    end_date DATETIME,    -- Updated to DATETIME
    description TEXT,
    FOREIGN KEY(folder_id) REFERENCES Folders(id)
);

CREATE TABLE Contacts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_id INTEGER,
    display_name TEXT,
    notes TEXT,
    email TEXT,
    FOREIGN KEY(folder_id) REFERENCES Folders(id)
);

CREATE TABLE Tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_id INTEGER,
    subject TEXT,
    body TEXT,
    due_date DATETIME,  -- Updated to DATETIME
    FOREIGN KEY(folder_id) REFERENCES Folders(id)
);

CREATE TABLE Attachments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    message_id INTEGER,
    file_name TEXT,
    file_data BLOB,
    FOREIGN KEY(message_id) REFERENCES Messages(id)
);
