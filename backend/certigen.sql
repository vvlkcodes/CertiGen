CREATE DATABASE IF NOT EXISTS certigen;
USE certigen;

CREATE TABLE IF NOT EXISTS employees (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    email VARCHAR(150) NOT NULL UNIQUE,
    password VARCHAR(255) NOT NULL,
    phone VARCHAR(30),
    role VARCHAR(30) DEFAULT 'employee'
);

CREATE TABLE IF NOT EXISTS certificates (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(150) NOT NULL,
    reg_no VARCHAR(100) NOT NULL UNIQUE,
    course_name VARCHAR(150) NOT NULL,
    course_type VARCHAR(80) NOT NULL,
    score VARCHAR(30) NOT NULL,
    passing_marks VARCHAR(30),
    total_marks VARCHAR(30),
    roll_number VARCHAR(100),
    he_she VARCHAR(20),
    start_date DATE NOT NULL,
    end_date DATE NOT NULL,
    email VARCHAR(150),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

INSERT INTO employees (name, email, password, phone, role)
VALUES ('Admin User', 'admin@certigen.local', 'admin123', '0000000000', 'admin')
ON DUPLICATE KEY UPDATE
name = VALUES(name),
password = VALUES(password),
phone = VALUES(phone),
role = VALUES(role);

INSERT INTO employees (name, email, password, phone, role)
VALUES ('Employee User', 'employee@certigen.local', 'employee123', '1111111111', 'employee')
ON DUPLICATE KEY UPDATE
name = VALUES(name),
password = VALUES(password),
phone = VALUES(phone),
role = VALUES(role);
