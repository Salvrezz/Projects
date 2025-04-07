create database EXCO

begin tran
-- Create table for departments
CREATE TABLE departments (
    dept_id INT PRIMARY KEY,
    department_name VARCHAR(50) UNIQUE NOT NULL
)
COMMIT

BEGIN TRAN
-- Create table for employees
CREATE TABLE employees (
    emp_id INT PRIMARY KEY,
    full_name VARCHAR(100) NOT NULL,
    age INT CHECK (age >= 18 AND age <= 65),
    gender varchar(100) NOT NULL,
    email VARCHAR(100) UNIQUE NOT NULL
)
COMMIT

BEGIN TRAN
-- Create table for employee job details
CREATE TABLE employee_details (
    detail_id INT PRIMARY KEY,
    emp_id INT,
    dept_id INT,
    salary_usd DECIMAL(10, 2),
    doj DATE NOT NULL,
    status varchar(100) DEFAULT 'Active',
    FOREIGN KEY (emp_id) REFERENCES employees(emp_id),
    FOREIGN KEY (dept_id) REFERENCES departments(dept_id)
)
COMMIT

BEGIN TRAN
-- Insert department data
INSERT INTO departments (dept_id, department_name) VALUES
(1, 'Finance'),
(2, 'Engineering'),
(3, 'Human Resources'),
(4, 'HR')
COMMIT

BEGIN TRAN
-- Insert employees
INSERT INTO employees (emp_id, full_name, age, gender, email) VALUES
(1, 'John Doe', 29, 'Male', 'johndoe@email.com'),
(2, 'Jane Smith', 33, 'Female', 'janesmith@email.com'),
(3, 'Ahmed Musa', NULL, 'Male', 'ahmedmusa@email.com'),
(4, 'Grace Lee', 26, 'Female', 'gracelee@email.com'),
(5, 'Mark O’Brian', 41, 'Male', 'markobrian@email.com'),
(6, 'Sarah Kim', 38, 'Female', 'sarahkim@email.com'),
(7, 'Samuel Johnson', 45, 'Male', 'samuel@email.com'),
(8, 'Chinaza Okafor', 29, 'Female', 'chinaza@email.com')
COMMIT

SELECT * FROM employees

BEGIN TRAN
-- Insert employee details
INSERT INTO employee_details (detail_id, emp_id, dept_id, salary_usd, doj, status) VALUES
(1, 1, 1, 75000.00, '2021-01-15', 'Active'),
(2, 2, 1, 80000.00, '2020-05-22', 'Active'),
(3, 3, 2, 90000.00, '2019-03-12', 'Active'),
(4, 4, 3, 67000.00, '2022-07-01', 'Inactive'),
(5, 5, 2, 93000.00, '2018-11-04', 'Active'),
(6, 6, 4, 67000.00, '2021-08-19', 'Active'),
(7, 7, 1, 75000.00, '2017-10-09', 'Active'),
(8, 8, 2, NULL, '2020-01-03', 'Active')
commit
SELECT * FROM employee_details
SELECT * FROM departments
SELECT * FROM employees

begin tran
SELECT COUNT(*) AS total_employees
FROM employees;
commit

begin tran
SELECT e.full_name, ed.salary_usd
FROM employees as e
JOIN employee_details as ed ON e.emp_id = ed.emp_id
WHERE ed.salary_usd = (
    SELECT MAX(salary_usd)
    FROM employee_details
)
commit
SELECT * FROM employee_details
SELECT * FROM departments
SELECT * FROM employees

 begin tran
SELECT e.full_name, ed.doj
FROM employees AS e
JOIN employee_details ed ON e.emp_id = ed.emp_id
WHERE ed.doj = (
    SELECT MAX(doj) FROM employee_details
)
COMMIT

BEGIN TRAN
SELECT E.full_name,ED.salary_usd
FROM employees AS E
JOIN employee_details AS ED ON E.emp_id = ED.emp_id
WHERE ED.salary_usd IS NULL
COMMIT