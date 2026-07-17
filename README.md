
# CertiGen

CertiGen is a Node.js + Express project for bulk certificate generation, ZIP download, email notification, employee/admin dashboards, and QR-based certificate verification.

## Project Structure

```text
CertiGen/
|-- backend/
|   |-- server.js
|   |-- package.json
|   |-- package-lock.json
|   |-- certigen.sql
|   |-- .env.example
|   |-- admin.local.example.json
|   |-- certificates/
|   |-- uploads/
|   `-- uploaded_templates/
|
|-- frontend/
|   |-- index.html
|   |-- admin.html
|   |-- employee-dashboard.html
|   |-- verify.html
|   |-- login.backup.html
|   |-- style.css
|   |-- styles.css
|   |-- image1.jpeg
|   |-- ParticipantList.xlsx
|   `-- templates/
|
|-- docs/
|   |-- Certi-Gen-APIs.txt
|   |-- Certi-Gen-Functions.txt
|   |-- Certi-Gen-Modules.txt
|   `-- System Architecture.txt
|
|-- Dockerfile
|-- DEPLOYMENT.md
|-- render.yaml
`-- README.md
```

## Features

- Bulk certificate generation
- ZIP download support
- Email notifications after certificate generation
- Admin and employee dashboards
- QR-based certificate verification
- MySQL support for local setup
- PostgreSQL support for free hosted deployment
- Docker deployment support

## Live Demo

Try the deployed application here:

https://certigen-cr0k.onrender.com

## Prerequisites

Install these first:

- Node.js 18+ or newer
- MySQL 8+ or compatible for local setup

For free hosted deployment, the app also supports PostgreSQL providers such as Neon.

## Setup

### 1. Go to the backend folder:

```powershell
cd backend
```

### 2. Install dependencies:

```powershell
npm install
```

### 3. Create the local MySQL database and tables:

```sql
SOURCE certigen.sql;
```

If your MySQL client does not support `SOURCE`, open `backend/certigen.sql` and run it manually.

The app also creates the required tables automatically when it starts, which is useful for hosted databases.

### 4. Set environment variables before starting the server

PowerShell example for local MySQL:

```powershell
$env:PORT="3000"
$env:DB_CLIENT="mysql"
$env:DB_HOST="localhost"
$env:DB_PORT="3306"
$env:DB_NAME="certigen"
$env:DB_USER="root"
$env:DB_PASSWORD="your_mysql_password"
$env:EMAIL_USER="your_gmail_address"
$env:EMAIL_PASS="your_gmail_app_password"

node server.js
```

You can leave `EMAIL_USER` and `EMAIL_PASS` empty if email functionality is not required.

### 5. Enable email notifications only

If database details are already configured and you only want email support:

```powershell
$env:EMAIL_USER="your_gmail_address"
$env:EMAIL_PASS="your_gmail_app_password"

node server.js
```

Use a Gmail App Password instead of your normal Gmail password.

### 6. Using default local MySQL setup

```powershell
$env:DB_HOST="localhost"
$env:DB_PORT="3306"
$env:DB_NAME="certigen"
$env:DB_USER="root"
$env:DB_PASSWORD="password"
$env:EMAIL_USER="your_gmail_address"
$env:EMAIL_PASS="your_gmail_app_password"

node server.js
```

These email variables are required for certificate email notifications.

### 7. Open the application

```text
http://localhost:3000
```

Do not use VS Code Live Server as the main application URL unless the backend server is also running with:

```powershell
node server.js
```

## Default Demo Accounts

These accounts are available for testing:

### Admin
- Email: `vkartheek007@gmail.com`
- Password: `admin@123`

### Employee
- Email: `employee@certigen.local`
- Password: `employee@123`

For public deployment, configure your own `ADMIN_EMAIL` and `ADMIN_PASSWORD` environment variables.

## GitHub Notes

The following folders and files are intentionally ignored because they contain generated files, dependencies, or private data:

- `backend/node_modules/`
- `backend/uploads/`
- `backend/uploaded_templates/`
- `backend/certificates/`
- `.env`
- `.env.local`
- `backend/.env`
- `backend/.env.local`
- `backend/admin.local.json`
- `Information.txt`
- `*.log`

Folder structures for:

- `backend/uploads/`
- `backend/uploaded_templates/`
- `backend/certificates/`

are preserved using `.gitkeep` files.

The following should never be pushed to GitHub:

- `Information.txt`
- `.env`
- `backend/admin.local.json`

because they may contain private credentials or configuration data.

## Run Commands

Run these commands from the `backend/` folder.

Start the application:

```powershell
npm start
```

Development mode:

```powershell
npm run dev
```

Run tests:

```powershell
npm test
```

## Deployment

This application requires:

- Node.js server
- Database service

GitHub Pages can host only static content, so deploy the application as a Node.js web service.

Deployment options included:

- Koyeb + Neon PostgreSQL
- Render deployment using `render.yaml`
- Docker deployment

See `DEPLOYMENT.md` for complete deployment steps.

## Important Note

- New passwords are stored as hashes.
- Existing plain-text passwords continue to work once and are automatically upgraded to hashed passwords after successful login.
- Email notifications require valid Gmail credentials and App Passwords.


## Note 
- If you want to generate the certificates, a demo excel file of the participants is already in frontend folder, use that kindly


##                          DEVELOPED BY VENKATA LALITH KARTHEEK VUPPULURI
=======
# CertiGen

