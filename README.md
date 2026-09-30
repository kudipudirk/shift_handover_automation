# ITConvergence Multi-Team Shift Handover Portal

A secure, multi-team daily shift handover web portal built with Flask, SQLite, background thread email alerts, and date-wise tabbed Excel reporting.

---

## 🚀 Overview & Key Features

- **3-Tier Role Management:**
  - **ADMIN:** Complete administrative access over user creation, team creation, password resets, and user management.
  - **LEAD:** Cross-team visibility and switching capabilities, Excel report exports, and audit history access without access to user management tabs.
  - **ENGINEER:** Standard team-restricted execution view for entering and managing daily shift tasks.
- **Automated P1/P2 Critical Email Alerts:** Sends immediate background email notifications when high-priority tickets are generated.
- **Force Password Change Workflow:** First-time logins require users to update their temporary passwords.
- **Date-Tabbed Excel Reports:** Formatted downloads featuring Dark Blue Banner headers, shift subheadings, color-coded priorities, and individual tabs per date.
- **Strict Team Audit Isolation:** Audit logs and history remain segmented by team scope.

---

## 🛠️ Deployment Instructions for New Server

### Step 1: System Dependencies
```bash
# RHEL / CentOS
sudo yum install -y python3 python3-pip python3-devel gcc

# Ubuntu / Debian
sudo apt update && sudo apt install -y python3 python3-pip python3-venv build-essential
```

### Step 2: Clone Repository
```bash
cd /opt
sudo git clone <YOUR_GIT_REPOSITORY_URL> handover_app
cd /opt/handover_app
```

### Step 3: Set Up Python Virtual Environment
```bash
python3 -m venv venv
./venv/bin/pip install --upgrade pip
./venv/bin/pip install -r requirements.txt
./venv/bin/pip install gunicorn
```

### Step 4: Initialize Fresh Database (`handover.db`)
```bash
./venv/bin/python3 -c "import sqlite3
from werkzeug.security import generate_password_hash

conn = sqlite3.connect('handover.db')
cursor = conn.cursor()

# Tasks Table (Updated with handover_notes & original_created_at)
cursor.execute('''CREATE TABLE IF NOT EXISTS tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    shift_name TEXT,
    customer TEXT,
    ticket_no TEXT,
    task_description TEXT,
    resource_from TEXT,
    resource_to TEXT,
    status TEXT DEFAULT 'PENDING',
    comments TEXT,
    priority TEXT DEFAULT 'P3',
    team_name TEXT DEFAULT 'Hosting',
    handover_notes TEXT DEFAULT '',
    original_created_at TIMESTAMP,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)''')

# Users Table
cursor.execute('''CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE,
    full_name TEXT,
    email TEXT,
    role TEXT,
    password TEXT,
    team_name TEXT,
    force_password_change INTEGER DEFAULT 1,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)''')

# Teams Table
cursor.execute('''CREATE TABLE IF NOT EXISTS teams (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    team_name TEXT UNIQUE
)''')

# Audit Logs Table
cursor.execute('''CREATE TABLE IF NOT EXISTS audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id INTEGER,
    action_type TEXT,
    details TEXT,
    performed_by TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)''')

# Default Teams
default_teams = ['Hosting', 'DBA', 'Linux/Unix', 'Windows/Cloud', 'Network', 'DevSecOps']
for t in default_teams:
    cursor.execute('INSERT OR IGNORE INTO teams (team_name) VALUES (?)', (t,))

# Default Admin Account (Temp Password: ITC@admin123!)
admin_pw = generate_password_hash('ITC@admin123!')
cursor.execute('''INSERT OR IGNORE INTO users (username, full_name, email, role, password, team_name, force_password_change)
VALUES ('admin', 'System Administrator', 'rkudipudi@itconvergence.com', 'ADMIN', ?, 'ALL_TEAMS', 1)''', (admin_pw,))

conn.commit()
conn.close()
print('Database and admin user initialized successfully!')"
```

### Step 5: Systemd Service Configuration
Create `/etc/systemd/system/handover.service`:

```ini
[Unit]
Description=ITConvergence Shift Handover Portal
After=network.target

[Service]
User=root
WorkingDirectory=/opt/handover_app
ExecStart=/opt/handover_app/venv/bin/gunicorn --workers 3 --bind 0.0.0.0:5000 app:app
Restart=always

[Install]
WantedBy=multi-user.target
```

Enable and start the service:
```bash
sudo systemctl daemon-reload
sudo systemctl enable handover
sudo systemctl start handover
```

---

#### Optional Nginx Reverse Proxy Config (`/etc/nginx/conf.d/handover.conf`):
```nginx
server {
    listen 80;
    server_name cms-365-stigtst09.itciss.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    }
}
```
### 💾 Automated Database Backup & Retention Setup:
To ensure data protection without overwhelming server storage, set up a daily automated backup script that automatically purges backups older than 2 days.

Step 1: Create Backup Directory and Script
```bash
mkdir -p /opt/handover_db_backups

cat << 'EOF' > /opt/handover_db_backups/backup_db.sh
#!/bin/bash

BACKUP_DIR="/opt/handover_db_backups"
DB_SOURCE="/opt/handover_app/handover.db"
DATE=$(date +%Y%m%d_%H%M%S)
BACKUP_FILE="$BACKUP_DIR/handover_$DATE.db"

# Create a new timestamped backup
if [ -f "$DB_SOURCE" ]; then
    cp "$DB_SOURCE" "$BACKUP_FILE"
    echo "[$DATE] ✅ Backup created: $BACKUP_FILE"
else
    echo "[$DATE] ❌ Error: Source database file missing!"
    exit 1
fi

# Automatically remove backup files older than 2 days
find "$BACKUP_DIR" -name "handover_*.db" -type f -mtime +2 -exec rm -f {} \;
echo "[$DATE] 🧹 Removed database backups older than 2 days."
EOF

chmod +x /opt/handover_db_backups/backup_db.sh
```

Step 2: Configure Daily Midnight Cron Job
```bash
(crontab -l 2>/dev/null; echo "0 0 * * * /opt/handover_db_backups/backup_db.sh >> /opt/handover_db_backups/backup.log 2>&1") | crontab -
```

### 🌐 Nginx Architecture & SSL/HTTPS Setup
Production Nginx Reverse Proxy & SSL Configuration (/etc/nginx/conf.d/handover.conf)

Step 1: Install Nginx
```bash
sudo yum install -y nginx
```

Step 2: Configure Nginx Site File
Create or update /etc/nginx/conf.d/handover.conf:
```bash
# 1. Redirect HTTP to HTTPS
server {
    listen 80;
    server_name cms-365-stigtst09.itciss.com handover.itconvergence.com;
    return 301 https://$host$request_uri;
}

# 2. HTTPS Reverse Proxy Server
server {
    listen 443 ssl;
    server_name cms-365-stigtst09.itciss.com handover.itconvergence.com;

    # SSL Certificates
    ssl_certificate /etc/pki/tls/certs/handover.crt;
    ssl_certificate_key /etc/pki/tls/private/handover.key;

    # Recommended SSL Protocols & Ciphers
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers HIGH:!aNULL:!MD5;
    ssl_prefer_server_ciphers on;

    location / {
        proxy_pass [http://127.0.0.1:5000](http://127.0.0.1:5000);
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Step 3: Enable Firewall Ports and Reload Nginx
```bash
sudo firewall-cmd --permanent --add-service=http
sudo firewall-cmd --permanent --add-service=https
sudo firewall-cmd --reload

sudo nginx -t
sudo systemctl enable nginx
sudo systemctl restart nginx
```
