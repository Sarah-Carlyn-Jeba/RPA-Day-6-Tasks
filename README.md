📧 RPA Day 6 – Email Automation using UiPath

This project demonstrates end-to-end email automation using UiPath, integrating Excel data processing, SMTP email sending, and IMAP email filtering. It simulates real-world business scenarios like sending personalized emails, organizing inbox messages, and notifying results automatically.

🚀 Overview

In this project, multiple email automation tasks were implemented:

Reading user data (Name, Email, Body, Marks) from Excel and Sending personalized emails using SMTP

Filtering and sorting emails using IMAP based on subject keywords

Sending automated result notifications based on marks

This project highlights how UiPath can streamline communication workflows efficiently.

🛠️ Features Implemented

📤 1. Sending Emails using SMTP
Used Send SMTP Mail Message activity
Dynamically passed:
Recipient Email (from Excel)
Subject
Personalized Body (using Name + Message)

📂 2. Email Filtering using IMAP
Used Get IMAP Mail Messages activity
Applied filtering logic based on subject:
Condition	Action
Subject contains "job"	Move to Job Folder
Subject contains "assignment"	Move to Assignment Folder

✅ Case-insensitive filtering implemented using:

mail.Subject.ToLower.Contains("job")
mail.Subject.ToLower.Contains("assignment")
Emails are automatically moved into:
📁 Job Folder
📁 Assignment Folder
Folders can be created manually in Gmail or automated using UiPath (if supported).

📊 3. Result Notification System
Based on Marks from Excel:
Send result notification email to users

🔄 Workflow Steps

Read Excel data using Read Range
Loop through each row using For Each Row
Send personalized email using SMTP
Fetch emails using IMAP
Apply subject-based filtering
Move emails to respective folders
Send result notification emails

🧰 Tools & Technologies

UiPath Studio
Excel Automation (Excel Process Scope)
SMTP Protocol (Send Email)
IMAP Protocol (Read & Filter Emails)

💡 Key Learnings

Integration of Excel with email automation
Dynamic email generation using data-driven approach
Email filtering using IMAP and string conditions
Building real-world automation workflows in UiPath

📎 Conclusion

This project showcases how UiPath can automate repetitive email tasks, improve productivity, and ensure efficient communication handling using structured workflows.

Author

Sarah Carlyn Jeba S
