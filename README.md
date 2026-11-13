# SendMail (console)

Small .NET console utility that reads email addresses from an Excel file and sends your resume to each address with a professional cold email message.

Features
- Reads an Excel (.xlsx) file and finds email addresses in a column named "Email"
- Sends a professional cold email with your resume (PDF) attached to each email address via SMTP
- Supports default paths in your Documents folder for convenience

Quick Start
1. Put these files in your Documents folder:
   - `contacts.xlsx` - Excel file with an "Email" column
   - `resume.pdf` - Your resume to attach to emails

2. Run the program once:
   ```bash
   dotnet run
   ```
   This will create a default `appsettings.json` with your Gmail settings.

3. Create a Gmail App Password:
   - Go to your Google Account settings
   - Security > 2-Step Verification > App passwords
   - Generate new app password for 'Mail'
   - Copy the 16-character password

4. Edit `appsettings.json` to add your Gmail App Password

5. Run again to send emails:
   ```bash
   dotnet run
   ```

Security Notes
- Use Gmail App Password, never your account password
- Enable 2-Step Verification for your Google Account first
- Keep appsettings.json private (contains password)

Default Paths
- Excel file: ~/Documents/contacts.xlsx
- Resume: ~/Documents/resume.pdf
- Config: ./appsettings.json

Usage
1. Build:
```bash
dotnet build
```

2. Run (using default paths):
```bash
dotnet run
```

3. Run (specify paths):
```bash
dotnet run -- path/to/contacts.xlsx path/to/appsettings.json
```

Email Template
- Professional cold email format included
- Customizable subject and body in appsettings.json
- Supports placeholders: {name} for your name

Notes & limitations
- This extracts emails with a regex from the PDF text. If the PDF encodes text unusually (images or scanned pages), you'll need OCR first.
- Be careful with sending limits and anti-spam rules. Add throttling / logging for production.
- For HR Contact Details, connect on work.harshgoyal@gmail.com

Enjoy!
