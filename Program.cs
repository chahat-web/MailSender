using System.Text.RegularExpressions;
using System.Text.Json;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

// Simple console utility:
// - Read an Excel (.xlsx) file where one column header is named "Email"
// - Send a personalized email with an attached resume to each address

// Setup default paths
var documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
var defaultExcelPath = Path.Combine(documentsPath, "contacts.xlsx");
var defaultResumePath = Path.Combine(documentsPath, "Resume.pdf");
var defaultConfigPath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");

// Use command line args if provided, otherwise use defaults
var excelPath = args.Length > 0 ? args[0] : defaultExcelPath;
var settingsPath = args.Length > 1 ? args[1] : defaultConfigPath;

// Create default config if it doesn't exist
if (!File.Exists(settingsPath))
{
    var defaultConfig = new Config
    {
        SmtpHost = "smtp.gmail.com",
        SmtpPort = 587,
        SmtpUser = "chahatbhandari30@gmail.com",
        FromName = "Chahat Bhnadari",
        FromEmail = "chahatbhandari30@gmail.com",
        Subject = "Full Stack Developer Position - Experienced Software Engineer",
        BodyTemplate = "Dear {recipientName},\n\nI hope this email finds you well. I am writing to express my strong interest in the Full Stack Developer position at your organization. With extensive experience in full-stack development, I bring expertise in modern web technologies including React, Node.js, .NET Core, and cloud platforms.\n\nMy technical skills include:\n- Frontend: React.js, TypeScript, HTML5/CSS3\n- Backend: .NET Core, Node.js, Python\n- Database: SQL Server, MongoDB\n- Cloud: AWS, Azure, Docker\n\nI believe my experience in building scalable applications and commitment to clean, maintainable code would make me a valuable addition to your team.\n\nI have attached my resume for your review. I would welcome the opportunity to discuss how my skills align with your team's needs.\n\nThank you for considering my application.\n\nBest regards,\n{name}\nPhone: {phone}\nEmail: {email}\nGitHub: https://github.com/chahat-web",
        ResumeAttachmentPath = defaultResumePath,
        UseStartTls = true,
        ContactPhone = "+91 9351601168"
    };
    
    File.WriteAllText(settingsPath, JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions 
    { 
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    }));
    
    Console.WriteLine($"Created default config at: {settingsPath}");
    Console.WriteLine("Please edit this file to add your Gmail App Password before running again.");
    Console.WriteLine("To create an App Password for Gmail:");
    Console.WriteLine("1. Go to your Google Account settings");
    Console.WriteLine("2. Select Security > 2-Step Verification > App passwords");
    Console.WriteLine("3. Generate a new app password for 'Mail'");
    Console.WriteLine("4. Copy the 16-character password into appsettings.json");
    return;
}

if (!File.Exists(excelPath))
{
	Console.WriteLine($"Excel file not found: {excelPath}");
	Console.WriteLine($"Please create an Excel file with an 'Email' column at: {excelPath}");
	Console.WriteLine("The Excel should have these columns in the first row:");
	Console.WriteLine("| Email              | Name (optional) | Company (optional) |");
	Console.WriteLine("| user@company.com   | John Smith     | Tech Corp         |");
	return;
}

Config? cfg = null;
try
{
	var json = File.ReadAllText(settingsPath);
	cfg = JsonSerializer.Deserialize<Config>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
}
catch (Exception ex)
{
	Console.WriteLine($"Failed to read/parse {settingsPath}: " + ex.Message);
	return;
}

if (cfg == null)
{
	Console.WriteLine("Invalid configuration format. Delete appsettings.json and run again to create a new one.");
	return;
}

// Debug: print which settings file we loaded and a short preview of the template.
try
{
	Console.WriteLine($"Using settings file: {settingsPath}");
	string templatePreview;
	if (string.IsNullOrEmpty(cfg.BodyTemplate))
		templatePreview = "<empty>";
	else
	{
		templatePreview = cfg.BodyTemplate.Replace("\n", "\\n");
		if (templatePreview.Length > 200)
			templatePreview = templatePreview.Substring(0, 200) + "...";
	}
	Console.WriteLine($"BodyTemplate preview: {templatePreview}");
}
catch { }

if (string.IsNullOrEmpty(cfg.SmtpPass) || cfg.SmtpPass == "your-app-password-here")
{
	Console.WriteLine("Please edit appsettings.json to add your Gmail App Password.");
	Console.WriteLine("To create an App Password:");
	Console.WriteLine("1. Go to your Google Account settings");
	Console.WriteLine("2. Select Security > 2-Step Verification > App passwords");
	Console.WriteLine("3. Generate a new app password for 'Mail'");
	Console.WriteLine("4. Copy the 16-character password into appsettings.json");
	return;
}

if (!File.Exists(cfg.ResumeAttachmentPath))
{
	Console.WriteLine($"Resume PDF not found: {cfg.ResumeAttachmentPath}");
	return;
}

Console.WriteLine($"Reading Excel file: {excelPath}");
Console.WriteLine("Extracting email addresses from 'Email' column...");
var emails = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
var recipientNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
var companies = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

using (SpreadsheetDocument document = SpreadsheetDocument.Open(excelPath, false))
{
	var workbookPart = document.WorkbookPart;
	var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>();
	var firstSheet = sheets?.FirstOrDefault();
	if (firstSheet == null)
	{
		Console.WriteLine("Excel contains no sheets.");
		return;
	}

	var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id!);
	var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
	if (sheetData == null)
	{
		Console.WriteLine("Excel sheet is empty.");
		return;
	}

	var rows = sheetData.Elements<Row>().ToList();
	if (rows.Count == 0)
	{
		Console.WriteLine("Excel appears empty.");
		return;
	}

	// header is first row
	var header = rows[0];
	int emailColIndex = -1;
	int nameColIndex = -1;
	int companyColIndex = -1;
	
	for (int i = 0; i < header.Elements<Cell>().Count(); i++)
	{
		var cell = header.Elements<Cell>().ElementAt(i);
		var val = GetCellValue(document, cell)?.Trim();
		if (string.Equals(val, "Email", StringComparison.OrdinalIgnoreCase))
		{
			emailColIndex = i;
		}
		else if (string.Equals(val, "Name", StringComparison.OrdinalIgnoreCase))
		{
			nameColIndex = i;
		}
		else if (string.Equals(val, "Company", StringComparison.OrdinalIgnoreCase))
		{
			companyColIndex = i;
		}
	}

	if (emailColIndex == -1)
	{
		Console.WriteLine("Could not find a header named 'Email' in the first sheet.");
		return;
	}

	var emailRegex = new Regex("[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}", RegexOptions.Compiled);
	for (int r = 1; r < rows.Count; r++)
	{
		var row = rows[r];
		var cells = row.Elements<Cell>().ToList();
		if (emailColIndex >= cells.Count)
			continue;

		var emailCell = cells[emailColIndex];
		var emailText = GetCellValue(document, emailCell)?.Trim();
		if (string.IsNullOrWhiteSpace(emailText))
			continue;

		string? recipientName = null;
		if (nameColIndex >= 0 && nameColIndex < cells.Count)
		{
			var nameCell = cells[nameColIndex];
			recipientName = GetCellValue(document, nameCell)?.Trim();
		}

		string? company = null;
		if (companyColIndex >= 0 && companyColIndex < cells.Count)
		{
			var companyCell = cells[companyColIndex];
			company = GetCellValue(document, companyCell)?.Trim();
		}

		// Extract each email from the cell
		foreach (var part in emailText.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
		{
			var email = part.Trim();
			var m = emailRegex.Match(email);
			if (m.Success)
			{
				var validEmail = m.Value;
				emails.Add(validEmail);
				
				if (!string.IsNullOrWhiteSpace(recipientName))
				{
					recipientNames[validEmail] = recipientName;
				}

				if (!string.IsNullOrWhiteSpace(company))
				{
					companies[validEmail] = company;
				}
			}
		}
	}
}

static string? GetCellValue(SpreadsheetDocument doc, Cell cell)
{
	if (cell == null)
		return null;

	var value = cell.CellValue?.InnerText;
	if (value == null)
		return null;

	if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
	{
		var sstPart = doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
		if (sstPart != null)
		{
			return sstPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
		}
	}

	return value;
}



if (emails.Count == 0)
{
	Console.WriteLine("No email addresses found in the Excel file.");
	return;
}

Console.WriteLine($"Found {emails.Count} unique email(s). Preparing to send...");

int sent = 0;
using var client = new SmtpClient();
try
{
	var secureOption = cfg.UseStartTls ? SecureSocketOptions.StartTls : (cfg.UseSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.Auto);
	Console.WriteLine($"Connecting to SMTP {cfg.SmtpHost}:{cfg.SmtpPort} (ssl:{cfg.UseSsl}, starttls:{cfg.UseStartTls})...");
	client.Connect(cfg.SmtpHost, cfg.SmtpPort, secureOption);

	if (!string.IsNullOrEmpty(cfg.SmtpUser))
	{
		client.Authenticate(cfg.SmtpUser, cfg.SmtpPass);
	}

	foreach (var emailAddress in emails)
	{
		try
		{
			var message = new MimeMessage();
			message.From.Add(new MailboxAddress(cfg.FromName ?? "", cfg.FromEmail));
			
			// Use recipient name in the To field if available
			string recipientName = recipientNames.TryGetValue(emailAddress, out var name) ? name : "Hiring Manager";
			message.To.Add(new MailboxAddress(recipientName, emailAddress));
			
			// Add company name to subject if available
			var company = companies.TryGetValue(emailAddress, out var companyName) ? companyName : "";
			message.Subject = string.IsNullOrWhiteSpace(cfg.Subject) ? 
				$"Software Engineer Application - {cfg.FromName} ({company})".Trim() : 
				cfg.Subject.Replace("{COMPANY}", company);

			// Add important headers to improve deliverability
			message.Headers.Add("List-Unsubscribe", $"<mailto:{cfg.FromEmail}?subject=unsubscribe>");
			message.Headers.Add("Precedence", "Bulk");
			message.Headers.Add("X-Auto-Response-Suppress", "OOF, AutoReply");
			
			// Set message priority to normal (spam often uses high priority)
			message.Priority = MessagePriority.Normal;

			var builder = new BodyBuilder();

				// Build the email body using the configured template or a sensible default.
				string defaultTemplate = "Dear {recipientName},\n\n" +
                    "I trust this email finds you well. I am writing to express my strong interest in the Full Stack Developer position at your organization. " +
                    "With extensive experience in full-stack development, I bring expertise in modern web technologies including React, Node.js, .NET Core, and cloud platforms.\n\n" +
                    "My technical skills include:\n" +
                    "- Languages: C#, Java, Python, JavaScript, TypeScript, SQL, C, C++\n" +
                    "- Frameworks: .NET Core, React.js, Spring Boot\n" +
                    "- Frontend: HTML, CSS, Ant Design, AG Grid, JavaScript\n" +
                    "- Databases: MySQL, MongoDB\n" +
                    "- Tools: Git, GitHub, Postman, Swagger, RabbitMQ\n" +
                    "- Practices: OOP, DSA, REST APIs, Microservices, HLD/LLD, Agile, Scrum\n" +
					"I believe my experience in building scalable applications and commitment to clean, maintainable code would make me a valuable addition to your team.\n\n" +
					"I have attached my resume for your review. I would welcome the opportunity to discuss how my skills align with your team's needs.\n\n" +
					"Thank you for considering my application.\n\n" +
					"Best regards,\n" +
					"{name}\n" +
					"Phone: {phone}\n" +
					"Email: {email}\n";

				// Choose which template to use. If PreferDefaultTemplate is true, always use the built-in default.
				string template = (!cfg.PreferDefaultTemplate && !string.IsNullOrWhiteSpace(cfg.BodyTemplate))
					? cfg.BodyTemplate
					: defaultTemplate;

				string bodyText = template
					.Replace("{recipientName}", recipientName)
					.Replace("{COMPANY}", company)
					.Replace("{email}", cfg.FromEmail)
					.Replace("{name}", cfg.FromName ?? "Harsh Goyal")
					.Replace("{phone}", cfg.ContactPhone);

				// Removed company name injection from greeting - keeping it clean and personal

				builder.TextBody = bodyText;

				// Attach resume
				using var resumeStream = File.OpenRead(cfg.ResumeAttachmentPath);
				builder.Attachments.Add(Path.GetFileName(cfg.ResumeAttachmentPath), resumeStream);

				message.Body = builder.ToMessageBody();

				// -- Bulk-sending safety features --
				// Local helper: try sending with retries and exponential backoff
				bool TrySendWithRetries(MimeMessage msg)
				{
					int attempt = 0;
					int maxAttempts = Math.Max(1, cfg.MaxRetries);
					while (true)
					{
						try
						{
							client.Send(msg);
							return true;
						}
						catch (Exception ex)
						{
							attempt++;
							if (attempt > maxAttempts)
							{
								Console.WriteLine($"Send failed after {attempt - 1} retries: {ex.Message}");
								return false;
							}
							int backoff = (int)Math.Pow(2, attempt) * 1000;
							Console.WriteLine($"Transient send error (attempt {attempt}/{maxAttempts}), backing off {backoff}ms: {ex.Message}");
							Thread.Sleep(backoff);
						}
					}
				}

				Console.WriteLine($"Sending to {emailAddress}...");
				var success = TrySendWithRetries(message);
				if (success)
				{
					sent++;
					Console.WriteLine($"Sent ({sent}/{emails.Count})");

					// Implement careful rate limiting for large batches
					int delayMs;
					if (emails.Count <= 50)
						delayMs = 3000;  // 3s for very small batches
					else if (emails.Count <= 200)
						delayMs = 5000;  // 5s for small batches
					else
						delayMs = 8000;  // 8s for large batches

					// Add some randomness to delays to look more natural
					var rnd = new Random();
					delayMs += rnd.Next(-500, 501);  // +/- 0.5s random variation
					Thread.Sleep(delayMs);

					// Take longer pauses more frequently for large batches
					if (sent % 50 == 0 && sent < emails.Count)
					{
						var pauseMinutes = 3;  // 3-minute pause every 50 emails
						Console.WriteLine($"Completed {sent} sends. Taking a {pauseMinutes}-minute pause to avoid spam filters...");
						Thread.Sleep(pauseMinutes * 60 * 1000);
					}
				}
				else
				{
					Console.WriteLine($"Failed to send to {emailAddress} after retries.");
				}
		}
		catch (Exception ex)
		{
			Console.WriteLine($"Failed to send to {emailAddress}: {ex.Message}");
		}
	}

	client.Disconnect(true);
}
catch (Exception ex)
{
	Console.WriteLine("SMTP error: " + ex.Message);
}

Console.WriteLine($"Completed. Sent {sent} messages (attempted {emails.Count}).");

// Config POCO
public class Config
{
	public string SmtpHost { get; set; } = "";
	public int SmtpPort { get; set; } = 587;
	public string SmtpUser { get; set; } = "";
	public string SmtpPass { get; set; } = "";
	public string FromName { get; set; } = "Chahat Bhandari";
	public string FromEmail { get; set; } = "";
	public string Subject { get; set; } = "";
	public string BodyTemplate { get; set; } = "";

	// Send controls
	public int MaxRetries { get; set; } = 3; // Number of retry attempts on failure

	// When true the program will ignore BodyTemplate from the config and use the built-in default template.
	public bool PreferDefaultTemplate { get; set; } = false;
	public string ResumeAttachmentPath { get; set; } = "Resume.pdf"; // Your CV/resume PDF to attach to each email
	public bool UseSsl { get; set; } = false;
	public bool UseStartTls { get; set; } = true;
	public string ContactPhone { get; set; } = "";
}
