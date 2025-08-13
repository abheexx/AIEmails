# Outlook Draft Creator

A Python tool that creates Outlook drafts from an Excel contact list using Microsoft Graph API. Perfect for cold email campaigns where you want to review and personalize drafts before sending.

## Features

- 📧 Create Outlook drafts (not sent) from Excel contact lists
- 🔐 OAuth2 authentication with Microsoft Graph API
- 📝 Jinja2 templating for personalized subject and body
- 🤖 Optional AI personalization using OpenAI API
- 📊 CSV logging of all draft creation results
- ⏱️ Rate limiting to avoid API throttling
- 💾 Token caching for seamless re-authentication

## Prerequisites

- Python 3.7+
- Microsoft 365 account (school/work account recommended)
- Azure AD App Registration (see setup below)
- Excel file with contact information

## Installation

1. Clone or download this project
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Copy `env.example` to `.env` and configure your settings

## Azure AD App Registration Setup

### Step 1: Create App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Fill in:
   - **Name**: `Outlook Draft Creator` (or any name you prefer)
   - **Supported account types**: `Accounts in this organizational directory only` (single tenant)
   - **Redirect URI**: Leave blank (we're using device code flow)
5. Click **Register**

### Step 2: Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph** → **Delegated permissions**
4. Add these permissions:
   - `Mail.ReadWrite` (to create drafts)
   - `offline_access` (for token refresh)
5. Click **Grant admin consent** (may require admin approval)

### Step 3: Get Client ID

1. In your app registration, go to **Overview**
2. Copy the **Application (client) ID**
3. Add it to your `.env` file:
   ```
   CLIENT_ID=your_client_id_here
   ```

## Excel File Format

Your Excel file should have these columns (some may be blank):

| Column | Description | Example |
|--------|-------------|---------|
| `first_name` | Contact's first name | John |
| `last_name` | Contact's last name | Smith |
| `email` | Contact's email address | john@company.com |
| `company` | Company name | TechCorp |
| `role` | Job title/role | Software Engineer |
| `observation` | Any notes/observations | Interested in Python |

## Customizing Email Templates

Edit the `SUBJECT_TEXT` and `BODY_TEXT` variables in `main.py`:

```python
# === USE EXACTLY; DO NOT CHANGE WORDING ===
SUBJECT_TEXT = """Connecting with {{first_name}} from {{company}}"""

BODY_TEXT = """Hi {{first_name}},

I noticed your role as {{role}} at {{company}} and wanted to connect.

{{observation}}

Best regards,
Your Name
"""
# ==========================================
```

### Available Placeholders

- `{{first_name}}` - Contact's first name
- `{{last_name}}` - Contact's last name  
- `{{email}}` - Contact's email
- `{{company}}` - Company name
- `{{role}}` - Job title/role
- `{{observation}}` - Any notes you've added

**Note**: If a placeholder doesn't exist in your Excel data, it will render as an empty string (no crashes).

## Usage

### Basic Usage

```bash
python main.py \
  --excel contacts_example.xlsx \
  --sheet Contacts \
  --from-name "Your Name" \
  --from-email "your.email@university.edu" \
  --client-id "your_azure_client_id" \
  --log-csv draft_log.csv
```

### With AI Personalization

```bash
python main.py \
  --excel contacts_example.xlsx \
  --sheet Contacts \
  --from-name "Your Name" \
  --from-email "your.email@university.edu" \
  --client-id "your_azure_client_id" \
  --log-csv draft_log.csv \
  --ai-personalize \
  --delay-ms 2000
```

### Command Line Arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `--excel` | Yes | Path to Excel file |
| `--sheet` | Yes | Sheet name in Excel |
| `--from-name` | Yes | Your name for logging |
| `--from-email` | Yes | Your email for logging |
| `--client-id` | Yes | Azure AD App Client ID |
| `--log-csv` | Yes | CSV file to log results |
| `--delay-ms` | No | Delay between drafts (default: 1000ms) |
| `--ai-personalize` | No | Enable AI personalization |

## Authentication Flow

1. **First Run**: The tool will display a URL and code
2. **Browser**: Open the URL and enter the code
3. **Sign In**: Use your Microsoft 365 credentials
4. **Token Caching**: Tokens are saved locally for future use
5. **Re-authentication**: Automatic token refresh when needed

## Output

### Draft Creation
- Creates drafts in your Outlook account (not sent)
- Each draft includes personalized subject and body
- Drafts appear in your Outlook Drafts folder

### CSV Logging
The tool logs results to your specified CSV file:

| Column | Description |
|--------|-------------|
| `email` | Contact's email address |
| `company` | Company name |
| `first_name` | Contact's first name |
| `subject` | Rendered subject line |
| `draft_id` | Outlook draft ID (if successful) |
| `status` | `success` or `error` |

## AI Personalization

When using `--ai-personalize`:

1. Requires `OPENAI_API_KEY` in your `.env` file
2. Analyzes contact information to generate personalized content
3. Appends ≤35-word tailored blurb to each email body
4. Silently skips if API key is missing or invalid

## Rate Limiting

- Default delay: 1 second between drafts
- Adjust with `--delay-ms` parameter
- Recommended: 1000-3000ms to avoid API throttling
- Microsoft Graph has rate limits that vary by plan

## Troubleshooting

### Common Issues

1. **"Authentication failed"**
   - Check your Azure AD App permissions
   - Ensure admin consent was granted
   - Verify client ID is correct

2. **"Permission denied"**
   - Ensure `Mail.ReadWrite` permission is granted
   - Check if admin consent is required

3. **"Excel file not found"**
   - Verify file path is correct
   - Check file permissions

4. **"Template rendering error"**
   - Check placeholder syntax (use `{{variable}}`)
   - Ensure Excel column names match placeholders

### Token Issues

- Delete `token_cache.bin` to force re-authentication
- Check if your Microsoft 365 account has the required permissions
- Verify your account hasn't been locked or expired

## Security Notes

- Tokens are stored locally in `token_cache.bin`
- Never commit `.env` or `token_cache.bin` to version control
- The tool only creates drafts (never sends emails automatically)
- Review all drafts before sending

## Compliance & Quotas

- **Microsoft Graph**: Check your plan's API limits
- **School/Work Accounts**: May have additional restrictions
- **Draft Storage**: Outlook has draft storage limits
- **Rate Limits**: Respect API throttling to avoid blocks

## Support

For issues related to:
- **Azure AD**: Check Azure Portal and Microsoft documentation
- **Microsoft Graph**: Review [Graph API documentation](https://docs.microsoft.com/en-us/graph/)
- **School/Work Policies**: Contact your IT administrator

## Example Workflow

1. **Prepare**: Set up Azure AD app and get client ID
2. **Configure**: Edit email templates in `main.py`
3. **Organize**: Prepare Excel file with contact information
4. **Run**: Execute the tool with appropriate parameters
5. **Review**: Check Outlook Drafts folder for created drafts
6. **Personalize**: Edit drafts as needed before sending
7. **Send**: Manually send when ready

## Files

- `main.py` - Main application script
- `requirements.txt` - Python dependencies
- `contacts_example.xlsx` - Sample contact data
- `env.example` - Environment variables template
- `README.md` - This documentation

Remember: This tool creates **drafts only**. You must manually review and send each email from Outlook. 