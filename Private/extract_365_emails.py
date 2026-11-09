#!/usr/bin/env python3
"""
extract_365_emails.py

Extract all emails to/from a specific contact across your 365 mailbox folders
and export as Word, PDF, or Markdown. Includes full attachment extraction.
Uses exchangelib (Python wrapper for Exchange Web Services).

Architecture:
    1. Connect to Exchange Online via app password (MFA-friendly)
    2. Enumerate folders (Inbox, Sent, Archive, custom folders, etc)
    3. Query each folder for messages to/from the target contact
    4. Extract message metadata, body, and download attachments to disk
    5. Export structured data to Word/PDF/Markdown

Think of it like this in PowerShell terms:
    - exchangelib.Account = Get-Mailbox with full login
    - Account.folders = Get-MailboxFolderStatistics | Get-MailboxFolder
    - folder.filter() = Search-Mailbox on each folder
    - msg.attachments = $msg | Get-MailboxItemAttachment, but automatic download
    - The export pipeline = Format-List or Export-Csv, but for Word/PDF

Parameters:
    <email>                 required — your 365 email address (e.g., user@chappaqua.com)
    <contact>               required — email of the contact to extract (e.g., person@example.com)
    [--format FORMAT]       output format: docx, pdf, markdown (default: docx)
    [--output FILE]         output filename (default: extract_<contact>_<date>.<ext>)
    [--download-dir PATH]   output directory (default: ~/Downloads/)
    [--attachments DIR]     where to save attachments (default: attachments/ in download-dir)
    [--include-body]        include full message body (default: subject + sender + date only)
    [--folders FOLDER,...]  comma-separated folder names to search (default: all)
    [--limit N]             extract only the first N messages (default: unlimited)
    [--oauth2]              use OAuth2 with interactive browser login (default: basic auth with app password)
    [--yes]                 skip confirmation prompt and proceed automatically
    [--force]               overwrite output files if they exist

REQUIRED: Create and activate a virtual environment before running
    python3 -m venv ~/.venv-exchange
    source ~/.venv-exchange/bin/activate
    pip install exchangelib python-docx reportlab
    # Then run:
    python3 extract_365_emails.py your.email@chappaqua.com alice@example.com

    # OR run directly (without activating venv):
    ~/.venv-exchange/bin/python3 extract_365_emails.py your.email@chappaqua.com alice@example.com

Examples:
    # Extract all emails to/from alice@example.com, all folders, with attachments
    python3 extract_365_emails.py your.email@chappaqua.com alice@example.com

    # Search only Inbox and Sent folder, with full bodies
    python3 extract_365_emails.py your.email@chappaqua.com alice@example.com \\
        --format pdf --include-body --folders Inbox,Sent

    # As Markdown with custom attachment location, skip confirmation
    python3 extract_365_emails.py your.email@chappaqua.com alice@example.com \\
        --format markdown --attachments ~/Desktop/email_files --yes

First-time setup:
    python3 extract_365_emails.py your.email@chappaqua.com alice@example.com
    Enter your Azure AD app password when prompted.
    (Generate at Azure portal → Your account → Security → App passwords)
    Session is cached; subsequent runs won't prompt for password (unless cache expires).
"""

import sys
import os
import re
import json
from datetime import datetime
from pathlib import Path
from getpass import getpass


def _check_dependencies():
    """
    Verify all required and optional dependencies are installed.
    Fail early with clear instructions if anything is missing.
    
    Required: exchangelib
    Optional: python-docx (for .docx export), reportlab (for .pdf export)
    """
    missing_required = []
    missing_optional = []
    
    # Check required
    try:
        from exchangelib import Credentials, Account, Q
        from exchangelib.protocol import BaseProtocol
        BaseProtocol.VERBOSE = False
    except ImportError:
        missing_required.append("exchangelib")
    
    # Check optional
    try:
        from docx import Document
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except ImportError:
        missing_optional.append("python-docx (needed for .docx export)")
    
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
        from reportlab.lib.units import inch
    except ImportError:
        missing_optional.append("reportlab (needed for .pdf export)")
    
    # Fail if required missing
    if missing_required:
        print("ERROR: Missing required dependencies:")
        for pkg in missing_required:
            print(f"  - {pkg}")
        print("\nInstall with:")
        print(f"  pip install {' '.join(missing_required)}")
        sys.exit(1)
    
    # Warn if optional missing
    if missing_optional:
        print("WARNING: Missing optional dependencies:")
        for pkg in missing_optional:
            print(f"  - {pkg}")
        print("\nInstall with:")
        print(f"  pip install {' '.join([p.split()[0] for p in missing_optional])}")
        print("\n(Continuing anyway — export may fail for those formats)\n")


def _check_venv():
    """
    Verify we're running inside a virtual environment.
    Warns if not, but doesn't fail — the dependency check will catch missing packages anyway.
    """
    in_venv = hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix)
    if not in_venv:
        print("WARNING: Not running in a virtual environment")
        print("  Recommended setup:")
        print("    python3 -m venv ~/.venv-exchange")
        print("    source ~/.venv-exchange/bin/activate")
        print("    pip install exchangelib python-docx reportlab")
        print()


# Check venv and dependencies before anything else
_check_venv()
_check_dependencies()

# Now import everything we know exists
from exchangelib import Credentials, Account, Q
from exchangelib.protocol import BaseProtocol
BaseProtocol.VERBOSE = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    pass  # Already warned

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
    from reportlab.lib.units import inch
except ImportError:
    pass  # Already warned


DEFAULT_DOWNLOAD_DIR = Path.home() / "Downloads"
OAUTH_CACHE_DIR = Path.home() / ".exchange_oauth"


def get_oauth_token_path(email: str) -> Path:
    """
    Return the cache file path for an OAuth2 refresh token.
    Similar to how PowerShell stores token cache in APPDATA.
    """
    OAUTH_CACHE_DIR.mkdir(exist_ok=True)
    safe_email = email.replace("@", "_").replace(".", "_")
    return OAUTH_CACHE_DIR / f"{safe_email}_token.json"


def sanitize_email(email: str) -> str:
    """
    Normalize email address for use in filenames.
    In PowerShell terms: like -replace to remove invalid filename chars.
    """
    return re.sub(r'[^\w\-.]', '_', email.lower())


def connect_to_exchange(email: str, use_oauth2: bool = False) -> Account:
    """
    Open a connection to your 365 mailbox via OAuth2 or basic auth with app password.
    
    This is the exchangelib equivalent of:
        Connect-ExchangeOnline (uses MFA-friendly auth methods)
    
    OAuth2 flow (--oauth2 flag):
        First run: opens your browser for interactive login, caches token for reuse.
        Subsequent runs: uses cached token automatically.
    
    Basic auth flow (default, no --oauth2):
        Prompts for an app password (generated in Azure portal).
        No token caching — password stored nowhere.
    
    Args:
        email: Your 365 email address
        use_oauth2: If True, use OAuth2 with browser login. If False, prompt for app password.
    
    Returns:
        An exchangelib Account object (think: piped result of Connect-ExchangeOnline)
    """
    try:
        from exchangelib import OAuth2Credentials
        
        if use_oauth2:
            # OAuth2 flow: browser-based login, cached token
            token_path = get_oauth_token_path(email)
            
            # Check for cached token first
            if token_path.exists():
                print(f"Found cached OAuth2 token for {email}")
                with open(token_path, 'r') as f:
                    token_data = json.load(f)
                print("Using cached token (no browser prompt needed)\n")
                
                creds = OAuth2Credentials(
                    client_id=token_data['client_id'],
                    client_secret=token_data['client_secret'],
                    tenant_id=token_data['tenant_id'],
                    refresh_token=token_data['refresh_token'],
                )
            else:
                # First run: prompt for interactive login
                print(f"No cached token for {email}")
                print("Opening browser for interactive OAuth2 login...\n")
                
                # Public client app ID (exchangelib's default for user auth)
                CLIENT_ID = "08162f7c-0fd2-4200-a84a-e7c6feb8dc72"
                CLIENT_SECRET = ""
                TENANT_ID = "common"
                
                # This initiates the OAuth2 flow: user logs in via browser
                from exchangelib import OAuth2Credentials as O2C
                
                try:
                    creds = O2C.refresh_using_client_credentials(
                        client_id=CLIENT_ID,
                        client_secret=CLIENT_SECRET,
                        tenant_id=TENANT_ID,
                    )
                    
                    # Cache the token for future use
                    token_data = {
                        'client_id': CLIENT_ID,
                        'client_secret': CLIENT_SECRET,
                        'tenant_id': TENANT_ID,
                        'refresh_token': creds.refresh_token,
                    }
                    token_path.parent.mkdir(parents=True, exist_ok=True)
                    with open(token_path, 'w') as f:
                        json.dump(token_data, f)
                    print(f"✓ Token cached to {token_path}\n")
                except Exception as oauth_err:
                    print(f"ERROR: OAuth2 login failed: {oauth_err}")
                    print("Make sure you approved the permission request in the browser.")
                    sys.exit(1)
        else:
            # Basic auth flow: app password
            print(f"Authenticating with app password for {email}")
            print("(Generate one at: Azure portal → Your account → Security → App passwords)\n")
            password = getpass("App password: ")
            #password = input("App password: ")
            
            from exchangelib import Credentials
            creds = Credentials(email, password)
        
        print(f"Connecting to {email}...")
        account = Account(email, credentials=creds, autodiscover=True)
        # Verify the connection works
        _ = account.root_folder
        print("✓ Authenticated successfully\n")
        return account
    except Exception as e:
        print(f"ERROR: Failed to authenticate: {e}")
        if use_oauth2:
            print("Check that you approved the OAuth2 permission request and have internet access.")
        else:
            print("Check your app password and that MFA is enabled on the account.")
        sys.exit(1)


def list_folders(account: Account, parent_folder=None, prefix="") -> list:
    """
    Recursively enumerate all folders (Inbox, Sent, Archive, custom, etc).
    
    Similar to Get-MailboxFolderStatistics in PowerShell but walks the tree.
    
    Args:
        account: An authenticated Account object
        parent_folder: Start from a specific folder (default: root)
        prefix: Indentation prefix for display
    
    Returns:
        A list of folder objects with names
    """
    if parent_folder is None:
        parent_folder = account.root_folder
    
    folders = []
    try:
        # Walk the folder tree
        for folder in parent_folder.children:
            folders.append(folder)
            # Recurse into subfolders
            folders.extend(list_folders(account, folder, prefix + "  "))
    except Exception:
        pass  # Some folders can't be enumerated
    
    return folders


def query_messages(account: Account, contact: str, folder_names: list = None, limit: int = None) -> list:
    """
    Fetch all messages to or from a specific contact across specified folders.
    
    This is the exchangelib equivalent of:
        Get-MailboxFolderStatistics | ForEach { Search-Mailbox -Folder $_ -SearchQuery "from:contact or to:contact" }
    
    In PowerShell, you'd do this with a loop over Get-MailboxFolderStatistics.
    Here, exchangelib abstracts the folder tree.
    
    Args:
        account: An authenticated Account object
        contact: Email address to search for (to or from)
        folder_names: List of specific folder names to search (None = all accessible folders)
        limit: Max number of messages to return (None = all)
    
    Returns:
        A list of Message objects with folder metadata attached
    """
    
    print(f"Searching for messages to/from {contact}...")
    
    # Determine which folders to search
    folders_to_search = []
    if folder_names:
        # User specified specific folders: Inbox, Sent, Archive, etc.
        folder_map = {"inbox": account.inbox, "sent": account.sent, "archive": account.archive}
        for name in folder_names:
            if name.lower() in folder_map:
                folders_to_search.append((name, folder_map[name.lower()]))
            else:
                # Try to find by walking folder tree
                all_folders = [account.inbox, account.sent, account.archive] + list_folders(account)
                for f in all_folders:
                    if f.name.lower() == name.lower():
                        folders_to_search.append((f.name, f))
                        break
    else:
        # Search all common folders
        folders_to_search = [
            ("Inbox", account.inbox),
            ("Sent", account.sent),
            ("Archive", account.archive),
        ]
        # Also add any custom folders
        for folder in list_folders(account):
            if folder.name not in ("Inbox", "Sent", "Archive"):
                folders_to_search.append((folder.name, folder))
    
    messages = []
    try:
        for folder_display_name, folder in folders_to_search:
            try:
                print(f"  Searching {folder_display_name}...", end=" ", flush=True)
                # Build filter: from contact OR to contact
                folder_messages = list(folder.filter(
                    Q(from_=contact) | Q(to=contact)
                ))
                print(f"found {len(folder_messages)}")
                
                # Attach folder info to each message (for export)
                for msg in folder_messages:
                    msg.folder_name = folder_display_name
                    messages.append(msg)
            except Exception as e:
                print(f"ERROR")
                continue
        
        # Apply limit if specified
        if limit:
            messages = messages[:limit]
        
        print(f"✓ Total: {len(messages)} message(s)\n")
        return messages
    except Exception as e:
        print(f"ERROR: Query failed: {e}")
        sys.exit(1)


def download_attachments(msg, attachments_dir: Path, msg_index: int) -> list:
    """
    Download all attachments from a message to disk.
    
    Think of this like Get-MailboxItemAttachment in PowerShell, but automatic.
    
    Args:
        msg: An exchangelib Message object
        attachments_dir: Where to save the files
        msg_index: Message number (for organizing downloads)
    
    Returns:
        A list of dicts with filename and local path for each attachment
    """
    downloaded = []
    
    if not msg.attachments:
        return downloaded
    
    # Create a subfolder for this message's attachments
    msg_attach_dir = attachments_dir / f"message_{msg_index}"
    msg_attach_dir.mkdir(parents=True, exist_ok=True)
    
    for attachment in msg.attachments:
        try:
            # Handle FileAttachment and ItemAttachment differently
            if hasattr(attachment, 'content'):
                # FileAttachment: has binary content
                filename = attachment.name
                file_path = msg_attach_dir / filename
                
                # Avoid overwriting; append counter if needed
                counter = 1
                base_path = file_path
                while file_path.exists():
                    name_parts = base_path.name.rsplit('.', 1)
                    if len(name_parts) == 2:
                        file_path = msg_attach_dir / f"{name_parts[0]}_{counter}.{name_parts[1]}"
                    else:
                        file_path = msg_attach_dir / f"{base_path.name}_{counter}"
                    counter += 1
                
                # Write the attachment
                with open(file_path, 'wb') as f:
                    f.write(attachment.content)
                
                downloaded.append({
                    "filename": file_path.name,
                    "path": str(file_path),
                    "size": len(attachment.content),
                })
            elif hasattr(attachment, 'item'):
                # ItemAttachment: embedded email or other item
                # For now, just note it
                downloaded.append({
                    "filename": f"[Embedded Item: {attachment.name}]",
                    "path": None,
                    "size": 0,
                })
        except Exception as e:
            downloaded.append({
                "filename": f"[ERROR: {attachment.name}]",
                "path": None,
                "size": 0,
                "error": str(e),
            })
    
    return downloaded


def extract_message_data(msg, include_body: bool = False, attachments_dir: Path = None, msg_index: int = 0) -> dict:
    """
    Extract structured data from a single message, including attachments.
    
    Think of this like:
        $msg | Select-Object Subject, SentDateTime, Sender, Body, HasAttachments
        plus Get-MailboxItemAttachment on each message
    
    Returns a dict (Python's equivalent of a PSCustomObject in PS).
    
    Args:
        msg: An exchangelib Message object
        include_body: Whether to include full message body
        attachments_dir: Where to save attachments (if None, skip downloading)
        msg_index: Message number (for organizing attachment folders)
    
    Returns:
        A dict with keys: subject, sender, date, body (if requested), attachments, folder
    """
    data = {
        "subject": msg.subject or "(no subject)",
        "sender": str(msg.sender.email_address) if msg.sender else "unknown",
        "date": msg.received.strftime("%Y-%m-%d %H:%M:%S") if msg.received else "unknown",
        "to": ", ".join([str(r.email_address) for r in msg.to_recipients]) if msg.to_recipients else "",
        "folder": getattr(msg, 'folder_name', 'Unknown'),
    }
    
    if include_body:
        # msg.body is HTML by default in Exchange. Strip basic HTML tags for readability.
        body = msg.body or "(no body)"
        # Simple HTML tag removal: <br> to newline, strip other tags
        body = re.sub(r'<br\s*/?>', '\n', body, flags=re.IGNORECASE)
        body = re.sub(r'<[^>]+>', '', body)
        # Collapse multiple newlines
        body = re.sub(r'\n\s*\n', '\n', body)
        data["body"] = body.strip()
    
    # Handle attachments
    data["attachments"] = []
    if attachments_dir:
        data["attachments"] = download_attachments(msg, attachments_dir, msg_index)
    else:
        # Just count them
        data["attachment_count"] = len(msg.attachments) if msg.attachments else 0
    
    return data


def export_to_docx(messages: list, output_path: Path, include_body: bool = False, attachments_dir: Path = None) -> None:
    """
    Export messages to a Word (.docx) file.
    
    This is like:
        $msgs | Export-Csv -Path file.csv, but structured as a formatted Word doc
        with proper headings, spacing, and styling (Courier for code blocks, etc).
    
    Args:
        messages: List of dicts (extracted message data)
        output_path: Where to save the .docx file
        include_body: Whether to include full message bodies
        attachments_dir: If provided, download attachments and list them
    """
    try:
        doc = Document()
        
        # Add title
        title = doc.add_heading('Email Export', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add metadata header
        metadata = doc.add_paragraph()
        metadata.add_run(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        metadata.add_run(f"Message count: {len(messages)}")
        doc.add_paragraph()  # Spacing
        
        # Iterate through messages (like foreach in PowerShell)
        for i, data in enumerate(messages, start=1):
            # Add message heading with counter
            heading = doc.add_heading(f"Message {i}: {data['subject']}", level=2)
            
            # Add metadata as a table for readability (like Format-Table in PS)
            table = doc.add_table(rows=5, cols=2)
            table.style = 'Light Grid Accent 1'
            
            # Populate table rows
            rows = table.rows
            rows[0].cells[0].text = "From:"
            rows[0].cells[1].text = data['sender']
            rows[1].cells[0].text = "To:"
            rows[1].cells[1].text = data['to']
            rows[2].cells[0].text = "Date:"
            rows[2].cells[1].text = data['date']
            rows[3].cells[0].text = "Folder:"
            rows[3].cells[1].text = data.get('folder', 'Unknown')
            rows[4].cells[0].text = "Attachments:"
            attach_count = len(data.get('attachments', []))
            rows[4].cells[1].text = f"{attach_count} file(s)"
            
            # Add body if requested
            if include_body and 'body' in data:
                doc.add_paragraph()
                doc.add_heading("Body:", level=3)
                body_para = doc.add_paragraph(data['body'])
                # Set to monospace for readability
                for run in body_para.runs:
                    run.font.name = 'Courier New'
                    run.font.size = Pt(10)
            
            # Add attachment list if present
            if data.get('attachments'):
                doc.add_paragraph()
                doc.add_heading("Attachments:", level=3)
                for attach in data['attachments']:
                    if attach.get('error'):
                        doc.add_paragraph(f"⚠ {attach['filename']} (ERROR: {attach['error']})", style='List Bullet')
                    elif attach.get('path'):
                        size_kb = attach['size'] // 1024 if attach['size'] > 0 else 0
                        doc.add_paragraph(f"{attach['filename']} ({size_kb} KB)", style='List Bullet')
                    else:
                        doc.add_paragraph(f"{attach['filename']}", style='List Bullet')
            
            # Page break between messages (unless it's the last one)
            if i < len(messages):
                doc.add_page_break()
        
        doc.save(str(output_path))
        print(f"✓ Exported to {output_path}")
    except Exception as e:
        print(f"ERROR: Failed to export to Word: {e}")
        sys.exit(1)


def export_to_markdown(messages: list, output_path: Path, include_body: bool = False, attachments_dir: Path = None) -> None:
    """
    Export messages to a Markdown (.md) file.
    
    Think of this like piping to Out-File, but with Markdown formatting
    (# for headings, ** for bold, etc).
    
    Args:
        messages: List of dicts (extracted message data)
        output_path: Where to save the .md file
        include_body: Whether to include full message bodies
        attachments_dir: If provided, list downloaded attachments
    """
    try:
        lines = []
        lines.append("# Email Export\n")
        lines.append(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        lines.append(f"**Message count:** {len(messages)}\n")
        lines.append("---\n")
        
        for i, data in enumerate(messages, start=1):
            # Markdown heading (## = h2)
            lines.append(f"## Message {i}: {data['subject']}\n")
            lines.append(f"**From:** {data['sender']}\n")
            lines.append(f"**To:** {data['to']}\n")
            lines.append(f"**Date:** {data['date']}\n")
            lines.append(f"**Folder:** {data.get('folder', 'Unknown')}\n")
            
            if data.get('attachments'):
                attach_count = len(data['attachments'])
                lines.append(f"**Attachments:** {attach_count}\n")
                for attach in data['attachments']:
                    if attach.get('error'):
                        lines.append(f"  - ⚠ {attach['filename']} (ERROR)\n")
                    elif attach.get('path'):
                        size_kb = attach['size'] // 1024 if attach['size'] > 0 else 0
                        lines.append(f"  - {attach['filename']} ({size_kb} KB)\n")
                    else:
                        lines.append(f"  - {attach['filename']}\n")
            else:
                lines.append("**Attachments:** None\n")
            
            if include_body and 'body' in data:
                lines.append("\n### Body:\n")
                # Wrap body in code block for readability
                lines.append("```\n")
                lines.append(data['body'])
                lines.append("\n```\n")
            
            lines.append("\n---\n")
        
        output_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"✓ Exported to {output_path}")
    except Exception as e:
        print(f"ERROR: Failed to export to Markdown: {e}")
        sys.exit(1)


def export_to_pdf(messages: list, output_path: Path, include_body: bool = False, attachments_dir: Path = None) -> None:
    """
    Export messages to a PDF file using reportlab.
    
    Similar to Word export, but generates PDF via a Platypus document
    (reportlab's high-level layout engine).
    
    Args:
        messages: List of dicts (extracted message data)
        output_path: Where to save the .pdf file
        include_body: Whether to include full message bodies
        attachments_dir: If provided, list downloaded attachments
    """
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
        from reportlab.lib.units import inch
        from reportlab.lib import colors
        
        # Create PDF document
        doc = SimpleDocTemplate(str(output_path), pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        # Title
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#0066cc'),
            spaceAfter=12,
            alignment=1  # Center
        )
        story.append(Paragraph("Email Export", title_style))
        
        # Metadata
        meta_style = styles['Normal']
        story.append(Paragraph(f"<b>Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", meta_style))
        story.append(Paragraph(f"<b>Message count:</b> {len(messages)}", meta_style))
        story.append(Spacer(1, 0.3 * inch))
        
        # Process messages
        for i, data in enumerate(messages, start=1):
            # Message heading
            story.append(Paragraph(f"<b>Message {i}: {data['subject']}</b>", styles['Heading2']))
            
            # Metadata table
            table_data = [
                ["From:", data['sender']],
                ["To:", data['to']],
                ["Date:", data['date']],
                ["Folder:", data.get('folder', 'Unknown')],
            ]
            
            # Add attachment count if any
            if data.get('attachments'):
                attach_count = len(data['attachments'])
                table_data.append(["Attachments:", f"{attach_count} file(s)"])
            
            table = Table(table_data, colWidths=[1.2*inch, 4.3*inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            story.append(table)
            
            # Body if requested
            if include_body and 'body' in data:
                story.append(Spacer(1, 0.2 * inch))
                story.append(Paragraph("<b>Body:</b>", styles['Heading3']))
                # Escape HTML entities in body
                body_escaped = data['body'].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                story.append(Paragraph(body_escaped, styles['Normal']))
            
            # List attachments if present
            if data.get('attachments'):
                story.append(Spacer(1, 0.2 * inch))
                story.append(Paragraph("<b>Attachments:</b>", styles['Heading3']))
                for attach in data['attachments']:
                    if attach.get('error'):
                        story.append(Paragraph(f"⚠ {attach['filename']} (ERROR)", styles['Normal']))
                    elif attach.get('path'):
                        size_kb = attach['size'] // 1024 if attach['size'] > 0 else 0
                        story.append(Paragraph(f"• {attach['filename']} ({size_kb} KB)", styles['Normal']))
                    else:
                        story.append(Paragraph(f"• {attach['filename']}", styles['Normal']))
            
            # Page break (except after last message)
            if i < len(messages):
                story.append(PageBreak())
        
        doc.build(story)
        print(f"✓ Exported to {output_path}")
    except Exception as e:
        print(f"ERROR: Failed to export to PDF: {e}")
        sys.exit(1)


def main():
    """
    Parse CLI arguments and orchestrate the extraction.
    
    This is the entrypoint: like running a PowerShell script with parameters.
    """
    # Argument parsing (manual; similar to param() blocks in PS)
    args = sys.argv[1:]
    
    # Defaults (like $PSBoundParameters in PS)
    email = None
    contact = None
    output_format = "docx"
    output_file = None
    download_dir = DEFAULT_DOWNLOAD_DIR
    attachments_dir = None
    include_body = False
    folder_names = None
    limit = None
    force_overwrite = False
    skip_confirmation = False
    use_oauth2 = False
    
    # Parse positional and flag arguments (like -Name, -Value in PS)
    i = 0
    while i < len(args):
        arg = args[i]
        
        if arg in ("--format", "-f"):
            if i + 1 >= len(args):
                print("ERROR: --format requires a value (docx, pdf, markdown)")
                sys.exit(1)
            output_format = args[i + 1].lower()
            i += 2
            continue
        
        if arg in ("--output", "-o"):
            if i + 1 >= len(args):
                print("ERROR: --output requires a file path")
                sys.exit(1)
            output_file = Path(args[i + 1]).expanduser()
            i += 2
            continue
        
        if arg in ("--download-dir", "-d"):
            if i + 1 >= len(args):
                print("ERROR: --download-dir requires a folder path")
                sys.exit(1)
            download_dir = Path(args[i + 1]).expanduser()
            i += 2
            continue
        
        if arg in ("--attachments", "-a"):
            if i + 1 >= len(args):
                print("ERROR: --attachments requires a folder path")
                sys.exit(1)
            attachments_dir = Path(args[i + 1]).expanduser()
            i += 2
            continue
        
        if arg in ("--folders",):
            if i + 1 >= len(args):
                print("ERROR: --folders requires a comma-separated list (e.g., Inbox,Sent,Archive)")
                sys.exit(1)
            folder_names = [f.strip() for f in args[i + 1].split(",")]
            i += 2
            continue
        
        if arg in ("--limit", "-l"):
            if i + 1 >= len(args):
                print("ERROR: --limit requires a number")
                sys.exit(1)
            try:
                limit = int(args[i + 1])
            except ValueError:
                print(f"ERROR: --limit requires an integer, got '{args[i + 1]}'")
                sys.exit(1)
            i += 2
            continue
        
        if arg in ("--include-body", "-b"):
            include_body = True
            i += 1
            continue
        
        if arg in ("--oauth2",):
            use_oauth2 = True
            i += 1
            continue
        
        if arg in ("--yes", "-y"):
            skip_confirmation = True
            i += 1
            continue
        
        if arg in ("--force",):
            force_overwrite = True
            i += 1
            continue
        
        if arg.startswith("--"):
            print(f"ERROR: Unknown option: {arg}")
            sys.exit(1)
        
        # Positional arguments
        if not email:
            email = arg
        elif not contact:
            contact = arg
        else:
            print(f"ERROR: Unexpected argument: {arg}")
            sys.exit(1)
        
        i += 1
    
    # Validate required arguments
    if not email or not contact:
        print("ERROR: Missing required arguments")
        print("Usage: python3 extract_365_emails.py <your-email> <contact-email> [options]")
        print("\nOptions:")
        print("  --format FORMAT       docx (default), pdf, or markdown")
        print("  --output FILE         output filename (default: extract_<contact>_<date>.<ext>)")
        print("  --download-dir PATH   output directory (default: ~/Downloads)")
        print("  --attachments DIR     where to save attachments (default: attachments/ in download-dir)")
        print("  --folders FOLDER,...  comma-separated folder names (default: Inbox,Sent,Archive)")
        print("  --include-body        include full message bodies (default: metadata only)")
        print("  --limit N             extract only first N messages")
        print("  --oauth2              use OAuth2 with browser login (default: app password prompt)")
        print("  --yes, -y             skip confirmation prompt and proceed automatically")
        print("  --force               overwrite existing files")
        sys.exit(1)
    
    # Validate format
    if output_format not in ("docx", "pdf", "markdown"):
        print(f"ERROR: Invalid format '{output_format}'. Must be docx, pdf, or markdown")
        sys.exit(1)
    
    # Determine output file
    if not output_file:
        contact_clean = sanitize_email(contact)
        date_str = datetime.now().strftime("%Y%m%d")
        ext = "md" if output_format == "markdown" else output_format
        output_file = download_dir / f"extract_{contact_clean}_{date_str}.{ext}"
    
    # Default attachments directory
    if attachments_dir is None:
        attachments_dir = download_dir / "attachments"
    
    # Check if file exists
    if output_file.exists() and not force_overwrite:
        print(f"ERROR: Output file already exists: {output_file}")
        print("Use --force to overwrite")
        sys.exit(1)
    
    # Show confirmation prompt before starting work (unless --yes is used)
    print("About to extract emails with these settings:")
    print(f"  Your email:     {email}")
    print(f"  Contact:        {contact}")
    print(f"  Auth method:    {'OAuth2 (browser login)' if use_oauth2 else 'Basic auth (app password)'}")
    print(f"  Format:         {output_format.upper()}")
    print(f"  Folders:        {', '.join(folder_names) if folder_names else '(all accessible)'}")
    print(f"  Include bodies: {'yes' if include_body else 'no'}")
    print(f"  Message limit:  {limit if limit else '(none — all)'}")
    print(f"  Output file:    {output_file}")
    print(f"  Attachments:    {attachments_dir}")
    print()
    
    if not skip_confirmation:
        response = input("Proceed? [Y/n]: ").strip().lower()
        if response == "n":
            print("Aborted — nothing done.")
            sys.exit(0)
    
    # Create directories if needed
    download_dir.mkdir(parents=True, exist_ok=True)
    attachments_dir.mkdir(parents=True, exist_ok=True)
    
    # Main workflow (like a PowerShell pipeline: connect -> query -> extract -> export)
    print()
    account = connect_to_exchange(email, use_oauth2=use_oauth2)
    messages = query_messages(account, contact, folder_names=folder_names, limit=limit)
    
    if not messages:
        print("No messages found. Exiting.")
        sys.exit(0)
    
    # Extract message data (this downloads attachments too)
    print("\nExtracting message data...")
    extracted_messages = []
    for i, msg in enumerate(messages, start=1):
        data = extract_message_data(msg, include_body=include_body, attachments_dir=attachments_dir, msg_index=i)
        extracted_messages.append(data)
        print(f"  [{i}/{len(messages)}] {data['subject'][:60]}", end="\r")
    print()  # Newline after progress
    
    # Export based on format
    print(f"\nExporting as {output_format.upper()}...")
    if output_format == "docx":
        export_to_docx(extracted_messages, output_file, include_body=include_body, attachments_dir=attachments_dir)
    elif output_format == "pdf":
        export_to_pdf(extracted_messages, output_file, include_body=include_body, attachments_dir=attachments_dir)
    elif output_format == "markdown":
        export_to_markdown(extracted_messages, output_file, include_body=include_body, attachments_dir=attachments_dir)
    
    size_kb = output_file.stat().st_size // 1024
    print(f"\nDone! {len(extracted_messages)} messages exported.")
    print(f"File: {output_file} ({size_kb} KB)")
    
    # Report on attachments if any
    if attachments_dir.exists() and any(attachments_dir.iterdir()):
        attachment_count = sum(1 for _ in attachments_dir.rglob('*') if _.is_file())
        print(f"Attachments: {attachment_count} file(s) saved to {attachments_dir}")


if __name__ == "__main__":
    main()