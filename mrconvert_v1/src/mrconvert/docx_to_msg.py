from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional

from .types import ConversionResult, ConversionError, EngineNotFoundError


def docx_to_msg(
    src: Path,
    dst: Path,
    subject: Optional[str] = None,
    sender: Optional[str] = None,
    recipients: Optional[list[str]] = None
) -> ConversionResult:
    """
    Convert DOCX file to Outlook MSG format
    
    Args:
        src: Source DOCX file path
        dst: Destination MSG file path
        subject: Email subject line
        sender: Sender email address
        recipients: List of recipient email addresses
        
    Returns:
        ConversionResult with conversion details
        
    Note:
        This function requires Windows and Microsoft Outlook installed.
        Uses win32com (pywin32) for Outlook COM interface.
    """
    if not src.exists():
        raise FileNotFoundError(f"Source file not found: {src}")
    
    # Check if running on Windows
    if sys.platform != "win32":
        raise EngineNotFoundError(
            "MSG conversion is only supported on Windows with Microsoft Outlook installed. "
            "Please use DOCX format instead."
        )
    
    try:
        import win32com.client
    except ImportError:
        raise EngineNotFoundError(
            "pywin32 is required for MSG conversion. "
            "Install it with: pip install pywin32"
        )
    
    try:
        # Initialize Outlook COM interface
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Create a new mail item
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        
        # Set email properties
        if subject:
            mail.Subject = subject
        else:
            mail.Subject = src.stem
        
        if sender:
            mail.SenderEmailAddress = sender
        
        if recipients:
            for recipient in recipients:
                mail.Recipients.Add(recipient)
        
        # Read DOCX content and convert to HTML
        # For simplicity, we'll attach the DOCX file and reference it
        # Or we can convert DOCX to HTML body
        
        # Attach DOCX file as attachment
        mail.Attachments.Add(str(src.absolute()))
        
        # Set body text
        mail.Body = f"Please see attached document: {src.name}"
        mail.BodyFormat = 2  # olFormatHTML (2) or olFormatPlain (1)
        
        # Save as MSG file
        dst.parent.mkdir(parents=True, exist_ok=True)
        mail.SaveAs(str(dst.absolute()), 3)  # 3 = olMSG
        
        return ConversionResult(src, dst, "win32com-outlook")
        
    except Exception as e:
        raise ConversionError(f"Failed to create MSG file: {e}")


def markdown_to_msg(
    src: Path,
    dst: Path,
    metadata: Optional[dict] = None
) -> ConversionResult:
    """
    Convert markdown email file directly to Outlook MSG format
    
    Args:
        src: Source markdown file path
        dst: Destination MSG file path
        metadata: Optional metadata dict with subject, sender, recipients, etc.
        
    Returns:
        ConversionResult with conversion details
    """
    if sys.platform != "win32":
        raise EngineNotFoundError(
            "MSG conversion is only supported on Windows with Microsoft Outlook installed."
        )
    
    try:
        import win32com.client
    except ImportError:
        raise EngineNotFoundError(
            "pywin32 is required for MSG conversion. Install with: pip install pywin32"
        )
    
    # Read markdown content
    content = src.read_text(encoding='utf-8')
    
    # Extract metadata if available
    if metadata is None:
        from .markdown_to_docx import extract_json_ld_from_markdown
        metadata = extract_json_ld_from_markdown(content)
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        
        # Set subject from metadata or filename
        if metadata and 'subject' in metadata:
            mail.Subject = metadata['subject']
        else:
            mail.Subject = src.stem
        
        # Note: SenderEmailAddress cannot be set directly in Outlook COM
        # It's automatically set to the current Outlook account
        
        # Add recipients from metadata
        # Note: Recipient.Name cannot be set directly, Outlook will resolve it from email
        if metadata and 'participants' in metadata:
            for participant in metadata['participants']:
                if 'email' in participant:
                    # Add as recipient (Outlook will resolve name from address book)
                    mail.Recipients.Add(participant['email'])
        
        # Convert markdown content to plain text for body
        # Remove JSON code block
        import re
        body_text = re.sub(r'```json\s*\{.*?\}\s*```', '', content, flags=re.DOTALL)
        # Remove markdown headers
        body_text = re.sub(r'^#+\s+', '', body_text, flags=re.MULTILINE)
        # Clean up
        body_text = body_text.strip()
        
        # Set body
        mail.Body = body_text[:50000]  # Limit body length
        mail.BodyFormat = 1  # olFormatPlain
        
        # Save as MSG
        dst.parent.mkdir(parents=True, exist_ok=True)
        mail.SaveAs(str(dst.absolute()), 3)
        
        return ConversionResult(src, dst, "win32com-outlook-markdown")
        
    except Exception as e:
        raise ConversionError(f"Failed to create MSG file from markdown: {e}")

