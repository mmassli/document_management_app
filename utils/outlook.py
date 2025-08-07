"""
Outlook integration utilities and helpers.
"""

# Try to import Outlook COM interface and test connection
try:
    import win32com.client
    try:
        # Test Outlook connection immediately
        outlook_test = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook_test.GetNamespace("MAPI")
        import logging
        logging.info("Outlook COM connection successful.")
        OUTLOOK_AVAILABLE = True
    except Exception as e:
        OUTLOOK_AVAILABLE = False
        import logging
        logging.exception("Failed to connect to Outlook:")
except ImportError:
    OUTLOOK_AVAILABLE = False
    import logging
    logging.error("pywin32 not installed.")


def get_outlook_connection():
    """Get connection to Outlook application"""
    if not OUTLOOK_AVAILABLE:
        raise ImportError("Outlook COM interface not available")

    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook


def get_inbox_messages(outlook, limit=50):
    """Get recent messages from Outlook inbox"""
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # Sort by received time, newest first

    return messages


def format_file_size(size_bytes):
    """Format file size in human readable format"""
    if size_bytes == 0:
        return "0 B"
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"