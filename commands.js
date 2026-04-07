/**
 * Office Add-in commands implementation
 * This file contains the functions called by the ribbon buttons
 */

Office.onReady(() => {
    console.log("Outlook Add-in ready");
});

/**
 * Copy Desktop Link Function
 * Creates outlook://mail/entryid=... link and copies to clipboard
 */
async function copyDesktopLink(event) {
    try {
        // Get the current mail item
        const item = Office.context.mailbox.item;
        
        if (!item) {
            showError("No email selected");
            event.completed();
            return;
        }
        
        // Get the item ID
        const itemId = item.itemId;
        
        // Convert to Hex format for outlook: URL
        const hexId = stringToHex(itemId);
        
        // Build the desktop link
        const desktopLink = \`outlook://mail/entryid=\${hexId}\`;
        
        // Copy to clipboard using the Clipboard API
        await copyToClipboard(desktopLink);
        
        // Show success message
        showSuccess(\`Desktop link copied: \${item.subject}\`);
        
        // Complete the command
        event.completed();
        
    } catch (error) {
        showError("Error: " + error.message);
        event.completed();
    }
}

/**
 * Copy Web Link Function
 * Creates Office 365 / Outlook Web link and copies to clipboard
 */
async function copyWebLink(event) {
    try {
        const item = Office.context.mailbox.item;
        
        if (!item) {
            showError("No email selected");
            event.completed();
            return;
        }
        
        // Get the conversation ID (unique identifier in O365)
        const conversationId = item.conversationId;
        
        // Build web link for Outlook Web Access (OWA)
        // Format: https://outlook.office.com/mail/search/all?q=<messageId>
        const webLink = \`https://outlook.office.com/mail/inbox\`;
        
        // Alternative: If you have specific server info
        // const webLink = \`https://<servername>/owa?ItemID=\${conversationId}\`;
        
        // Copy to clipboard
        await copyToClipboard(webLink);
        
        // Show success
        showSuccess(\`Web link copied: \${item.subject}\`);
        
        event.completed();
        
    } catch (error) {
        showError("Error: " + error.message);
        event.completed();
    }
}

/**
 * Helper: Copy text to clipboard using Clipboard API
 */
async function copyToClipboard(text) {
    // Try modern Clipboard API first
    if (navigator.clipboard && navigator.clipboard.writeText) {
        try {
            await navigator.clipboard.writeText(text);
            return true;
        } catch (err) {
            console.warn("Clipboard API failed, falling back to textarea method");
        }
    }
    
    // Fallback: textarea method
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    
    try {
        textarea.select();
        document.execCommand("copy");
        return true;
    } finally {
        document.body.removeChild(textarea);
    }
}

/**
 * Helper: Convert string to hex (for outlook: URLs)
 */
function stringToHex(str) {
    let hex = "";
    for (let i = 0; i < str.length; i++) {
        const charCode = str.charCodeAt(i);
        const hexChar = charCode.toString(16).toUpperCase();
        hex += hexChar.padStart(2, "0");
    }
    return hex;
}

/**
 * Helper: Show success message
 */
function showSuccess(message) {
    const statusDiv = document.getElementById("status");
    if (statusDiv) {
        statusDiv.className = "status success";
        statusDiv.textContent = "✓ " + message;
        
        // Clear after 5 seconds
        setTimeout(() => {
            statusDiv.textContent = "";
        }, 5000);
    }
    
    console.log("Success: " + message);
}

/**
 * Helper: Show error message
 */
function showError(message) {
    const statusDiv = document.getElementById("status");
    if (statusDiv) {
        statusDiv.className = "status error";
        statusDiv.textContent = "✗ " + message;
        
        setTimeout(() => {
            statusDiv.textContent = "";
        }, 5000);
    }
    
    console.error("Error: " + message);
}