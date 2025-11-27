const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// Drafts folder paths
const DRAFTS_FOLDER = path.join(__dirname, '..', 'drafts');
const ATTACHMENTS_FOLDER = path.join(DRAFTS_FOLDER, 'attachments');

// Create folders if they don't exist
if (!fs.existsSync(DRAFTS_FOLDER)) {
  fs.mkdirSync(DRAFTS_FOLDER, { recursive: true });
}
if (!fs.existsSync(ATTACHMENTS_FOLDER)) {
  fs.mkdirSync(ATTACHMENTS_FOLDER, { recursive: true });
}

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' })); // Support for large files

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'OK', message: 'Backend is running!' });
});

// Draft creation endpoint
app.post('/create-drafts', async (req, res) => {
  try {
    const { recipients, subject, body, attachmentData, attachmentName } = req.body;

    // Validation
    if (!recipients || recipients.length === 0) {
      return res.status(400).json({ error: 'At least one recipient is required' });
    }

    if (!subject) {
      return res.status(400).json({ error: 'Subject is required' });
    }

    let attachmentPath = null;

    // Save attachment if exists
    if (attachmentData && attachmentName) {
      const timestamp = Date.now();
      const safeFileName = attachmentName.replace(/[^a-zA-Z0-9._-]/g, '_');
      const fileName = `${timestamp}_${safeFileName}`;
      attachmentPath = path.join(ATTACHMENTS_FOLDER, fileName);

      // Convert Base64 to file
      const base64Data = attachmentData.replace(/^data:.*?;base64,/, '');
      fs.writeFileSync(attachmentPath, base64Data, 'base64');
      
      console.log(`📎 File saved: ${attachmentPath}`);
    }

    // Create JSON file for each recipient
    const draftFiles = [];
    for (let i = 0; i < recipients.length; i++) {
      const recipient = recipients[i];
      const timestamp = Date.now();
      const draftFileName = `draft_${timestamp}_${i}.json`;
      const draftFilePath = path.join(DRAFTS_FOLDER, draftFileName);

      const draftData = {
        to: recipient,
        subject: subject,
        body: body,
        attachmentPath: attachmentPath,
        attachmentName: attachmentName,
        createdAt: new Date().toISOString()
      };

      fs.writeFileSync(draftFilePath, JSON.stringify(draftData, null, 2), 'utf-8');
      draftFiles.push(draftFileName);
      
      console.log(`✉️ Draft created: ${draftFileName} for ${recipient}`);
    }

    res.json({
      success: true,
      draftsCreated: recipients.length,
      files: draftFiles,
      message: 'Drafts created successfully!'
    });

  } catch (error) {
    console.error('❌ Error creating drafts:', error);
    res.status(500).json({ error: 'Internal server error', details: error.message });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`🚀 Backend Server running on http://localhost:${PORT}`);
  console.log(`📁 Drafts folder: ${DRAFTS_FOLDER}`);
  console.log(`📎 Attachments folder: ${ATTACHMENTS_FOLDER}`);
});