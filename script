// Google Apps Script for Donation Certificate Generator
// Follow the setup steps in order!

// ============================================
// STEP 1: RUN THIS FIRST - Create Certificate Template
// ============================================
function Step1_CreateCertificateTemplate() {
  try {
    // Create a new Google Doc for the certificate template
    const doc = DocumentApp.create('Donation Certificate Template');
    const body = doc.getBody();
    
    // Clear default content
    body.clear();
    
    // Set page margins (optional)
    body.setMarginTop(72);    // 1 inch
    body.setMarginBottom(72);  // 1 inch
    body.setMarginLeft(72);    // 1 inch
    body.setMarginRight(72);   // 1 inch
    
    // Add certificate content with placeholders
    // Add some decorative spacing at top
    body.appendParagraph('').setSpacingAfter(20);
    
    // Title
    const title = body.appendParagraph('CERTIFICATE OF DONATION');
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    title.setFontSize(28);
    title.setBold(true);
    title.setForegroundColor('#667eea');
    title.setSpacingAfter(20);
    
    // Decorative line
    const line1 = body.appendParagraph('━━━━━━━━━━━━━━━━━━━━━━━━━');
    line1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    line1.setForegroundColor('#764ba2');
    line1.setSpacingAfter(20);
    
    // Subtitle
    const subtitle = body.appendParagraph('This is to certify that');
    subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    subtitle.setFontSize(14);
    subtitle.setItalic(true);
    subtitle.setSpacingAfter(15);
    
    // Donor name (placeholder)
    const donorName = body.appendParagraph('{{FULL_NAME}}');
    donorName.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    donorName.setFontSize(22);
    donorName.setBold(true);
    donorName.setForegroundColor('#333333');
    donorName.setSpacingAfter(15);
    
    // Social media handle
    const social = body.appendParagraph('( {{SOCIAL_MEDIA}} )');
    social.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    social.setFontSize(11);
    social.setItalic(true);
    social.setForegroundColor('#666666');
    social.setSpacingAfter(20);
    
    // Description
    const description = body.appendParagraph('has made a generous donation in support of AESPA and our cause');
    description.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    description.setFontSize(13);
    description.setSpacingAfter(25);
    
    // Donation details box
    const amountLabel = body.appendParagraph('Donation Amount');
    amountLabel.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    amountLabel.setFontSize(11);
    amountLabel.setForegroundColor('#666666');
    amountLabel.setSpacingAfter(5);
    
    const amount = body.appendParagraph('{{AMOUNT}}');
    amount.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    amount.setFontSize(18);
    amount.setBold(true);
    amount.setForegroundColor('#667eea');
    amount.setSpacingAfter(20);
    
    // Date
    const date = body.appendParagraph('Given on {{DATE}}');
    date.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    date.setFontSize(12);
    date.setSpacingAfter(25);
    
    // Message section (if provided)
    const messageTitle = body.appendParagraph('Message of Support');
    messageTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    messageTitle.setFontSize(12);
    messageTitle.setBold(true);
    messageTitle.setForegroundColor('#764ba2');
    messageTitle.setSpacingAfter(10);
    
    const message = body.appendParagraph('"{{MESSAGE}}"');
    message.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    message.setFontSize(11);
    message.setItalic(true);
    message.setForegroundColor('#555555');
    message.setSpacingAfter(30);
    
    // Decorative line
    const line2 = body.appendParagraph('━━━━━━━━━━━━━━━━━━━━━━━━━');
    line2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    line2.setForegroundColor('#764ba2');
    line2.setSpacingAfter(15);
    
    // Certificate number
    const certNumber = body.appendParagraph('Certificate No: {{CERTIFICATE_NUMBER}}');
    certNumber.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    certNumber.setFontSize(10);
    certNumber.setItalic(true);
    certNumber.setForegroundColor('#999999');
    certNumber.setSpacingAfter(10);
    
    // Footer
    const footer = body.appendParagraph('Thank you for supporting AESPA!');
    footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    footer.setFontSize(14);
    footer.setBold(true);
    footer.setForegroundColor('#667eea');
    
    // Save and close
    doc.saveAndClose();
    
    // Log the results
    const templateId = doc.getId();
    const templateUrl = doc.getUrl();
    
    Logger.log('✅ Certificate template created successfully!');
    Logger.log('📋 Template ID: ' + templateId);
    Logger.log('🔗 Template URL: ' + templateUrl);
    Logger.log('');
    Logger.log('⚠️ IMPORTANT: Copy the Template ID above and paste it in Step 2!');
    
    // Also return the ID for convenience
    return templateId;
    
  } catch (error) {
    Logger.log('❌ Error creating template: ' + error.toString());
    throw error;
  }
}

// ============================================
// STEP 2: UPDATE THIS CONFIGURATION
// ============================================
const CONFIG = {
  // PASTE YOUR TEMPLATE ID HERE (from Step 1)
  TEMPLATE_DOC_ID: '10BOujxNJT1ALSCLKM5rwWwAg1d34zyHOnvrjwhnqXhs',
  
  // OPTIONAL: Create a folder in Google Drive and paste its ID here
  // To get folder ID: Open the folder in Drive, look at the URL after /folders/
  CERTIFICATES_FOLDER_ID: '1ExtKfdMgwgDh4zTX6zJBEy93KwU8ghJD',  // Leave empty if you don't want to organize certificates
  
  // Email configuration
  EMAIL_SUBJECT: 'Thank You for Your Donation to AESPA - Certificate Attached',
  SENDER_NAME: 'AESPA Support Team',
  
  // Certificate numbering prefix
  CERTIFICATE_PREFIX: 'AESPA-2025-',
  
  // Admin notification settings
  ADMIN_EMAIL: Session.getActiveUser().getEmail(),  // Automatically uses your email
  ENABLE_ADMIN_NOTIFICATIONS: true,  // Set to false to disable admin notifications
  ADMIN_EMAIL_SUBJECT: 'New Donation Received for AESPA'
};

// ============================================
// STEP 3: RUN THIS - Create and Setup Form
// ============================================
function Step3_CreateAndSetupForm() {
  try {
    // Check if template ID is configured
    if (CONFIG.TEMPLATE_DOC_ID === '10BOujxNJT1ALSCLKM5rwWwAg1d34zyHOnvrjwhnqXhs') {
      throw new Error('Please update TEMPLATE_DOC_ID in CONFIG first! Run Step 1 to get the ID.');
    }
    
    // Create a new form
    const form = FormApp.create('AESPA Donation Form');
    
    // Customize form settings
    form.setDescription('Thank you for supporting AESPA! Please fill out this form to receive your donation certificate.');
    form.setConfirmationMessage('Thank you for your donation! You will receive your certificate via email shortly.');
    form.setCollectEmail(true);  // Automatically collect email addresses
    
    // Add form fields
    const nameItem = form.addTextItem();
    nameItem.setTitle('Full Name')
      .setHelpText('Please enter your full name as you want it to appear on the certificate')
      .setRequired(true);
    
    const socialItem = form.addTextItem();
    socialItem.setTitle('Social Media Handle')
      .setHelpText('Your Instagram, Twitter, or other social media username')
      .setRequired(true);
    
    const amountItem = form.addTextItem();
    amountItem.setTitle('Donation Amount')
      .setHelpText('Please enter the amount you donated (e.g., $50 or 50 USD)')
      .setRequired(true);
    
    const messageItem = form.addParagraphTextItem();
    messageItem.setTitle('Message for AESPA (Optional)')
      .setHelpText('Leave a message of support for AESPA (this will appear on your certificate)')
      .setRequired(false);
    
    // Receipt upload - using URL for compatibility
    const receiptItem = form.addTextItem();
    receiptItem.setTitle('Transfer Receipt Link')
      .setHelpText('Please provide a link to your payment receipt (Google Drive, Dropbox, Imgur, etc.)')
      .setRequired(true)
      .setValidation(
        FormApp.createTextValidation()
          .requireTextIsUrl()
          .setHelpText('Please enter a valid URL')
          .build()
      );
    
    // Get the form trigger
    const formId = form.getId();
    
    // Set up form submit trigger
    ScriptApp.newTrigger('onFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();
    
    // Get URLs
    const editUrl = form.getEditUrl();
    const publishedUrl = form.getPublishedUrl();
    
    // Use console.log for better visibility
    console.log('✅ FORM CREATED SUCCESSFULLY!');
    console.log('=====================================');
    console.log('📝 FORM EDIT URL:');
    console.log(editUrl);
    console.log('=====================================');
    console.log('🔗 FORM PUBLIC URL (Share this with donors):');
    console.log(publishedUrl);
    console.log('=====================================');
    console.log('📄 FORM ID: ' + formId);
    console.log('=====================================');
    
    // Also use Logger for backup
    Logger.log('✅ Form created and trigger set up successfully!');
    Logger.log('📝 Form Edit URL: ' + editUrl);
    Logger.log('🔗 Form Public URL: ' + publishedUrl);
    Logger.log('Form ID: ' + formId);
    
    // Save URLs to Script Properties for later retrieval
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('FORM_EDIT_URL', editUrl);
    scriptProperties.setProperty('FORM_PUBLIC_URL', publishedUrl);
    scriptProperties.setProperty('FORM_ID', formId);
    
    // Create a summary message
    const summary = `
FORM CREATED SUCCESSFULLY!
==========================
Form Name: AESPA Donation Form
Form ID: ${formId}

📝 TO EDIT YOUR FORM:
${editUrl}

🔗 TO SHARE WITH DONORS:
${publishedUrl}

The form is now connected to the certificate generator!
Test it by submitting a response.
    `;
    
    console.log(summary);
    
    return {
      editUrl: editUrl,
      publishedUrl: publishedUrl,
      formId: formId,
      summary: summary
    };
    
  } catch (error) {
    console.log('❌ ERROR: ' + error.toString());
    Logger.log('❌ Error creating form: ' + error.toString());
    throw error;
  }
}

// ============================================
// MAIN FUNCTION - Triggered on Form Submit
// ============================================
function onFormSubmit(e) {
  try {
    Logger.log('📥 New form submission received');
    
    // Check configuration
    if (!CONFIG.TEMPLATE_DOC_ID || CONFIG.TEMPLATE_DOC_ID === 'PASTE_YOUR_TEMPLATE_ID_HERE') {
      throw new Error('Template ID not configured. Please run setup steps first.');
    }
    
    // Get form responses
    const responses = e.response;
    const itemResponses = responses.getItemResponses();
    
    // Extract donor information
    const donorData = {
      email: responses.getRespondentEmail(),
      fullName: '',
      socialMedia: '',
      donationAmount: '',
      message: 'Thank you for your support!',
      receiptUrl: '',
      donationDate: new Date().toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
      }),
      certificateNumber: generateCertificateNumber(),
      timestamp: new Date().toLocaleString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      })
    };
    
    // Parse responses by question title
    for (let i = 0; i < itemResponses.length; i++) {
      const response = itemResponses[i];
      const title = response.getItem().getTitle();
      const answer = response.getResponse();
      
      if (title.includes('Full Name')) {
        donorData.fullName = answer;
      } else if (title.includes('Social Media')) {
        donorData.socialMedia = answer;
      } else if (title.includes('Donation Amount')) {
        donorData.donationAmount = answer;
      } else if (title.includes('Message')) {
        donorData.message = answer || 'Thank you for your support!';
      } else if (title.includes('Receipt')) {
        // Handle receipt URL - fix for file uploads or invalid URLs
        donorData.receiptUrl = processReceiptUrl(answer);
      }
    }
    
    Logger.log('👤 Processing certificate for: ' + donorData.fullName);
    
    // Generate the certificate
    const certificateFile = createCertificate(donorData);
    
    // Send certificate to donor
    sendCertificateEmail(donorData, certificateFile);
    Logger.log('✅ Certificate sent to donor: ' + donorData.email);
    
    // Send admin notification if enabled
    if (CONFIG.ENABLE_ADMIN_NOTIFICATIONS) {
      try {
        sendAdminNotification(donorData);
        Logger.log('📧 Admin notification sent to: ' + CONFIG.ADMIN_EMAIL);
      } catch (adminError) {
        Logger.log('❌ Failed to send admin notification: ' + adminError.toString());
      }
    } else {
      Logger.log('ℹ️ Admin notifications are disabled');
    }
    
    Logger.log('✅ All processing completed successfully');
    
  } catch (error) {
    Logger.log('❌ Error processing submission: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    
    // Try to notify admin
    try {
      notifyAdminOfError(error, e);
    } catch (notifyError) {
      Logger.log('Could not send admin notification: ' + notifyError.toString());
    }
  }
}

// ============================================
// NEW FUNCTION - Send Admin Notification
// ============================================
function sendAdminNotification(donorData) {
  // Create admin notification HTML
  const adminHtmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
      </style>
    </head>
    <body style="margin: 0; padding: 0; font-family: 'Inter', Arial, sans-serif; background-color: #f5f5f5;">
      <div style="max-width: 600px; margin: 20px auto; background-color: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
        
        <!-- Header -->
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 30px; text-align: center;">
          <h2 style="color: white; margin: 0; font-size: 24px;">
            New Donation Received!
          </h2>
        </div>
        
        <!-- Content -->
        <div style="padding: 30px;">
          <p style="color: #333; font-size: 16px; margin: 0 0 20px 0;">
            Great news! A new donation has been received and processed.
          </p>
          
          <!-- Donor Details Card -->
          <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
            <h3 style="color: #333; margin: 0 0 15px 0; font-size: 18px;">
              Donation Details
            </h3>
            
            <table style="width: 100%; font-size: 14px; border-collapse: collapse;">
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666; width: 40%;">Donor Name:</td>
                <td style="padding: 10px 0; color: #333; font-weight: 600;">${donorData.fullName}</td>
              </tr>
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666;">Email:</td>
                <td style="padding: 10px 0; color: #333;">
                  <a href="mailto:${donorData.email}" style="color: #667eea; text-decoration: none;">
                    ${donorData.email}
                  </a>
                </td>
              </tr>
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666;">Amount:</td>
                <td style="padding: 10px 0; color: #28a745; font-weight: 700; font-size: 16px;">
                  ${donorData.donationAmount}
                </td>
              </tr>
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666;">Social Media:</td>
                <td style="padding: 10px 0; color: #333;">${donorData.socialMedia}</td>
              </tr>
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666;">Certificate #:</td>
                <td style="padding: 10px 0; color: #333; font-family: monospace;">
                  ${donorData.certificateNumber}
                </td>
              </tr>
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666;">Date & Time:</td>
                <td style="padding: 10px 0; color: #333;">${donorData.timestamp}</td>
              </tr>
              <tr>
                <td style="padding: 10px 0; color: #666;">Receipt:</td>
                <td style="padding: 10px 0;">
                  ${donorData.receiptUrl.startsWith('http') ? 
                    `<a href="${donorData.receiptUrl}" target="_blank" style="color: #667eea; text-decoration: none;">
                      View Receipt →
                    </a>` : 
                    `<span style="color: #666; font-style: italic;">${donorData.receiptUrl}</span>`
                  }
                </td>
              </tr>
            </table>
          </div>
          
          <!-- Message from Donor -->
          ${donorData.message && donorData.message !== 'Thank you for your support!' ? `
          <div style="background: #fff3e0; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #ffc107;">
            <p style="color: #856404; font-weight: 600; margin: 0 0 8px 0; font-size: 14px;">
              Message from Donor:
            </p>
            <p style="color: #333; margin: 0; font-style: italic; font-size: 14px;">
              "${donorData.message}"
            </p>
          </div>
          ` : ''}
          
          <!-- Status -->
          <div style="background: #d4edda; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #28a745;">
            <p style="color: #155724; margin: 0; font-size: 14px;">
              ✅ <strong>Certificate Status:</strong> Successfully generated and sent to donor
            </p>
          </div>
          
          <!-- Quick Actions -->
          <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0;">
            <p style="color: #666; font-size: 13px; margin: 0;">
              <strong>Quick Actions:</strong><br>
              • Certificate has been automatically sent to the donor<br>
              • PDF copy saved in Google Drive<br>
              • You can reply directly to ${donorData.email} if needed
            </p>
          </div>
        </div>
        
        <!-- Footer -->
        <div style="background: #f8f9fa; padding: 15px; text-align: center; border-top: 1px solid #e0e0e0;">
          <p style="color: #999; font-size: 12px; margin: 0;">
            This is an automated notification from AESPA Donation System
          </p>
        </div>
      </div>
    </body>
    </html>
  `;
  
  // Create plain text version
  const adminPlainText = `
New Donation Received!

DONATION DETAILS:
-----------------
Donor Name: ${donorData.fullName}
Email: ${donorData.email}
Amount: ${donorData.donationAmount}
Social Media: ${donorData.socialMedia}
Certificate #: ${donorData.certificateNumber}
Date & Time: ${donorData.timestamp}
Receipt: ${donorData.receiptUrl}

${donorData.message && donorData.message !== 'Thank you for your support!' ? 
`Message from Donor:
"${donorData.message}"

` : ''}
STATUS: Certificate successfully generated and sent to donor.

---
This is an automated notification from AESPA Donation System
  `;
  
  // Send admin notification email (no attachments)
  GmailApp.sendEmail(
    CONFIG.ADMIN_EMAIL,
    CONFIG.ADMIN_EMAIL_SUBJECT + ' - ' + donorData.fullName,
    adminPlainText,
    {
      htmlBody: adminHtmlBody,
      name: CONFIG.SENDER_NAME,
      noReply: true
    }
  );
}

// Generate unique certificate number
function generateCertificateNumber() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let currentNumber = scriptProperties.getProperty('CERT_NUMBER') || '1000';
  currentNumber = parseInt(currentNumber) + 1;
  scriptProperties.setProperty('CERT_NUMBER', currentNumber.toString());
  return CONFIG.CERTIFICATE_PREFIX + currentNumber;
}

// Create certificate from template
function createCertificate(donorData) {
  // Open the template document
  const templateDoc = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
  
  // Create a copy of the template
  const certificateName = `Certificate_${donorData.certificateNumber}_${donorData.fullName.replace(/[^a-zA-Z0-9]/g, '_')}`;
  const certificateCopy = templateDoc.makeCopy(certificateName);
  
  // Move to certificates folder if specified
  if (CONFIG.CERTIFICATES_FOLDER_ID && CONFIG.CERTIFICATES_FOLDER_ID !== '') {
    try {
      const folder = DriveApp.getFolderById(CONFIG.CERTIFICATES_FOLDER_ID);
      certificateCopy.moveTo(folder);
    } catch (error) {
      Logger.log('Warning: Could not move to folder: ' + error.toString());
    }
  }
  
  // Open the copied document
  const doc = DocumentApp.openById(certificateCopy.getId());
  const body = doc.getBody();
  
  // Replace placeholders with actual data
  body.replaceText('{{FULL_NAME}}', donorData.fullName);
  body.replaceText('{{CERTIFICATE_NUMBER}}', donorData.certificateNumber);
  body.replaceText('{{DATE}}', donorData.donationDate);
  body.replaceText('{{AMOUNT}}', donorData.donationAmount);
  body.replaceText('{{SOCIAL_MEDIA}}', donorData.socialMedia);
  body.replaceText('{{MESSAGE}}', donorData.message);
  
  // Save and close the document
  doc.saveAndClose();
  
  // Convert to PDF
  const pdfBlob = certificateCopy.getAs('application/pdf');
  pdfBlob.setName(certificateName + '.pdf');
  
  // Create PDF in Drive
  const pdfFile = DriveApp.createFile(pdfBlob);
  
  // Move PDF to folder if specified
  if (CONFIG.CERTIFICATES_FOLDER_ID && CONFIG.CERTIFICATES_FOLDER_ID !== '') {
    try {
      const folder = DriveApp.getFolderById(CONFIG.CERTIFICATES_FOLDER_ID);
      pdfFile.moveTo(folder);
    } catch (error) {
      Logger.log('Warning: Could not move PDF to folder: ' + error.toString());
    }
  }
  
  // Delete the temporary Google Doc copy (keep only PDF)
  certificateCopy.setTrashed(true);
  
  return pdfFile;
}

// ============================================
// NEW FUNCTION - Send Admin Notification
// ============================================
function sendAdminNotification(donorData) {
  // Create admin notification HTML
  const adminHtmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
      </style>
    </head>
    <body style="margin: 0; padding: 0; font-family: 'Inter', Arial, sans-serif; background-color: #f5f5f5;">
      <div style="max-width: 600px; margin: 20px auto; background-color: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
        
        <!-- Header -->
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 30px; text-align: center;">
          <h2 style="color: white; margin: 0; font-size: 24px;">
            🎉 New Donation Received!
          </h2>
        </div>
        
        <!-- Content -->
        <div style="padding: 30px;">
          <p style="color: #333; font-size: 16px; margin: 0 0 20px 0;">
            Great news! A new donation has been received and processed.
          </p>
          
          <!-- Donor Details Card -->
          <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
            <h3 style="color: #333; margin: 0 0 15px 0; font-size: 18px;">
              📋 Donation Details
            </h3>
            
            <table style="width: 100%; font-size: 14px; border-collapse: collapse;">
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666; width: 40%;">Donor Name:</td>
                <td style="padding: 10px 0; color: #333; font-weight: 600;">${donorData.fullName}</td>
              </tr>
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666;">Email:</td>
                <td style="padding: 10px 0; color: #333;">
                  <a href="mailto:${donorData.email}" style="color: #667eea; text-decoration: none;">
                    ${donorData.email}
                  </a>
                </td>
              </tr>
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666;">Amount:</td>
                <td style="padding: 10px 0; color: #28a745; font-weight: 700; font-size: 16px;">
                  ${donorData.donationAmount}
                </td>
              </tr>
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666;">Social Media:</td>
                <td style="padding: 10px 0; color: #333;">${donorData.socialMedia}</td>
              </tr>
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666;">Certificate #:</td>
                <td style="padding: 10px 0; color: #333; font-family: monospace;">
                  ${donorData.certificateNumber}
                </td>
              </tr>
              <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px 0; color: #666;">Date & Time:</td>
                <td style="padding: 10px 0; color: #333;">${donorData.timestamp}</td>
              </tr>
              <tr>
                <td style="padding: 10px 0; color: #666;">Receipt:</td>
                <td style="padding: 10px 0;">
                  ${donorData.receiptUrl && donorData.receiptUrl.startsWith('http') ? 
                    `<a href="${donorData.receiptUrl}" target="_blank" style="color: #667eea; text-decoration: none;">
                      View Receipt →
                    </a>` : 
                    `<span style="color: #666; font-style: italic;">${donorData.receiptUrl || 'No receipt provided'}</span>`
                  }
                </td>
              </tr>
            </table>
          </div>
          
          <!-- Message from Donor -->
          ${donorData.message && donorData.message !== 'Thank you for your support!' ? `
          <div style="background: #fff3e0; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #ffc107;">
            <p style="color: #856404; font-weight: 600; margin: 0 0 8px 0; font-size: 14px;">
              Message from Donor:
            </p>
            <p style="color: #333; margin: 0; font-style: italic; font-size: 14px;">
              "${donorData.message}"
            </p>
          </div>
          ` : ''}
          
          <!-- Status -->
          <div style="background: #d4edda; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #28a745;">
            <p style="color: #155724; margin: 0; font-size: 14px;">
              ✅ <strong>Certificate Status:</strong> Successfully generated and sent to donor
            </p>
          </div>
          
          <!-- Quick Actions -->
          <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0;">
            <p style="color: #666; font-size: 13px; margin: 0;">
              <strong>Quick Actions:</strong><br>
              • Certificate has been automatically sent to the donor<br>
              • PDF copy saved in Google Drive<br>
              • You can reply directly to ${donorData.email} if needed
            </p>
          </div>
        </div>
        
        <!-- Footer -->
        <div style="background: #f8f9fa; padding: 15px; text-align: center; border-top: 1px solid #e0e0e0;">
          <p style="color: #999; font-size: 12px; margin: 0;">
            This is an automated notification from AESPA Donation System
          </p>
        </div>
      </div>
    </body>
    </html>
  `;
  
  // Create plain text version
  const adminPlainText = `
New Donation Received!

DONATION DETAILS:
-----------------
Donor Name: ${donorData.fullName}
Email: ${donorData.email}
Amount: ${donorData.donationAmount}
Social Media: ${donorData.socialMedia}
Certificate #: ${donorData.certificateNumber}
Date & Time: ${donorData.timestamp}
Receipt: ${donorData.receiptUrl}

${donorData.message && donorData.message !== 'Thank you for your support!' ? 
`Message from Donor:
"${donorData.message}"

` : ''}
STATUS: Certificate successfully generated and sent to donor.

---
This is an automated notification from AESPA Donation System
  `;
  
  // Send admin notification email (no attachments)
  GmailApp.sendEmail(
    CONFIG.ADMIN_EMAIL,
    CONFIG.ADMIN_EMAIL_SUBJECT + ' - ' + donorData.fullName,
    adminPlainText,
    {
      htmlBody: adminHtmlBody,
      name: CONFIG.SENDER_NAME,
      noReply: true
    }
  );
  
  console.log('Admin notification sent to: ' + CONFIG.ADMIN_EMAIL);
}

// Send certificate via email
function sendCertificateEmail(donorData, certificateFile) {
  // Get the PDF blob
  const pdfBlob = certificateFile.getBlob();
  
  // Create HTML email body
  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap');
      </style>
    </head>
    <body style="margin: 0; padding: 0; font-family: 'Poppins', Arial, sans-serif; background-color: #f4f4f4;">
      <div style="max-width: 600px; margin: 20px auto; background-color: white; border-radius: 15px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        
        <!-- Header with gradient -->
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 40px 30px; text-align: center;">
          <h1 style="color: white; margin: 0; font-size: 28px; font-weight: 700; text-shadow: 2px 2px 4px rgba(0,0,0,0.1);">
            Thank You for Supporting AESPA! 
          </h1>
          <p style="color: rgba(255,255,255,0.9); margin: 10px 0 0 0; font-size: 16px;">
            Your generosity makes a difference
          </p>
        </div>
        
        <!-- Main content -->
        <div style="padding: 40px 30px;">
          <p style="font-size: 18px; color: #333; margin: 0 0 20px 0;">
            Dear <strong>${donorData.fullName}</strong>,
          </p>
          
          <p style="font-size: 15px; color: #555; line-height: 1.8; margin: 0 0 25px 0;">
            We are incredibly grateful for your generous donation. Your support means everything 
            to us and helps us continue our mission. We're honored to have you as part of our community!
          </p>
          
          <!-- Certificate details card -->
          <div style="background: linear-gradient(135deg, #f5f7ff 0%, #f0e6ff 100%); padding: 25px; border-radius: 12px; margin: 25px 0; border-left: 4px solid #667eea;">
            <h3 style="color: #667eea; margin: 0 0 15px 0; font-size: 18px; font-weight: 600;">
              📜 Your Certificate Details
            </h3>
            <table style="width: 100%; font-size: 14px;">
              <tr>
                <td style="padding: 8px 0; color: #666;">Certificate Number:</td>
                <td style="padding: 8px 0; color: #333; font-weight: 600;">${donorData.certificateNumber}</td>
              </tr>
              <tr>
                <td style="padding: 8px 0; color: #666;">Date Issued:</td>
                <td style="padding: 8px 0; color: #333; font-weight: 600;">${donorData.donationDate}</td>
              </tr>
              <tr>
                <td style="padding: 8px 0; color: #666;">Donation Amount:</td>
                <td style="padding: 8px 0; color: #333; font-weight: 600;">${donorData.donationAmount}</td>
              </tr>
            </table>
          </div>
          
          ${donorData.message && donorData.message !== 'Thank you for your support!' ? `
          <!-- Message display -->
          <div style="background: #fff3e0; padding: 20px; border-radius: 10px; margin: 25px 0; border-left: 4px solid #ff9800;">
            <p style="color: #e65100; font-weight: 600; margin: 0 0 8px 0; font-size: 14px;">
              💬 Your message for AESPA:
            </p>
            <p style="color: #555; margin: 0; font-style: italic; font-size: 14px; line-height: 1.6;">
              "${donorData.message}"
            </p>
          </div>
          ` : ''}
          
          <!-- Call to action -->
          <div style="text-align: center; margin: 35px 0;">
            <p style="color: #666; font-size: 14px; margin: 0 0 15px 0;">
              Your official donation certificate is attached to this email.
            </p>
            <div style="display: inline-block; background: #667eea; color: white; padding: 14px 35px; border-radius: 50px; font-weight: 600; font-size: 15px; text-decoration: none; box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);">
              📎 See Attached Certificate (PDF)
            </div>
          </div>
          
          <!-- Social media -->
          <div style="text-align: center; padding: 25px 0; border-top: 1px solid #eee; margin-top: 35px;">
            <p style="color: #999; font-size: 13px; margin: 0 0 10px 0;">
              Share your support with the community!
            </p>
            <p style="color: #667eea; font-size: 14px; margin: 0;">
              Tag us: ${donorData.socialMedia}
            </p>
          </div>
        </div>
        
        <!-- Footer -->
        <div style="background: #f8f9fa; padding: 20px; text-align: center;">
          <p style="color: #999; font-size: 12px; margin: 0;">
            This certificate confirms your donation and can be kept for your records.
          </p>
          <p style="color: #999; font-size: 12px; margin: 5px 0 0 0;">
            Thank you for being an amazing supporter! 💜
          </p>
        </div>
      </div>
    </body>
    </html>
  `;
  
  // Send email with attachment
  GmailApp.sendEmail(
    donorData.email,
    CONFIG.EMAIL_SUBJECT,
    `Thank you for your donation, ${donorData.fullName}! Your certificate is attached.\n\nCertificate Number: ${donorData.certificateNumber}\nAmount: ${donorData.donationAmount}\n\nThank you for supporting AESPA!`,
    {
      htmlBody: htmlBody,
      attachments: [pdfBlob],
      name: CONFIG.SENDER_NAME
    }
  );
}

// Notify admin of errors
function notifyAdminOfError(error, formEvent) {
  const adminEmail = Session.getActiveUser().getEmail();
  
  let eventDetails = 'No form data available';
  try {
    if (formEvent && formEvent.response) {
      eventDetails = 'Form submission from: ' + formEvent.response.getRespondentEmail();
    }
  } catch (e) {
    // Ignore errors in getting event details
  }
  
  GmailApp.sendEmail(
    adminEmail,
    '⚠️ Error in AESPA Donation Certificate Generator',
    `An error occurred while processing a donation certificate.\n\n` +
    `Error: ${error.toString()}\n\n` +
    `Details: ${eventDetails}\n\n` +
    `Stack trace:\n${error.stack}\n\n` +
    `Please check the script logs for more information.`
  );
}

// ============================================
// HELPER FUNCTION - Get Form URLs
// ============================================
function getFormURLs() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const editUrl = scriptProperties.getProperty('FORM_EDIT_URL');
  const publicUrl = scriptProperties.getProperty('FORM_PUBLIC_URL');
  const formId = scriptProperties.getProperty('FORM_ID');
  
  if (!editUrl || !publicUrl) {
    console.log('No saved form URLs found. Please run Step3_CreateAndSetupForm first.');
    return null;
  }
  
  console.log('=====================================');
  console.log('📝 FORM EDIT URL:');
  console.log(editUrl);
  console.log('=====================================');
  console.log('🔗 FORM PUBLIC URL (Share with donors):');
  console.log(publicUrl);
  console.log('=====================================');
  console.log('📄 FORM ID: ' + formId);
  console.log('=====================================');
  
  return {
    editUrl: editUrl,
    publicUrl: publicUrl,
    formId: formId
  };
}

// ============================================
// ALTERNATIVE: Find Your Form Manually
// ============================================
function findMyForm() {
  try {
    const forms = DriveApp.getFilesByType(MimeType.GOOGLE_FORMS);
    let foundForm = false;
    
    console.log('🔍 SEARCHING FOR FORMS IN YOUR GOOGLE DRIVE...');
    console.log('=====================================');
    
    while (forms.hasNext()) {
      const form = forms.next();
      const formName = form.getName();
      
      // Look for our donation form
      if (formName.includes('AESPA') || formName.includes('Donation')) {
        foundForm = true;
        const formId = form.getId();
        const formUrl = form.getUrl();
        
        console.log('📋 FOUND: ' + formName);
        console.log('🔗 URL: ' + formUrl);
        console.log('📄 ID: ' + formId);
        console.log('-------------------------------------');
        
        // Try to get the actual form object for edit URL
        try {
          const formApp = FormApp.openById(formId);
          const editUrl = formApp.getEditUrl();
          const publishedUrl = formApp.getPublishedUrl();
          
          console.log('📝 EDIT URL: ' + editUrl);
          console.log('🔗 PUBLIC URL: ' + publishedUrl);
          console.log('=====================================');
        } catch (e) {
          console.log('(Could not get additional URLs for this form)');
        }
      }
    }
    
    if (!foundForm) {
      console.log('❌ No donation forms found in your Drive.');
      console.log('Please run Step3_CreateAndSetupForm to create one.');
    }
    
  } catch (error) {
    console.log('Error searching for forms: ' + error.toString());
  }
}

// ============================================
// HELPER FUNCTION - Process Receipt URL
// ============================================
function processReceiptUrl(rawUrl) {
  // If no URL provided
  if (!rawUrl) {
    return 'No receipt provided';
  }
  
  // Convert to string and trim
  const urlString = String(rawUrl).trim();
  
  // Check if it's a Google Drive file ID (from file upload in forms)
  // File IDs are typically 33 characters long and contain letters, numbers, -, _
  if (urlString.match(/^[a-zA-Z0-9_-]{20,50}$/) && !urlString.includes('http')) {
    // It's a file ID, convert to Google Drive URL
    return `https://drive.google.com/file/d/${urlString}/view`;
  }
  
  // Check if it's already a valid URL
  if (urlString.startsWith('http://') || urlString.startsWith('https://')) {
    return urlString;
  }
  
  // Check if it's a drive.google.com URL without https
  if (urlString.startsWith('drive.google.com')) {
    return 'https://' + urlString;
  }
  
  // Check if it's an array (happens with file uploads)
  if (Array.isArray(rawUrl) && rawUrl.length > 0) {
    // Get the first item and recursively process it
    return processReceiptUrl(rawUrl[0]);
  }
  
  // For any other format, try to make it a URL
  if (urlString.includes('.') && !urlString.includes(' ')) {
    // Looks like a domain, add https://
    return 'https://' + urlString;
  }
  
  // If we can't process it, return a message
  return 'Receipt link format not recognized: ' + urlString.substring(0, 50);
}

// ============================================
// ALTERNATIVE: Create Form with File Upload Option
// ============================================
function Step3b_CreateFormWithFileUpload() {
  try {
    // Check if user has Google Workspace
    const userEmail = Session.getActiveUser().getEmail();
    console.log('Creating form for user: ' + userEmail);
    
    // Create a new form
    const form = FormApp.create('AESPA Donation Form - File Upload Version');
    
    // Customize form settings
    form.setDescription('Thank you for supporting AESPA! Please fill out this form to receive your donation certificate.');
    form.setConfirmationMessage('Thank you for your donation! You will receive your certificate via email shortly.');
    form.setCollectEmail(true);
    
    // Add form fields
    form.addTextItem()
      .setTitle('Full Name')
      .setHelpText('Please enter your full name as you want it to appear on the certificate')
      .setRequired(true);
    
    form.addTextItem()
      .setTitle('Social Media Handle')
      .setHelpText('Your Instagram, Twitter, or other social media username')
      .setRequired(true);
    
    form.addTextItem()
      .setTitle('Donation Amount')
      .setHelpText('Please enter the amount you donated (e.g., $50 or 50 USD)')
      .setRequired(true);
    
    form.addParagraphTextItem()
      .setTitle('Message for AESPA (Optional)')
      .setHelpText('Leave a message of support for AESPA (this will appear on your certificate)')
      .setRequired(false);
    
    // Multiple options for receipt
    form.addMultipleChoiceItem()
      .setTitle('How would you like to provide your receipt?')
      .setChoices([
        form.createChoice('Upload an image file'),
        form.createChoice('Provide a link to the receipt')
      ])
      .setRequired(true);
    
    // Option 1: URL input (works for everyone)
    form.addTextItem()
      .setTitle('Receipt Link (URL)')
      .setHelpText('Please paste the link to your receipt (Google Drive, Dropbox, Imgur, etc.)')
      .setRequired(false);
    
    // Option 2: For text description if they can't upload
    form.addParagraphTextItem()
      .setTitle('Receipt Description (Alternative)')
      .setHelpText('If you cannot provide a link, please describe your payment details (transaction ID, date, payment method)')
      .setRequired(false);
    
    // Set up trigger
    ScriptApp.newTrigger('onFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();
    
    // Get URLs
    const editUrl = form.getEditUrl();
    const publishedUrl = form.getPublishedUrl();
    
    console.log('✅ Form with alternative receipt options created!');
    console.log('📝 Form Edit URL: ' + editUrl);
    console.log('🔗 Form Public URL: ' + publishedUrl);
    
    return {
      editUrl: editUrl,
      publishedUrl: publishedUrl
    };
    
  } catch (error) {
    console.log('❌ Error creating form: ' + error.toString());
    throw error;
  }
}

// ============================================
// DEBUG FUNCTION - Check Admin Email Settings
// ============================================
function debugAdminEmail() {
  console.log('🔍 DEBUGGING ADMIN EMAIL SETTINGS');
  console.log('=====================================');
  console.log('ADMIN_EMAIL: ' + CONFIG.ADMIN_EMAIL);
  console.log('ENABLE_ADMIN_NOTIFICATIONS: ' + CONFIG.ENABLE_ADMIN_NOTIFICATIONS);
  console.log('Current User Email: ' + Session.getActiveUser().getEmail());
  console.log('=====================================');
  
  if (CONFIG.ENABLE_ADMIN_NOTIFICATIONS) {
    console.log('✅ Admin notifications are ENABLED');
    
    // Test sending a notification
    const testData = {
      email: 'test@example.com',
      fullName: 'Debug Test',
      socialMedia: '@debug_test',
      donationAmount: '$TEST',
      message: 'This is a debug test',
      receiptUrl: 'https://example.com/debug',
      donationDate: new Date().toLocaleDateString(),
      certificateNumber: 'DEBUG-001',
      timestamp: new Date().toLocaleString()
    };
    
    try {
      console.log('Attempting to send test admin email...');
      sendAdminNotification(testData);
      console.log('✅ TEST EMAIL SENT SUCCESSFULLY!');
      console.log('Check inbox for: ' + CONFIG.ADMIN_EMAIL);
    } catch (error) {
      console.log('❌ ERROR SENDING TEST EMAIL: ' + error.toString());
    }
  } else {
    console.log('❌ Admin notifications are DISABLED');
    console.log('Set ENABLE_ADMIN_NOTIFICATIONS to true to enable');
  }
}
function Step4_TestCertificateGeneration() {
  // Test data
  const testData = {
    email: Session.getActiveUser().getEmail(),  // Will send to your email
    fullName: 'Test Donor',
    socialMedia: '@testdonor',
    donationAmount: '$100 USD',
    message: 'This is a test message for AESPA!',
    receiptUrl: 'https://example.com/receipt',
    donationDate: new Date().toLocaleDateString('en-US', { 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric' 
    }),
    certificateNumber: 'TEST-' + new Date().getTime()
  };
  
  try {
    Logger.log('🧪 Starting test certificate generation...');
    
    // Generate certificate
    const certificateFile = createCertificate(testData);
    Logger.log('✅ Certificate created: ' + certificateFile.getName());
    
    // Send email
    sendCertificateEmail(testData, certificateFile);
    Logger.log('✅ Email sent to: ' + testData.email);
    
    Logger.log('🎉 Test completed successfully! Check your email.');
    
  } catch (error) {
    Logger.log('❌ Test failed: ' + error.toString());
    throw error;
  }
}
