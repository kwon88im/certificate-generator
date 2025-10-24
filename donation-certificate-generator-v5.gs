// Google Apps Script for Donation Certificate Generator
// Follow the setup steps in order!

// ============================================
// STEP 1: RUN THIS FIRST - Create Certificate Template
// ============================================
function Step1_CreateCertificateTemplate() {
  try {
    const doc = DocumentApp.create('Donation Certificate Template');
    const body = doc.getBody();

    // Clear default content
    body.clear();

    // Page margins (1 inch)
    body.setMarginTop(72);
    body.setMarginBottom(72);
    body.setMarginLeft(72);
    body.setMarginRight(72);

    // Decorative spacing
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

    // Donor name
    const donorName = body.appendParagraph('{{FULL_NAME}}');
    donorName.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    donorName.setFontSize(22);
    donorName.setBold(true);
    donorName.setForegroundColor('#333333');
    donorName.setSpacingAfter(15);

    // Social media
    const social = body.appendParagraph('( {{SOCIAL_MEDIA}} )');
    social.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    social.setFontSize(11);
    social.setItalic(true);
    social.setForegroundColor('#666666');
    social.setSpacingAfter(25);

    // Description
    const description = body.appendParagraph('has made a generous donation in support of');
    description.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    description.setFontSize(13);
    description.setSpacingAfter(10);

    // Project name
    const projectName = body.appendParagraph('AESPA AEXIS LINE Fan Project');
    projectName.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    projectName.setFontSize(18);
    projectName.setBold(true);
    projectName.setForegroundColor('#667eea');
    projectName.setSpacingAfter(30);

    // Date
    const date = body.appendParagraph('Given on {{DATE}}');
    date.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    date.setFontSize(12);
    date.setSpacingAfter(40);

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

    doc.saveAndClose();

    const templateId = doc.getId();
    const templateUrl = doc.getUrl();

    Logger.log('✅ Certificate template created successfully!');
    Logger.log('📋 Template ID: ' + templateId);
    Logger.log('🔗 Template URL: ' + templateUrl);
    Logger.log('');
    Logger.log('⚠️ IMPORTANT: Copy the Template ID above and paste it in Step 2!');

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
  // Your template id from Step 1
  TEMPLATE_DOC_ID: '13WLhFYDcIPLILmgOn6uzGNgGV0TfAPKO2s9b3rqQogQ',

  // OPTIONAL: folder to store certificates
  CERTIFICATES_FOLDER_ID: '1vl2wiODYynu_AGa_N2tmy1QhMeH9JPO5',

  // Email configuration
  EMAIL_SUBJECT: 'Thank You for Your Donation to aespa - Certificate Attached',
  SENDER_NAME: 'aespa INA UNION',

  // Certificate numbering prefix
  CERTIFICATE_PREFIX: 'aespa-2025-',

  // Admin notification settings
  ADMIN_EMAIL: 'aiufankit@gmail.com',
  ENABLE_ADMIN_NOTIFICATIONS: true,
  ADMIN_EMAIL_SUBJECT: 'New Donation Received!',

// Just enable/disable announcement
  ENABLE_ANNOUNCEMENT: true,

  // Announcement shown/used wherever you need
  ANNOUNCEMENT_TEXT: `Giveaway untuk Donatur dengan ketentuan berikut:
- Donasi min. IDR 50,000 : Card holder, MD official hair clip, PC aespa official
- Donasi min. IDR 75,000 : Voucher alfamart 50k, PC aespa official
- Donasi min. IDR 100,000 : T-shirt jikjik + sticker winter 2025, Magazine Cover Karina
- Donasi min. IDR 125,000 : Giselle slogan cheering kit, Magazine Cover Giselle
- Donasi min. IDR 150,000 : PC aespa official set, Magazine cover member aespa
Ikuti sosial media @aespainaunion di X (Twitter) dan Instagram untuk informasi jadwal pengundian.`
};

// ============================================
// STEP 3: RUN THIS - Create and Setup Form
// ============================================
function Step3_CreateAndSetupForm() {
  try {
    // Only check that an ID exists
    if (!CONFIG.TEMPLATE_DOC_ID) {
      throw new Error('Please set CONFIG.TEMPLATE_DOC_ID first (run Step 1 to get the ID).');
    }

    const form = FormApp.create('Fan Donation Form');

    form.setDescription('Thank you for supporting aespa! Please fill out this form to receive your donation certificate.');
    form.setConfirmationMessage('Thank you for your donation! You will receive your certificate via email shortly.');
    form.setCollectEmail(true);

    form.addTextItem()
      .setTitle('Full Name')
      .setHelpText('Masukkan nama lengkap Anda sesuai yang ingin ditampilkan pada sertifikat.')
      .setRequired(true);

    form.addTextItem()
      .setTitle('Social Media')
      .setHelpText('Instagram / X (Twitter)')
      .setRequired(true);

    form.addTextItem()
      .setTitle('Donation Amount')
      .setHelpText('Silahkan masukan jumlah donasi mu (contoh: 50.000 / Rp 50.000)')
      .setRequired(true);

    // Use URL link (compatible for all users)
    form.addTextItem()
      .setTitle('Upload Bukti Transfer')
      .setHelpText('Please paste a link to your payment receipt (Google Drive, Dropbox, Imgur, etc.)')
      .setRequired(true)
      .setValidation(
        FormApp.createTextValidation()
          .requireTextIsUrl()
          .setHelpText('Please enter a valid URL')
          .build()
      );

    const formId = form.getId();

    ScriptApp.newTrigger('onFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();

    const editUrl = form.getEditUrl();
    const publishedUrl = form.getPublishedUrl();

    console.log('✅ FORM CREATED SUCCESSFULLY!');
    console.log('📝 EDIT URL:', editUrl);
    console.log('🔗 PUBLIC URL:', publishedUrl);
    console.log('📄 FORM ID:', formId);

    Logger.log('✅ Form created and trigger set up successfully!');
    Logger.log('📝 Form Edit URL: ' + editUrl);
    Logger.log('🔗 Form Public URL: ' + publishedUrl);
    Logger.log('Form ID: ' + formId);

    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('FORM_EDIT_URL', editUrl);
    scriptProperties.setProperty('FORM_PUBLIC_URL', publishedUrl);
    scriptProperties.setProperty('FORM_ID', formId);

    const summary = `
FORM CREATED SUCCESSFULLY!
==========================
Form Name: Fan Donation Form
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
    console.log('❌ ERROR:', error.toString());
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

    if (!CONFIG.TEMPLATE_DOC_ID) {
      throw new Error('Template ID not configured. Please run setup steps first.');
    }

    const responses = e.response;
    const itemResponses = responses.getItemResponses();

// Extract donor information from form responses
const donorData = {
  email: responses.getRespondentEmail(),
  fullName: '',
  socialMedia: '',
  whatsappNumber: '',           // NEW FIELD
  paymentMethod: '',            // NEW FIELD (BCA/QRIS)
  paymentDate: '',              // NEW FIELD
  accountOwnerName: '',         // NEW FIELD
  donationAmount: '',
  message: 'Thank you for your support!',
  receiptUrl: '',
donationDate: new Date().toLocaleDateString('id-ID', {
  year: 'numeric',
  month: '2-digit',
  day: '2-digit'
}).replace(/\./g, '/'),
  certificateNumber: generateCertificateNumber(),
  timestamp: new Date().toLocaleString('en-US', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  })
};

// Match Indonesian field names
for (let i = 0; i < itemResponses.length; i++) {
  const response = itemResponses[i];
  const title = response.getItem().getTitle();
  const answer = response.getResponse();

  // Update to Indonesian field names
  if (title.includes('Nama Lengkap')) {
    donorData.fullName = answer;
  } else if (title.includes('Sosial Media')) {
    donorData.socialMedia = answer;
  } else if (title.includes('Nomor Whatsapp')) {
    donorData.whatsappNumber = answer;
  } else if (title.includes('Pembayaran melalui')) {
    donorData.paymentMethod = answer;
  } else if (title.includes('Tanggal Pembayaran')) {
    donorData.paymentDate = answer;
  } else if (title.includes('Nama Pemilik Rekening')) {
    donorData.accountOwnerName = answer;
  } else if (title.includes('Nominal Donasi')) {
    donorData.donationAmount = answer;
  } else if (title.includes('Upload Bukti Transfer')) {
    donorData.receiptUrl = processReceiptUrl(answer);
  }
}

    Logger.log('👤 Processing certificate for: ' + donorData.fullName);

    const certificateFile = createCertificate(donorData);
    sendCertificateEmail(donorData, certificateFile);
    Logger.log('✅ Certificate sent to donor: ' + donorData.email);

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
    try {
      notifyAdminOfError(error, e);
    } catch (notifyError) {
      Logger.log('Could not send admin notification: ' + notifyError.toString());
    }
  }
}

// ============================================
// ADMIN NOTIFICATION (single version)
// ============================================
function sendAdminNotification(donorData) {
  const adminHtmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');</style>
    </head>
    <body style="margin:0;padding:0;font-family:'Inter',Arial,sans-serif;background-color:#f5f5f5;">
      <div style="max-width:600px;margin:20px auto;background-color:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);">
        <div style="background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);padding:30px;text-align:center;">
          <h2 style="color:#fff;margin:0;font-size:24px;">New Donation Received!</h2>
        </div>
        <div style="padding:30px;">
          <p style="color:#333;font-size:16px;margin:0 0 20px 0;">Great news! A new donation has been received and processed.</p>
          <div style="background:#f8f9fa;padding:20px;border-radius:8px;margin:20px 0;">
            <h3 style="color:#333;margin:0 0 15px 0;font-size:18px;">Donation Details</h3>
            <table style="width:100%;font-size:14px;border-collapse:collapse;">
              <tr style="border-bottom:1px solid #e0e0e0;"><td style="padding:10px 0;color:#666;width:40%;">Donor Name:</td><td style="padding:10px 0;color:#333;font-weight:600;">${donorData.fullName}</td></tr>
              <tr style="border-bottom:1px solid #e0e0e0;"><td style="padding:10px 0;color:#666;">Email:</td><td style="padding:10px 0;color:#333;"><a href="mailto:${donorData.email}" style="color:#667eea;text-decoration:none;">${donorData.email}</a></td></tr>
              <tr style="border-bottom:1px solid #e0e0e0;"><td style="padding:10px 0;color:#666;">Amount:</td><td style="padding:10px 0;color:#28a745;font-weight:700;font-size:16px;">${donorData.donationAmount}</td></tr>
              <tr style="border-bottom:1px solid #e0e0e0;"><td style="padding:10px 0;color:#666;">Social Media:</td><td style="padding:10px 0;color:#333;">${donorData.socialMedia}</td></tr>
              <tr style="border-bottom:1px solid #e0e0e0;"><td style="padding:10px 0;color:#666;">Certificate #:</td><td style="padding:10px 0;color:#333;font-family:monospace;">${donorData.certificateNumber}</td></tr>
              <tr style="border-bottom:1px solid #e0e0e0;"><td style="padding:10px 0;color:#666;">Date & Time:</td><td style="padding:10px 0;color:#333;">${donorData.timestamp}</td></tr>
              <tr style="border-bottom:1px solid #e0e0e0;"><td style="padding:10px 0;color:#666;">WhatsApp:</td><td style="padding:10px 0;color:#333;">${donorData.whatsappNumber}</td></tr>
              <tr style="border-bottom:1px solid #e0e0e0;"><td style="padding:10px 0;color:#666;">Payment Method:</td><td style="padding:10px 0;color:#333;">${donorData.paymentMethod}</td></tr>
              <tr style="border-bottom:1px solid #e0e0e0;"><td style="padding:10px 0;color:#666;">Account Owner:</td><td style="padding:10px 0;color:#333;">${donorData.accountOwnerName}</td></tr>
              <tr style="border-bottom:1px solid #e0e0e0;"><td style="padding:10px 0;color:#666;">Payment Date:</td><td style="padding:10px 0;color:#333;">${donorData.paymentDate}</td></tr>
              <tr><td style="padding:10px 0;color:#666;">Receipt:</td>
                <td style="padding:10px 0;">
                  ${donorData.receiptUrl && donorData.receiptUrl.startsWith('http')
                    ? `<a href="${donorData.receiptUrl}" target="_blank" style="color:#667eea;text-decoration:none;">View Receipt →</a>`
                    : `<span style="color:#666;font-style:italic;">${donorData.receiptUrl || 'No receipt provided'}</span>`}
                </td>
              </tr>
            </table>
          </div>
          ${donorData.message && donorData.message !== 'Thank you for your support!' ? `
          <div style="background:#fff3e0;padding:15px;border-radius:8px;margin:20px 0;border-left:4px solid #ffc107;">
            <p style="color:#856404;font-weight:600;margin:0 0 8px 0;font-size:14px;">💬 Message from Donor:</p>
            <p style="color:#333;margin:0;font-style:italic;font-size:14px;">"${donorData.message}"</p>
          </div>` : ''}
          <div style="background:#d4edda;padding:15px;border-radius:8px;margin:20px 0;border-left:4px solid #28a745;">
            <p style="color:#155724;margin:0;font-size:14px;">✅ <strong>Certificate Status:</strong> Successfully generated and sent to donor</p>
          </div>
          <div style="margin-top:30px;padding-top:20px;border-top:1px solid #e0e0e0;">
            <p style="color:#666;font-size:13px;margin:0;">
              <strong>Quick Actions:</strong><br>
              • Certificate sent to donor<br>
              • PDF copy saved in Google Drive<br>
              • You can reply directly to ${donorData.email} if needed
            </p>
          </div>
        </div>
        <div style="background:#f8f9fa;padding:15px;text-align:center;border-top:1px solid #e0e0e0;">
          <p style="color:#999;font-size:12px;margin:0;">This is an automated notification from aespa INA UNION Donation System</p>
        </div>
      </div>
    </body>
    </html>
  `;

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

` : ''}STATUS: Certificate successfully generated and sent to donor.

---
This is an automated notification from AESPA Donation System
  `;

  GmailApp.sendEmail(
    CONFIG.ADMIN_EMAIL,
    CONFIG.ADMIN_EMAIL_SUBJECT + ' - ' + donorData.fullName,
    adminPlainText,
    { htmlBody: adminHtmlBody, name: CONFIG.SENDER_NAME, noReply: true }
  );

  console.log('Admin notification sent to: ' + CONFIG.ADMIN_EMAIL);
}

// ============================================
// CERTIFICATE NUMBER
// ============================================
function generateCertificateNumber() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let currentNumber = scriptProperties.getProperty('CERT_NUMBER') || '1000';
  currentNumber = parseInt(currentNumber, 10) + 1;
  scriptProperties.setProperty('CERT_NUMBER', currentNumber.toString());
  return CONFIG.CERTIFICATE_PREFIX + currentNumber;
}

// ============================================
// CREATE CERTIFICATE (from Doc template to PDF)
// ============================================
function createCertificate(donorData) {
  const templateDoc = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);

  const certificateName =
    `Certificate_${donorData.certificateNumber}_${donorData.fullName.replace(/[^a-zA-Z0-9]/g, '_')}`;
  const certificateCopy = templateDoc.makeCopy(certificateName);

  if (CONFIG.CERTIFICATES_FOLDER_ID && CONFIG.CERTIFICATES_FOLDER_ID !== '') {
    try {
      const folder = DriveApp.getFolderById(CONFIG.CERTIFICATES_FOLDER_ID);
      certificateCopy.moveTo(folder);
    } catch (error) {
      Logger.log('Warning: Could not move to folder: ' + error.toString());
    }
  }

  const doc = DocumentApp.openById(certificateCopy.getId());
  const body = doc.getBody();

  body.replaceText('{{FULL_NAME}}', donorData.fullName);
  body.replaceText('{{CERTIFICATE_NUMBER}}', donorData.certificateNumber);
  body.replaceText('{{DATE}}', donorData.donationDate);
  body.replaceText('{{SOCIAL_MEDIA}}', donorData.socialMedia);

  // Clean up any old placeholders if present
  body.replaceText('{{AMOUNT}}', '');
  body.replaceText('{{MESSAGE}}', '');

  doc.saveAndClose();

  const pdfBlob = certificateCopy.getAs('application/pdf');
  pdfBlob.setName(certificateName + '.pdf');

  const pdfFile = DriveApp.createFile(pdfBlob);

  if (CONFIG.CERTIFICATES_FOLDER_ID && CONFIG.CERTIFICATES_FOLDER_ID !== '') {
    try {
      const folder = DriveApp.getFolderById(CONFIG.CERTIFICATES_FOLDER_ID);
      pdfFile.moveTo(folder);
    } catch (error) {
      Logger.log('Warning: Could not move PDF to folder: ' + error.toString());
    }
  }

  // Trash the intermediate Google Doc copy (keep only PDF)
  certificateCopy.setTrashed(true);

  return pdfFile;
}

// ============================================
// SEND CERTIFICATE EMAIL (with optional ANNOUNCEMENT_TEXT)
// ============================================
function sendCertificateEmail(donorData, certificateFile) {
  const pdfBlob = certificateFile.getBlob();

  // Safely materialize the announcement text with donor data
  const announcementText = (CONFIG.ANNOUNCEMENT_TEXT || '')
    .replace(/{{FULL_NAME}}/g, donorData.fullName)
    .replace(/{{DONATION_AMOUNT}}/g, donorData.donationAmount)
    .replace(/{{CERT_NUMBER}}/g, donorData.certificateNumber)
    .replace(/{{DATE}}/g, donorData.donationDate);

  const announcementBlock = CONFIG.ENABLE_ANNOUNCEMENT
  ? `
  <div style="background:#fffbe6;border-left:4px solid #f5c518;padding:18px;border-radius:10px;margin:25px 0;">
    <h3 style="margin:0 0 15px 0;font-size:18px;color:#7a5a00;text-align:center;">
      Giveaway untuk Donatur
    </h3>
    
    <!-- Responsive Table Container -->
    <div style="overflow-x:auto;">
      <table style="width:100%;min-width:400px;border-collapse:collapse;font-size:14px;background:#fff;border-radius:8px;overflow:hidden;">
        <thead>
          <tr style="background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:#fff;">
            <th style="padding:12px 8px;text-align:left;font-weight:600;border-bottom:2px solid #5568d3;min-width:100px;">Donasi Min.</th>
            <th style="padding:12px 8px;text-align:left;font-weight:600;border-bottom:2px solid #5568d3;">Hadiah</th>
          </tr>
        </thead>
        <tbody>
          <tr style="background:#f8f9ff;">
            <td style="padding:12px 8px;border-bottom:1px solid #e0e0e0;font-weight:600;color:#667eea;white-space:nowrap;">Rp 50.000</td>
            <td style="padding:12px 8px;border-bottom:1px solid #e0e0e0;color:#333;line-height:1.5;">Card holder, MD official hair clip, PC aespa official</td>
          </tr>
          <tr style="background:#fff;">
            <td style="padding:12px 8px;border-bottom:1px solid #e0e0e0;font-weight:600;color:#667eea;white-space:nowrap;">Rp 75.000</td>
            <td style="padding:12px 8px;border-bottom:1px solid #e0e0e0;color:#333;line-height:1.5;">Voucher alfamart 50k, PC aespa official</td>
          </tr>
          <tr style="background:#f8f9ff;">
            <td style="padding:12px 8px;border-bottom:1px solid #e0e0e0;font-weight:600;color:#667eea;white-space:nowrap;">Rp 100.000</td>
            <td style="padding:12px 8px;border-bottom:1px solid #e0e0e0;color:#333;line-height:1.5;">T-shirt jikjik + sticker winter 2025, Magazine Cover Karina</td>
          </tr>
          <tr style="background:#fff;">
            <td style="padding:12px 8px;border-bottom:1px solid #e0e0e0;font-weight:600;color:#667eea;white-space:nowrap;">Rp 125.000</td>
            <td style="padding:12px 8px;border-bottom:1px solid #e0e0e0;color:#333;line-height:1.5;">Giselle slogan cheering kit, Magazine Cover Giselle</td>
          </tr>
          <tr style="background:#f8f9ff;">
            <td style="padding:12px 8px;font-weight:600;color:#667eea;white-space:nowrap;">Rp 150.000</td>
            <td style="padding:12px 8px;color:#333;line-height:1.5;">PC aespa official set, Magazine cover member aespa</td>
          </tr>
        </tbody>
      </table>
    </div>
    
    <!-- Footer Note -->
    <div style="margin-top:15px;padding:12px;background:#fff;border-radius:6px;text-align:center;">
      <p style="margin:0;font-size:13px;color:#666;line-height:1.6;">
        Ikuti sosial media <strong style="color:#667eea;">@aespainaunion</strong> di X (Twitter) dan Instagram untuk informasi jadwal pengundian.
      </p>
    </div>
  </div>
  `
  : '';

  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap');
      </style>
    </head>
    <body style="margin:0;padding:0;font-family:'Poppins',Arial,sans-serif;background-color:#f4f4f4;">
      <div style="max-width:600px;margin:20px auto;background-color:#fff;border-radius:15px;overflow:hidden;box-shadow:0 4px 6px rgba(0,0,0,0.1);">
        <div style="background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);padding:40px 30px;text-align:center;">
          <h1 style="color:#fff;margin:0;font-size:28px;font-weight:700;text-shadow:2px 2px 4px rgba(0,0,0,0.1);">
            Thank You for Supporting aespa aeXIS LINE in Jakarta - Fan Project!
          </h1>
          <p style="color:rgba(255,255,255,0.9);margin:10px 0 0 0;font-size:16px;">Your generosity makes a difference</p>
        </div>

        <div style="padding:40px 30px;">
          <p style="font-size:18px;color:#333;margin:0 0 20px 0;">Dear <strong>${donorData.fullName}</strong>,</p>

          <p style="font-size:15px;color:#555;line-height:1.8;margin:0 0 25px 0;">
            Thank you for your generous donation. Your support means everything to us and makes it possible to create beautiful memories with aespa through this fan project.
          </p>

          <div style="background:linear-gradient(135deg,#f5f7ff 0%,#f0e6ff 100%);padding:25px;border-radius:12px;margin:25px 0;border-left:4px solid #667eea;">
            <h3 style="color:#667eea;margin:0 0 15px 0;font-size:18px;font-weight:600;">Your Certificate Details</h3>
            <table style="width:100%;font-size:14px;">
              <tr>
                <td style="padding:8px 0;color:#666;">Certificate Number:</td>
                <td style="padding:8px 0;color:#333;font-weight:600;">${donorData.certificateNumber}</td>
              </tr>
              <tr>
                <td style="padding:8px 0;color:#666;">Date Issued:</td>
                <td style="padding:8px 0;color:#333;font-weight:600;">${donorData.donationDate}</td>
              </tr>
              <tr>
                <td style="padding:8px 0;color:#666;">Donation Amount:</td>
                <td style="padding:8px 0;color:#333;font-weight:600;">${donorData.donationAmount}</td>
              </tr>
            </table>
          </div>

         ${announcementBlock}
        </div>

          <div style="text-align:center;padding:25px 0;border-top:1px solid #eee;margin-top:35px;">
            <p style="color:#999;font-size:16px;margin:0 0 10px 0;">Share your certificate with the community!</p>
            <p style="color:#667eea;font-size:18px;margin:0;">Tag us on X (Twitter) / Instagram: @aespainaunion</p>
          </div>

        <div style="background:#f8f9fa;padding:20px;text-align:center;">
          <p style="color:#999;font-size:15px;margin:0;">This certificate confirms your donation and can be kept for your records.</p>
          <p style="color:#999;font-size:15px;margin:5px 0 0 0;">Thank you for being an amazing supporter, MYne!</p>
        </div>
      </div>
    </body>
    </html>
  `;

  GmailApp.sendEmail(
    donorData.email,
    CONFIG.EMAIL_SUBJECT,
    `Thank you for your donation, ${donorData.fullName}! Your certificate is attached.\n\nCertificate Number: ${donorData.certificateNumber}\nAmount: ${donorData.donationAmount}\n\nThank you for supporting aespa!`,
    { htmlBody: htmlBody, attachments: [pdfBlob], name: CONFIG.SENDER_NAME }
  );
}

// ============================================
// ERROR NOTIFY
// ============================================
function notifyAdminOfError(error, formEvent) {
  const adminEmail = Session.getActiveUser().getEmail();

  let eventDetails = 'No form data available';
  try {
    if (formEvent && formEvent.response) {
      eventDetails = 'Form submission from: ' + formEvent.response.getRespondentEmail();
    }
  } catch (e) { /* ignore */ }

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

  return { editUrl, publicUrl, formId };
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

      if (formName.includes('AESPA') || formName.includes('Donation')) {
        foundForm = true;
        const formId = form.getId();
        const formUrl = form.getUrl();

        console.log('📋 FOUND: ' + formName);
        console.log('🔗 URL: ' + formUrl);
        console.log('📄 ID: ' + formId);
        console.log('-------------------------------------');

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
// HELPER FUNCTION - Format Amount as Rupiah (NOT WORKING)
// ============================================
function formatRupiah(amount) {
  // Handle null/undefined
  if (!amount) return 'Rp 0';
  
  // Convert to string and remove all non-numeric characters except digits
  let cleanAmount = String(amount).replace(/[^\d]/g, '');
  
  // If nothing left, return 0
  if (!cleanAmount || cleanAmount === '') return 'Rp 0';
  
  // Convert to number
  let numericAmount = parseInt(cleanAmount, 10);
  
  // Format with thousand separators (dots for Indonesian style)
  const formatted = numericAmount.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  
  // Return with Rp prefix
  return 'Rp ' + formatted;
}

function testFormatRupiah() {
  console.log(formatRupiah('50000'));        // Rp 50.000
  console.log(formatRupiah('50.000'));       // Rp 50.000
  console.log(formatRupiah('Rp 50.000'));    // Rp 50.000
  console.log(formatRupiah('50,000'));       // Rp 50.000
  console.log(formatRupiah('1234567'));      // Rp 1.234.567
  console.log(formatRupiah('100'));          // Rp 100
  console.log(formatRupiah(null));           // Rp 0
}

// ============================================
// HELPER FUNCTION - Process Receipt URL
// ============================================
function processReceiptUrl(rawUrl) {
  if (!rawUrl) return 'No receipt provided';

  if (Array.isArray(rawUrl) && rawUrl.length > 0) {
    return processReceiptUrl(rawUrl[0]);
  }

  const urlString = String(rawUrl).trim();

  // Possibly a Drive file ID (no http, 20–50 chars)
  if (urlString.match(/^[a-zA-Z0-9_-]{20,50}$/) && !urlString.includes('http')) {
    return `https://drive.google.com/file/d/${urlString}/view`;
  }

  if (urlString.startsWith('http://') || urlString.startsWith('https://')) {
    return urlString;
  }

  if (urlString.startsWith('drive.google.com')) {
    return 'https://' + urlString;
  }

  if (urlString.includes('.') && !urlString.includes(' ')) {
    return 'https://' + urlString;
  }

  return 'Receipt link format not recognized: ' + urlString.substring(0, 50);
}

// ============================================
// ALTERNATIVE: Create Form with File Upload Option (link/alt text)
// ============================================
function Step3b_CreateFormWithFileUpload() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    console.log('Creating form for user: ' + userEmail);

    const form = FormApp.create('AESPA Donation Form - File Upload Version');
    form.setDescription('Thank you for supporting AESPA! Please fill out this form to receive your donation certificate.');
    form.setConfirmationMessage('Thank you for your donation! You will receive your certificate via email shortly. Share your certificate with the community and tag us @aespainaunion !');
    form.setCollectEmail(true);

    form.addTextItem().setTitle('Full Name').setHelpText('Please enter your full name as you want it to appear on the certificate').setRequired(true);
    form.addTextItem().setTitle('Social Media Handle').setHelpText('Your Instagram, Twitter, or other social media username').setRequired(true);
    form.addTextItem().setTitle('Donation Amount').setHelpText('Please enter the amount you donated (e.g., $50 or 50 USD)').setRequired(true);
    form.addParagraphTextItem().setTitle('Message for aespa (Optional)').setHelpText('Leave a message of support for aespa (this will appear on your certificate)').setRequired(false);

    form.addMultipleChoiceItem()
      .setTitle('How would you like to provide your receipt?')
      .setChoices([ form.createChoice('Upload an image file'), form.createChoice('Provide a link to the receipt') ])
      .setRequired(true);

    form.addTextItem()
      .setTitle('Receipt Link (URL)')
      .setHelpText('Please paste the link to your receipt (Google Drive, Dropbox, Imgur, etc.)')
      .setRequired(false);

    form.addParagraphTextItem()
      .setTitle('Receipt Description (Alternative)')
      .setHelpText('If you cannot provide a link, please describe your payment details (transaction ID, date, payment method)')
      .setRequired(false);

    ScriptApp.newTrigger('onFormSubmit').forForm(form).onFormSubmit().create();

    const editUrl = form.getEditUrl();
    const publishedUrl = form.getPublishedUrl();

    console.log('✅ Form with alternative receipt options created!');
    console.log('📝 Form Edit URL: ' + editUrl);
    console.log('🔗 Form Public URL: ' + publishedUrl);

    return { editUrl, publishedUrl };

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

// ============================================
// STEP 4: Test Certificate Generation
// ============================================
function Step4_TestCertificateGeneration() {
  const testData = {
    email: Session.getActiveUser().getEmail(),
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
    const certificateFile = createCertificate(testData);
    Logger.log('✅ Certificate created: ' + certificateFile.getName());
    sendCertificateEmail(testData, certificateFile);
    Logger.log('✅ Email sent to: ' + testData.email);
    Logger.log('🎉 Test completed successfully! Check your email.');
  } catch (error) {
    Logger.log('❌ Test failed: ' + error.toString());
    throw error;
  }
}
