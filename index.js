const TESTING_MODE = false; // change to true for tests
const TEST_EMAIL = "matiasl@nassaunationalcable.com"; // admin email for tests

function getSparkPostToken() {
    return PropertiesService.getScriptProperties().getProperty('SPARKPOST_TOKEN') ||
        PropertiesService.getUserProperties().getProperty('SPARKPOST_TOKEN');
}
const SPARKPOST_URL = 'https://api.sparkpost.com/api/v1/transmissions';


const WORKSHEETS = [
    'Offline Orders',
    'ebay, Amazon & Walmart',
    'NNC NES & non-wire'
];

function onEdit(e) {
    try {
        if (!e || !e.source) {
            console.log('No event object or source provided. Make sure this function is called by a trigger.');
            return;
        }

        const sheet = e.source.getActiveSheet();
        const sheetName = sheet.getName();
        if (!WORKSHEETS.includes(sheetName)) {
            return;
        }

        const range = e.range;
        const row = range.getRow();
        const col = range.getColumn();
        const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
        const statusColumnIndex = findColumnIndex(headers, ['status']);

        if (col === statusColumnIndex && range.getValue().toString().toLowerCase() === 'delivered') {

            if (TESTING_MODE) {
                console.log(`Test mode: Status changed to delivered on line ${row} of page ${sheetName}`);
                console.log(`Will send email to: ${TEST_EMAIL}`);
            } else {
                console.log(`Status changed to delivered on row ${row} of page ${sheetName}`);
            }

            processDeliveredOrder(sheet, row);
        }

    } catch (error) {
        console.error('Error on onEdit:', error);
    }
}


function findColumnIndex(headers, possibleNames) {
    for (let i = 0; i < headers.length; i++) {
        const header = headers[i].toString().toLowerCase().trim();
        if (possibleNames.some(name => header.includes(name))) {
            return i + 1;
        }
    }
    return -1;
}


function processDeliveredOrder(sheet, row) {
    try {
        const data = getRowData(sheet, row);

        if (!data.email && !TESTING_MODE) {
            console.log("Couldn't find email for row", row);
            return;
        }
        if (TESTING_MODE && !data.email) {
            data.email = "client-without-email@example.com";
            data.customerName = data.customerName || "Test client";
        }
        if (wasEmailAlreadySent(sheet, row)) {
            console.log('Mail already sent for this row', row);
            return;
        }

        // Small delay to avoid  multiple simultaneous triggers
        Utilities.sleep(1000);

        const customerOrders = getUnsentDeliveredOrdersForCustomer(sheet, data.email);

        if (customerOrders && customerOrders.length > 0) {
            console.log(`Processing ${customerOrders.length} order for ${data.email}`);
            
            //check if an order was already processed by another trigger
            const stillNeedProcessing = customerOrders.filter(order => 
                !wasEmailAlreadySent(sheet, order.row)
            );
            
            if (stillNeedProcessing.length > 0) {
                console.log(`Sending email with ${stillNeedProcessing.length} orders pending`);
                sendDeliveryEmail(data.email, stillNeedProcessing);
                markEmailsAsSent(sheet, stillNeedProcessing);
            } else {
                console.log(`All orders of ${data.email} have been already processed`);
            }
        }

    } catch (error) {
        console.error('Error processing order:', error);
    }
}

function getRowData(sheet, row) {
    const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    const data = {};
    headers.forEach((header, index) => {
        const key = header.toString().toLowerCase().trim();
        data[key] = values[index];
    });

    // En case there are multiple emails
    let emailField = data['email'] || '';
    let primaryEmail = '';

    if (emailField) {
        // Separate  multiple emails and use the first
        const emails = emailField.toString().split(/[,;\s]+/).filter(email =>
            email.trim() && email.includes('@')
        );
        primaryEmail = emails.length > 0 ? emails[0].trim() : '';

        if (emails.length > 1) {
            console.log(`Row ${row}: found multiple emails: ${emails.join(', ')}`);
            console.log(`Using the first: ${primaryEmail}`);
        }
    }

    return {
        po: data['po'] || '',
        date: data['date'] || '',
        type: data['type'] || '',
        customerName: data['contact'] || '',
        qty: data['qty'] || '',
        products: data['product'] || '',
        status: data['status'] || '',
        trackingNumber: data['tracing number'] || data['tracking number'] || '',
        email: primaryEmail,
        allEmails: emailField,
        row: row
    };
}

function wasEmailAlreadySent(sheet, row) {
    const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    let emailSentColumnIndex = findColumnIndex(headers, ['email sent', 'correo enviado', 'email_sent']);

    if (emailSentColumnIndex === -1) {
        emailSentColumnIndex = sheet.getLastColumn() + 1;
        sheet.getRange(2, emailSentColumnIndex).setValue('Email Sent');
    }

    const emailSentValue = sheet.getRange(row, emailSentColumnIndex).getValue();
    return emailSentValue === 'YES' || emailSentValue === 'SÍ' || emailSentValue === true;
}


function getUnsentDeliveredOrdersForCustomer(sheet, customerEmail) {
    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const allData = sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).getValues();

    const customerOrders = [];

    allData.forEach((row, index) => {
        const data = {};
        headers.forEach((header, colIndex) => {
            const key = header.toString().toLowerCase().trim();
            data[key] = row[colIndex];
        });

        const orderData = {
            po: data['po'] || '',
            date: data['date'] || '',
            type: data['type'] || '',
            customerName: data['contact'] || '',
            qty: data['qty'] || '',
            products: data['product'] || '',
            status: data['status'] || '',
            trackingNumber: data['tracing number'] || data['tracking number'] || '',
            email: data['email'] || '',
            row: index + 3
        };

        if (orderData.email === customerEmail &&
            orderData.status.toString().toLowerCase() === 'delivered' &&
            !wasEmailAlreadySent(sheet, orderData.row)) {

            customerOrders.push(orderData);
        }
    });

    return customerOrders;
}

function sendDeliveryEmail(originalEmail, orders) {
    try {
        const sparkpostToken = getSparkPostToken();
        if (!sparkpostToken) {
            console.error('SPARKPOST_TOKEN not configured in PropertiesService');
            return false;
        }

        const customerName = orders[0].customerName || 'Valued Customer';
        const orderNumbers = orders.map(order => order.po).filter(po => po).join('\n');
        const productList = orders.map(order => {
            const qty = order.qty || '';
            const product = order.products || '';
            return qty && product ? `-${product} (${qty})` : `-${product}`;
        }).filter(p => p !== '-').join('\n');
        const trackingNumbers = orders.map(order => order.trackingNumber).filter(t => t).join('\n-');
        const formattedTracking = trackingNumbers ? `-${trackingNumbers}` : 'No tracking numbers available';

        const destinationEmail = TESTING_MODE ? TEST_EMAIL : originalEmail;
        const subjectPrefix = TESTING_MODE ? "[TESTING] " : "";
        const fromEmail = originalEmail.split('@')[0] + '@nnc.NNCNationalCable.com';

        const emailContent = {
            "use_sandbox": false,
            "recipients": [{
                "address": {
                    "email": destinationEmail,
                    "name": TESTING_MODE ? "Tester" : customerName
                }
            }],
            "content": {
                "from": {
                    "email": fromEmail,
                    "name": "NNC Orders"
                },
                "subject": subjectPrefix + "Your order has been delivered",
                "html": generateEmailHTML(customerName, orderNumbers, productList, formattedTracking, originalEmail),
                "text": generateEmailText(customerName, orderNumbers, productList, formattedTracking, originalEmail)
            }
        };

        const options = {
            'method': 'POST',
            'headers': {
                'Authorization': sparkpostToken,
                'Content-Type': 'application/json'
            },
            'payload': JSON.stringify(emailContent)
        };

        console.log(' Sending Email');
        console.log(' Recipient:', destinationEmail);
        console.log(' Subject:', emailContent.content.subject);
        console.log(' From:', emailContent.content.from.email);
        console.log(' SparkPost URL:', SPARKPOST_URL);

        const response = UrlFetchApp.fetch(SPARKPOST_URL, options);
        const responseData = JSON.parse(response.getContentText());
        const responseCode = response.getResponseCode();

        console.log('SparkPost Response Code:', responseCode);
        console.log('SparkPost Response:', JSON.stringify(responseData, null, 2));

        if (responseCode === 200 || responseCode === 202) {
            if (responseData.results && responseData.results.total_accepted_recipients < 1) {
                console.error('SparkPost response: no accepted recipients');
                throw new Error('Email not sent, no accepted recipients');
            }

            const acceptedRecipients = responseData.results ? responseData.results.total_accepted_recipients : 'unknown';
            const rejectedRecipients = responseData.results ? responseData.results.total_rejected_recipients : 'unknown';

            console.log(`SparkPost accepted: ${acceptedRecipients}  recipients`);
            console.log(`SparkPost rejected: ${rejectedRecipients} recipients`);

            if (TESTING_MODE) {
                console.log(`TEST EMAIL sent to : ${destinationEmail}`);
                console.log(`ORIGINAL client email: ${originalEmail}`);
                console.log(`Orders include: ${orderNumbers}`);
            } else {
                console.log('Email sent successfully to:', originalEmail);
            }
            return true;
        } else {
            console.error(' SparkPost Error Response:', response.getContentText());
            console.error(' Response Code:', responseCode);
            return false;
        }

    } catch (error) {
        console.error('Error en sendDeliveryEmail:', error);
        return false;
    }
}

function generateEmailHTML(customerName, orderNumbers, productList, trackingNumbers, originalEmail = '') {
    const testingBanner = TESTING_MODE ? `
    <div style="background-color: #fff3cd; border: 1px solid #ffeaa7; padding: 15px; margin-bottom: 20px; border-radius: 5px;">
      <h3 style="color: #856404; margin: 0;"> TESTING MODE</h3>
      <p style="margin: 5px 0; color: #856404;">
        <strong>Original email for client:</strong> ${originalEmail || 'Not specified'}<br>
        <strong>This email was sent to:</strong> ${TEST_EMAIL}<br>
        <strong>Timestamp:</strong> ${new Date().toLocaleString()}
      </p>
    </div>
  ` : '';

    return `
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
        ${testingBanner}
        
        <h2 style="color: #2c3e50;">Hi ${customerName},</h2>
        
        <p>Your order has been delivered, and we'd like to know if everything met your expectations.</p>
        
        <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin: 20px 0;">
          <h3 style="color: #2c3e50; margin-top: 0;">Order Number's:</h3>
          <p style="font-weight: bold;">${orderNumbers || 'N/A'}</p>
          
          <h3 style="color: #2c3e50;">Products:</h3>
          <div style="white-space: pre-line;">${productList || 'No products specified'}</div>
          
          <h3 style="color: #2c3e50;">Tracking Information:</h3>
          <div style="font-family: monospace; white-space: pre-line;">${trackingNumbers || 'No tracking numbers'}</div>
        </div>
        
        <p>If anything needs attention or you have a question, just reply - we'll take care of it.</p>
        
        <p>Thanks again for choosing NNC. We're looking forward to your next order.</p>
        
        <p style="margin-top: 30px;">
          Best regards,<br>
          <strong>NNC Team</strong>
        </p>
      </div>
    </body>
    </html>
  `;
}


function generateEmailText(customerName, orderNumbers, productList, trackingNumbers, originalEmail = '') {
    const testingHeader = TESTING_MODE ? `
TESTING MODE
Original Client Email: ${originalEmail || 'Not specified'}
The emails was sent to: ${TEST_EMAIL}
Timestamp: ${new Date().toLocaleString()}-------------------` : '';

    return `${testingHeader}Hi ${customerName},

Your order has been delivered, and we'd like to know if everything met your expectations.

Order Number's:
${orderNumbers || 'N/A'}

Products:
${productList || 'No products specified'}

Tracking Information:
${trackingNumbers || 'No tracking numbers'}

If anything needs attention or you have a question, just reply - we'll take care of it.

Thanks again for choosing NNC. We're looking forward to your next order.

Best regards,
NNC Team`;
}


function markEmailsAsSent(sheet, orders) {
    const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    let emailSentColumnIndex = findColumnIndex(headers, ['email sent', 'correo enviado', 'email_sent']);

    if (emailSentColumnIndex === -1) {
        emailSentColumnIndex = sheet.getLastColumn() + 1;
        sheet.getRange(2, emailSentColumnIndex).setValue('Email Sent');
    }

    orders.forEach(order => {
        const markValue = TESTING_MODE ? 'TEST-SENT' : 'YES';

        try {
            const targetCell = sheet.getRange(order.row, emailSentColumnIndex);

            const mergedRanges = sheet.getRange(order.row, emailSentColumnIndex, 1, 1).getMergedRanges();

            if (mergedRanges.length > 0) {
                console.log(`Row ${order.row}: Has a merged cell, marking it all`);
                mergedRanges[0].setValue(markValue);
            } else {
                targetCell.setValue(markValue);
            }

            console.log(` Marking "Email Sent" = "${markValue}" in row ${order.row}`);

        } catch (error) {
            console.error(`Error marking email as sent in row ${order.row}:`, error);
        }
    });
}


function setupTrigger() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'onEdit') {
            ScriptApp.deleteTrigger(trigger);
        }
    });


    if (TESTING_MODE) {
        console.log('Trigger made in testing mode');
        console.log('All emails sent to: ' + TEST_EMAIL);
    } else {
        console.log('Trigger set up for production');
    }
}



// IMPORTANT: Delete this entire block after configuring token for security
// function setSparkPostToken() {
//     const token = 'PASTE_REAL_TOKEN_HERE_THEN_RUN_ONCE_AND_DELETE';
//     PropertiesService.getScriptProperties().setProperty('SPARKPOST_TOKEN', token);
//     console.log('SparkPost token configured successfully');
// }


