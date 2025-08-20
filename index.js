const TESTING_MODE = false; // change to true for tests
const TEST_EMAIL = "matiasl@nassaunationalcable.com"; // admin email for tests

function getSparkPostToken() {
    return PropertiesService.getScriptProperties().getProperty('SPARKPOST_TOKEN') ||
        PropertiesService.getUserProperties().getProperty('SPARKPOST_TOKEN');
}
const SPARKPOST_URL = 'https://api.sparkpost.com/api/v1/transmissions';


const WORKSHEETS = [
    'Offline Orders',
    'eBay, Amazon & Walmart',
    'NNC NES & Non-Wire'
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

    // Check if status column is merged and collect all products from merged rows
    const statusColumnIndex = findColumnIndex(headers, ['status']);
    const productColumnIndex = findColumnIndex(headers, ['product']);
    const qtyColumnIndex = findColumnIndex(headers, ['qty']);
    const poColumnIndex = findColumnIndex(headers, ['po']);
    const trackingColumnIndex = findColumnIndex(headers, ['tracing number', 'tracking number']);
    
    if (statusColumnIndex > 0) {
        const statusRange = sheet.getRange(row, statusColumnIndex);
        const mergedRanges = statusRange.getMergedRanges();
        
        if (mergedRanges.length > 0) {
            const mergedRange = mergedRanges[0];
            const startRow = mergedRange.getRow();
            const numRows = mergedRange.getNumRows();
            
            let allProducts = [];
            let allQtys = [];
            let allPOs = [];
            let allTracking = [];
            
            for (let i = 0; i < numRows; i++) {
                const currentRow = startRow + i;
                const rowValues = sheet.getRange(currentRow, 1, 1, sheet.getLastColumn()).getValues()[0];
                
                if (productColumnIndex > 0 && rowValues[productColumnIndex - 1]) {
                    const product = rowValues[productColumnIndex - 1].toString().trim();
                    if (product && product !== '') allProducts.push(product);
                }
                
                if (qtyColumnIndex > 0 && rowValues[qtyColumnIndex - 1]) {
                    const qty = rowValues[qtyColumnIndex - 1].toString().trim();
                    if (qty && qty !== '') allQtys.push(qty);
                }
                
                if (poColumnIndex > 0 && rowValues[poColumnIndex - 1]) {
                    const po = rowValues[poColumnIndex - 1].toString().trim();
                    if (po && po !== '') allPOs.push(po);
                }
                
                if (trackingColumnIndex > 0 && rowValues[trackingColumnIndex - 1]) {
                    const tracking = rowValues[trackingColumnIndex - 1].toString().trim();
                    if (tracking && tracking !== '') allTracking.push(tracking);
                }
            }
            
            if (allProducts.length > 0) {
                data['product'] = allProducts.join('\n');
            }
            if (allQtys.length > 0) {
                data['qty'] = allQtys.join('\n');
            }
            if (allPOs.length > 0) {
                data['po'] = allPOs.join(',');
            }
            if (allTracking.length > 0) {
                data['tracing number'] = allTracking.join('\n');
                data['tracking number'] = allTracking.join('\n');
            }
        }
    }

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
        const rowNumber = index + 3;
        const orderData = getRowData(sheet, rowNumber);

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
            const products = (order.products || '').split('\n').filter(p => p.trim());
            const qtys = (order.qty || '').toString().split('\n').filter(q => q.trim());
            
            if (products.length > 1) {
                return products.map((product, index) => {
                    const qty = qtys[index] || '';
                    return qty ? `${product} (${qty})` : product;
                }).join('\n');
            } else {
                const qty = order.qty || '';
                const product = order.products || '';
                return qty && product ? `${product} (${qty})` : product;
            }
        }).filter(p => p.trim()).join('\n');
        const trackingNumbers = orders.map(order => order.trackingNumber).filter(t => t).join('\n-');
        const formattedTracking = trackingNumbers ? `-${trackingNumbers}` : 'No tracking numbers available';

        const destinationEmail = TESTING_MODE ? TEST_EMAIL : originalEmail;
        const subjectPrefix = TESTING_MODE ? "[TESTING] " : "";
        const fromEmail = 'noreply@nnc.nncnationalcable.com';

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
                "html": generateEmailHTML(customerName, orderNumbers, productList, formattedTracking, originalEmail, orders),
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

function generateEmailHTML(customerName, orderNumbers, productList, trackingNumbers, originalEmail = '', orders = []) {
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

    const hasSpecialPO = orders.some(order => {
        const po = order.po ? order.po.toString().toUpperCase() : '';
        return po.includes('A') || po.includes('W') || po.includes('E');
    });

    const nnc_info = hasSpecialPO ? '' : `
                        <font face="Tahoma, sans-serif" color="#666666"><b><span
                                            class="il">Phone: 516-482-6313</span></b><br></font>
                                <font face="Tahoma, sans-serif" color="#666666"><a style="text-decoration: none;"
                                        href="https://nassaunationalcable.com">www.nassaunationalcable.com</a>
                        `;

    const garrie_image = hasSpecialPO ? '' : `<img src="https://i.imgur.com/wvxALVs.jpeg" alt="Imagen" width="500">`;
    
    const nassau_footer = hasSpecialPO ? '' : `
                                        <div><span style="font-size:10pt;font-family:Tahoma,sans-serif">
                                                <font color="#666666">Nassau National Cable | NY 11021</font><br>
                                            </span>
                                            <font size="1"><i>
                                                    <font color="#999999">Argentina |&nbsp;Colombia | India | Poland |
                                                        Ukraine | United States | Uruguay</font>
                                                </i></font>
                                        </div>`;

    const nnc_team_header = hasSpecialPO ? '' : `<font face="Tahoma, sans-serif" color="#666666"><b><span
                                            class="il">NNC Customer Service Team</span></b><br></font>`;

    const nnc_logo = hasSpecialPO ? '' : `<img width="96" height="46" src="https://s3.us-east-2.amazonaws.com/starkflow.us/nnc.png"
                style="color:rgb(34,34,34)" class="CToWUd" data-bit="iit">`;

    const body = `
                    <strong>Hi ${customerName}, </strong>
                    <br><br>
                    Your order has been delivered, and we'd like to know if everything met your expectations.
                    <br><br>
                    <strong>Order Number${orderNumbers.includes('<br>') || orderNumbers.includes(',') ? "'s" : ""}: </strong><br>
                    ${orderNumbers.replace(/,/g, '<br>')}
                    <br><br>
                    <strong>Product${productList.includes("<br>-") ? "s" : ""}:</strong><br>
                    -${productList}
                    <br><br>
                    <strong>Tracking Information:</strong><br>
                    -${trackingNumbers}
                    <br><br>
                    If anything needs attention or you have a question, just reply - we'll take care of it.
                    <br><br>
                    Thanks again for choosing NNC. We're looking forward to your next order.
                    <br><br>
                    Best regards,
                    <br>
                    `;

    const body_html = `<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
</head>

<body><br>
    ${testingBanner}
    <p style="font-size:10pt;font-family:Tahoma,sans-serif"><span>${body}</span></p><br>
    <div dir="ltr" class="gmail_signature" data-smartmail="gmail_signature">
        <div dir="ltr">${nnc_logo}
            <font color="#888888"><br style="color:rgb(34,34,34)">
                <table cellpadding="0" cellspacing="0"
                    style="color:rgb(34,34,34);background-color:transparent;border-spacing:0px;border-collapse:collapse">
                    <tbody>
                        <tr><br>
                            <td style="padding:0px">
                                <font color="#00006f" face="Tahoma, sans-serif"><b><span class="il"></span></b>
                                </font>
                                ${nnc_team_header}
                                ${nnc_info}
                                    <font face="Tahoma, sans-serif" color="#666666">
                                        ${nassau_footer}
                            </td>
                        </tr>
                    </tbody>
                </table><a style="text-decoration:none;" href="https://www.linkedin.com/company/nassaunationalcable/"
                    target="_blank"
                    data-saferedirecturl="https://www.google.com/url?q=https://www.linkedin.com/company/nassaunationalcable/&amp;source=gmail&amp;ust=1687284573416000&amp;usg=AOvVaw0lv_0znO6DQioGdUZ5jHg0"><img
                        src="https://ci3.googleusercontent.com/mail-sig/AIorK4x9YsL6RScQKGevrjyhcIi7-QMOZEumdVELt78Y3TR9C0bLAo0nVu0TDBOx55ab_YK1jxAHGas"
                        class="CToWUd" data-bit="iit"></a>&nbsp;&nbsp;&nbsp;<a style="text-decoration:none;"
                    href="https://www.facebook.com/nassaunationalcable" target="_blank"
                    data-saferedirecturl="https://www.google.com/url?q=https://www.facebook.com/nassaunationalcable&amp;source=gmail&amp;ust=1687284573416000&amp;usg=AOvVaw3zpnI-XABSj--PVCVDKScF"><img
                        src="https://ci3.googleusercontent.com/mail-sig/AIorK4znOVUFqstVsa5DSxVa0rKoiEIvb6FHWD0HyM-zUtl0ZRttlAVLsPODbADYpV_-8XNrNlUomp8"
                        class="CToWUd" data-bit="iit"></a>&nbsp;&nbsp;&nbsp;<a style="text-decoration:none;"
                    href="https://twitter.com/CableNassau?s=20" target="_blank"
                    data-saferedirecturl="https://www.google.com/url?q=https://twitter.com/CableNassau?s%3D20&amp;source=gmail&amp;ust=1687284573416000&amp;usg=AOvVaw3PGIaXzeJbZAbEGGeMNiuq"><img
                        src="https://ci3.googleusercontent.com/mail-sig/AIorK4xOM0kDPDtiqMXZfm0Yb_C-x5ALOWNC9VRd2_dghiINYuQvivUa6l0OdhP1Z0fDmkDRJFz5GG0"
                        class="CToWUd" data-bit="iit"></a>&nbsp;&nbsp;&nbsp;<a style="text-decoration:none;"
                    href="https://www.instagram.com/nassaunationalcable/" target="_blank"
                    data-saferedirecturl="https://www.google.com/url?q=https://www.instagram.com/nassaunationalcable/&amp;source=gmail&amp;ust=1687284573416000&amp;usg=AOvVaw3cqkHlGXuHKipSCcTRMlNl"><img
                        src="https://ci3.googleusercontent.com/mail-sig/AIorK4y7fNWamYzfOQaXBGuk8aN4sUhoeHPtHHBU9RS6Ws7rEOR_YthRQ3dINQcENzM_MZe6LvxV0N4"
                        class="CToWUd" data-bit="iit"></a>&nbsp; &nbsp;<a style="text-decoration:none;"
                    href="https://www.youtube.com/channel/UCDKDyoBSSi9W3QgGMySDyAQ" target="_blank"
                    data-saferedirecturl="https://www.google.com/url?q=https://www.youtube.com/channel/UCDKDyoBSSi9W3QgGMySDyAQ&amp;source=gmail&amp;ust=1687284573416000&amp;usg=AOvVaw0WBWlyFAAax-X4h2ySTxQ6"><img
                        src="https://ci3.googleusercontent.com/mail-sig/AIorK4xTKiZnhvQJaL0SPCNMxqfQl9Nh44LK9R2Ckj0-z3nhkxz64wWkmmFTfL49Q_LCsCvLvGMF_fM"
                        class="CToWUd" data-bit="iit"></a><br>
            </font><br><br>
        </div>
    </div>
    ${garrie_image}
</body>

</html>`;

    return body_html;
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

