console.log("E-Mail Generator v0.2");

const queryString = window.location.search;
const urlParams = new URLSearchParams(queryString);
const sheetKey = urlParams.get('sheetkey');

var previewModal;

async function getData(sheetKey) {
    return Promise.all([
        readSheetData(sheetKey, 1),
        readSheetData(sheetKey, 2),
        readSheetData(sheetKey, 3),
    ])
}

getData(sheetKey).then(run);

var meta, user, emails;
function run(data) {

    data = _makeDictFromData(data)

    // Prepare meta data
    meta = data["meta"];
    metaDict = {}
    for (let i = 0; i < meta.length; i++) {
        metaDict[meta[i]["key"]] = meta[i]["value"]
    }
    meta = metaDict;

    // Prepare the user
    user = data["user"];

    // Prepare the mails
    emails = data["emails"];

    for (let i = 0; i < emails.length; i++) {

        let mail = emails[i];

        let subject_template = Handlebars.compile(mail.subject);
        mail.subject = subject_template(meta);


        let content_template = Handlebars.compile(mail.content);
        mail.content = content_template(meta);

        converter = new showdown.Converter();
        mail.contentHtml = converter.makeHtml(mail.content);

        // To whom send this mail?
        let toString = mail.to;
        let toSet = _resolveRecipientsByTag(toString);

        // To whom not to send this mail?
        let notToString = mail.notto;
        let notToSet = _resolveRecipientsByTag(notToString);

        // Subtract both sets (to)
        toSet = _subtractSets(toSet, notToSet);
        mail.to = _emailSetToString(toSet);

        // Who is in cc?
        let ccString = mail.cc;
        let ccSet = _resolveRecipientsByTag(ccString);

        // Who is not in cc?
        let notCcString = mail.notcc;
        let notCcSet = _resolveRecipientsByTag(notCcString);

        // Subtract both sets (cc)
        ccSet = _subtractSets(ccSet, notCcSet);
        ccString = _emailSetToString(ccSet);
        mail.cc = ccString == "" ? null : ccString;

        // And who is in bcc?
        let bccString = mail.bcc;
        let bccSet = _resolveRecipientsByTag(bccString);

        // Who is not in bcc?
        let notBccString = mail.notbcc;
        let notBccSet = _resolveRecipientsByTag(notBccString);

        // Subtract bot sets (bcc)
        bccSet = _subtractSets(bccSet, notBccSet);
        bccString = _emailSetToString(bccSet);
        mail.bcc = bccString == "" ? null : bccString;

    }

    app = new Vue({
        el: '#app',
        data: {
            emails: emails,
            activemail: emails[0]
        }
    })
}

async function readSheetData(workbookId, sheetNumber) {
    let values;
    let json;

    try {
        values = await fetch('https://spreadsheets.google.com/feeds/list/' + workbookId + '/' + sheetNumber + '/public/values?alt=json');
        json = await values.json();
    }
    catch (error) {
        if (error.name === 'FetchError') {
            console.error("No data returned. Maybe sheet not published to web, wrong workbook ID, or sheet " + sheetNumber + " does not exist in sheet?");
        }

        return { "error": "No data returned. Maybe sheet not published to web, wrong workbook ID, or sheet " + sheetNumber + " does not exist in sheet?" };
    }

    let rows = json.feed.entry;

    let data = {};
    data.title = json.feed.title['$t']
    let dataRows = [];

    if (rows) {
        for (let i = 0; i < rows.length; i++) {
            let row = rows[i];
            let rowObj = {}
            for (column in row) {
                if (column.startsWith('gsx$')) {
                    let columnName = column.split("$")[1];
                    rowObj[columnName] = row[column]["$t"];
                }
            }
            dataRows.push(rowObj);
        }
    }

    data.rows = dataRows;
    return data;
}

function downloadEmail(btn) {
    let mailId = btn.dataset.mailId;
    let mail = emails[mailId];

    _createEmail(mail.to, mail.subject, mail.contentHtml, mail.cc, mail.bcc);
}

function previewEmail(btn) {
    let mailId = btn.dataset.mailId;
    let mail = emails[mailId];

    let previewSubjectElement = document.getElementById("previewSubject");
    previewSubjectElement.innerHTML = mail.subject;
    
    let previewContentElement = document.getElementById("previewContent");
    previewContentElement.innerHTML = mail.contentHtml;

    let previewDownloadBtnElement = document.getElementById("previewDownloadBtn");
    previewDownloadBtnElement.dataset.mailId = mail.id;

    _previewEmail(mail);
}

function _createEmail(to, subject, content, cc = null, bcc = null) {

    var emlContent = "data:message/rfc822 eml,";
    emlContent += 'To: ' + to + '\n';
    if (cc !== null)
        emlContent += 'Cc: ' + cc + '\n';
    if (bcc !== null)
        emlContent += 'Bcc: ' + bcc + '\n';
    emlContent += 'Subject: ' + subject + '\n';
    emlContent += 'X-Unsent: 1' + '\n';
    emlContent += 'Content-Type: text/html; charset=UTF-8' + '\n';
    emlContent += '' + '\n';
    emlContent += content;

    _createAndDownloadMail(emlContent, subject + '.eml')
}

function _createAndDownloadMail(content, fileName) {
    var encodedUri = encodeURI(content);
    var a = document.createElement('a');
    var linkText = document.createTextNode("fileLink");
    a.appendChild(linkText);
    a.href = encodedUri;
    a.id = 'fileLink';
    a.download = fileName;
    a.style = "display:none;";
    document.body.appendChild(a);
    document.getElementById('fileLink').click();
    document.body.removeChild(a);
}

function _previewEmail(mail) {

    previewModal = new bootstrap.Modal(document.getElementById('previewModal'), {
        keyboard: true
    });

    // Set the mail content for preview

    previewModal.toggle();

    // Open dialogue
}

function closePreview() {

    previewModal.toggle();

}


function _makeDictFromData(data) {
    dataDict = {};

    // Go through sheets
    for (let i = 0; i < data.length; i++) {

        let entries = [];
        // Go through each row in a sheet
        for (let j = 0; j < data[i].rows.length; j++) {
            rowObj = _parseProperties(data[i].rows[j])
            rowObj.id = j;
            entries.push(rowObj);
        }

        dataDict[data[i].title.toLowerCase()] = entries;
    }

    return dataDict;
}

function _parseProperties(obj) {

    result = {}
    for (var key in obj) {

        if (!obj.hasOwnProperty(key)) continue;

        if (key === "name")
            result.name = obj[key];
        else if (["TRUE", "FALSE"].includes(obj[key])) {
            result[key] = obj[key] === "TRUE" ? true : false;
        }
        else {
            result[key] = obj[key];
        }
    }

    return result;
}

function _resolveRecipientsByTag(tagString) {
    let emails = new Set();

    if (tagString === "")
        return emails;

    let tags = tagString.replace(/ /g, '').split(";")

    for (let i = 0; i < tags.length; i++) {
        let emailsForTag = _resolveEmailsForTag(tags[i]);
        emails = _unionSets(emails, emailsForTag);
    }
    return emails;
}

function _resolveEmailsForTag(tag) {
    let emails = new Set()

    if (_isEmail(tag) === true) {
        emails.add(tag);
        return emails;
    }

    for (let i = 0; i < user.length; i++) {
        if (user[i][tag] && user[i][tag] === true) {
            emails.add(user[i].email);
        }
    }

    return emails;
}

function _unionSets(setA, setB) {
    let _union = new Set(setA)
    for (let elem of setB) {
        _union.add(elem)
    }
    return _union
}

function _subtractSets(setA, setB) {
    let _difference = new Set(setA)
    for (let elem of setB) {
        _difference.delete(elem)
    }
    return _difference
}

function _emailSetToString(emailSet) {
    emailString = "";
    for (let e of emailSet) {
        emailString += e + ";";
    }

    return emailString.slice(0, -1);
}

function _isEmail(tagString) {
    return tagString.includes("@");
}