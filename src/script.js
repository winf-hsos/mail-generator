console.log("E-Mail Generator v0.2");

const queryString = window.location.search;
const urlParams = new URLSearchParams(queryString);
var sheetKey = urlParams.get('sheetkey');

var meta, user, emails, index;

var isIndex = false;

// Initialize preview modal
var previewModal;

if (sheetKey === null) {
    console.error("Please provide sheet key via url.")
    noKey();

} else {

    document.getElementById("linkToMailSheet").setAttribute("href", `https://docs.google.com/spreadsheets/d/${sheetKey}/edit`);
    document.getElementById("linkToIndexSheet").setAttribute("href", `https://docs.google.com/spreadsheets/d/${sheetKey}/edit`);
    getData(sheetKey).then(run);
}

async function getData(sheetKey) {
    return Promise.all([
        readSheetData(sheetKey, 1),
        readSheetData(sheetKey, 2),
        readSheetData(sheetKey, 3),
    ])
}

function noKey() {
    let sectionInputKey = document.getElementById("sheetKeyPrompt");
    sectionInputKey.removeAttribute("hidden");

    let sectionEmailList = document.getElementById("emailList");
    sectionEmailList.setAttribute("hidden", "");
}

function submitSheetKey() {
    let inputSheetKey = document.getElementById("inputSheetKey");
    sheetKey = inputSheetKey.value;
    setSheetKey(sheetKey);
}

function openMailSheet(btn) {
    let sheetKey = btn.dataset.sheetKey;
    setSheetKey(sheetKey);
}

function backToIndex() {
    var indexSheetKey = urlParams.get('indexkey');
    setSheetKey(indexSheetKey);
}

function setSheetKey(newSheetKey) {
    if (newSheetKey !== "") {
        if (index) {
            window.location.href = "?sheetkey=" + newSheetKey + "&indexkey=" + sheetKey;
        }
        else {
            window.location.href = "?sheetkey=" + newSheetKey;
        }
    }
}

function tryTemplateSheet() {
    window.location.href = "?sheetkey=1l1ZvHf3lCBFLUMFXsDCfhEw_z1u1UD63aubOiS8JPSk";
}

function tryIndexTemplateSheet() {
    window.location.href = "?sheetkey=1-Zx7KdBwFyh0-FMm1pTlZFlNlxeipU2W9eyjDQIn2L0";
}

function run(data) {

    // Index sheet, if first (and only) sheet is named "Index"
    if (data[0].title === "Index") {
        console.info("This is an index sheet.");
        data = [data.shift()];
        isIndex = true;
    }

    let sectionInputKey = document.getElementById("sheetKeyPrompt");
    sectionInputKey.setAttribute("hidden", "");

    data = _makeDictFromData(data)

    if (!isIndex) {
        mailSheet(data);

        let sectionEmailList = document.getElementById("emailList");
        sectionEmailList.removeAttribute("hidden");
        
        if (urlParams.get('indexkey') !== null)
            document.getElementById("btnBackToIndex").removeAttribute("hidden");
    }
    else {
        indexSheet(data);

        let sectionEmailIndexList = document.getElementById("indexList");
        sectionEmailIndexList.removeAttribute("hidden");

    }

    app = new Vue({
        el: '#app',
        data: {
            emails: emails,
            indexSheets: index
        }
    })
}


// Run if this is a mail sheet
function mailSheet(data) {

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
}

// Run if this is an index sheet
function indexSheet(data) {
    index = [];
    let indexRaw = data["index"];
    for (let i = 0; i < indexRaw.length; i++) {
        index.push(indexRaw[i]);
    }
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

function _fillPreview(mail) {
    let previewSubjectElement = document.getElementById("previewSubject");
    previewSubjectElement.innerHTML = mail.subject;

    let previewContentElement = document.getElementById("previewContent");
    previewContentElement.innerHTML = mail.contentHtml;

    let previewDownloadBtnElement = document.getElementById("previewDownloadBtn");
    previewDownloadBtnElement.dataset.mailId = mail.id;

    let previewCopyContentBtnElement = document.getElementById("previewCopyContentBtn");
    previewCopyContentBtnElement.dataset.mailId = mail.id;
}

function previewEmail(btn) {
    let mailId = btn.dataset.mailId;
    let mail = emails[mailId];
    _fillPreview(mail);
    _previewEmail(mail);
}

function copyMailFormattedToClipboard(btn) {
    let selection = _selectText("previewContent");
    document.execCommand("copy");
    selection.removeAllRanges();
}

function _selectText(elementId) {
    var doc = document
    var text = doc.getElementById(elementId);
    var range, selection;

    if (doc.body.createTextRange) {
        range = document.body.createTextRange();
        range.moveToElementText(text);
        range.select();
    } else if (window.getSelection) {
        selection = window.getSelection();
        range = document.createRange();
        range.selectNodeContents(text);
        selection.removeAllRanges();
        selection.addRange(range);
        return selection;
    }
}

function _createEmail(to, subject, content, cc = null, bcc = null) {

    console.dir(content)

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

    // Open dialogue
    previewModal.toggle();
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