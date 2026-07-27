/**
 * Mill Creek Fab Lab Queue — Apps Script backend
 *
 * Handles file uploads to Google Drive and milestone automated email notifications
 * for status changes (Submitted, Started, Complete, Flagged, Stop requested).
 *
 * Deploy as: Web App
 *   Execute as: Me
 *   Who has access: Anyone
 */

const ROOT_FOLDER_NAME = 'Mill Creek Fab Lab Submissions';

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'uploadFile') {
      return handleUpload(body);
    }

    if (body.action === 'sendStatusEmail') {
      return handleStatusEmail(body);
    }

    if (body.action === 'getStorageUsage') {
      return getStorageUsage();
    }

    return jsonOut({ ok: false, error: 'Unknown action: ' + body.action });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function doGet(e) {
  return jsonOut({ ok: true, message: 'Mill Creek Apps Script backend is running.' });
}

function handleUpload(body) {
  const { submissionId, partId, fileName, fileDataBase64, mimeType } = body;

  if (!submissionId || !fileName || !fileDataBase64) {
    return jsonOut({ ok: false, error: 'Missing required fields (submissionId, fileName, fileDataBase64).' });
  }

  const root = getRootFolder();
  const submissionFolder = getOrCreateSubfolder(root, submissionId);

  const bytes = Utilities.base64Decode(fileDataBase64);
  const blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName);
  const file = submissionFolder.createFile(blob);

  // Anyone with link can view for in-app browser previews
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return jsonOut({
    ok: true,
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId(),
    partId: partId || null,
  });
}

/**
 * Sends automated emails based on job milestone updates
 */
function handleStatusEmail(body) {
  const { recipientEmail, label, status, partId, failureReason, flagMessage, actor } = body;

  if (!recipientEmail || !status || !label) {
    return jsonOut({ ok: false, error: 'Missing required email parameters.' });
  }

  let subject = '';
  let bodyText = '';

  switch (status) {
    case 'Submitted':
      subject = `[Fab Lab] Job Received: ${label}`;
      bodyText = `Hello,\n\nYour job submission "${label}" has been received by the Fab Lab queue. We will notify you as soon as physical processing begins.\n\nThank you,\nMill Creek Fab Lab Team`;
      break;

    case 'Printing/Lasering':
      subject = `[Fab Lab] Job Started: ${label}`;
      bodyText = `Hello,\n\nYour job "${label}" is now on the machine and actively printing/lasering!\n\nThank you,\nMill Creek Fab Lab Team`;
      break;

    case 'Complete':
      subject = `[Fab Lab] Job Complete: ${label}`;
      bodyText = `Hello,\n\nGreat news! Your job "${label}" has completed. Please swing by the Fab Lab during open hours to pick up your part.\n\nThank you,\nMill Creek Fab Lab Team`;
      break;

    case 'Flagged':
      subject = `[Fab Lab] Issue Flagged: ${label}`;
      bodyText = `Hello,\n\nThere was an issue with your job "${label}".\n\nReason: ${failureReason || 'Needs Attention'}\nDetails: ${flagMessage || 'No extra notes provided.'}\n\nPlease visit the Fab Lab or check your dashboard to make updates.\n\nThank you,\nMill Creek Fab Lab Team`;
      break;

    case 'Withdrawn':
      subject = `[Fab Lab] Job Withdrawn: ${label}`;
      bodyText = `Hello,\n\nThis email confirms that your job "${label}" was successfully withdrawn from the queue.\n\nThank you,\nMill Creek Fab Lab Team`;
      break;

    case 'Stop Requested':
      subject = `[Fab Lab] Urgent Stop Requested: ${label}`;
      bodyText = `Hello,\n\nA stop request was submitted for your running job "${label}". A peer leader has been notified to intervene physically.\n\nThank you,\nMill Creek Fab Lab Team`;
      break;

    default:
      return jsonOut({ ok: true, message: 'No email trigger required for status: ' + status });
  }

  try {
    MailApp.sendEmail(recipientEmail, subject, bodyText);
    return jsonOut({ ok: true, message: 'Email sent successfully to ' + recipientEmail });
  } catch (err) {
    return jsonOut({ ok: false, error: 'Failed to send email: ' + err.message });
  }
}

/**
 * Walks the Mill Creek Fab Lab Submissions folder (and every submission
 * subfolder inside it) and totals up file count / bytes used, so the
 * teacher dashboard can show real Drive usage on demand.
 */
function getStorageUsage() {
  const root = getRootFolder();
  let totalBytes = 0;
  let fileCount = 0;
  let submissionFolderCount = 0;

  function walk(folder, isRoot) {
    if (!isRoot) submissionFolderCount++;

    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      totalBytes += f.getSize();
      fileCount++;
    }

    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      walk(subfolders.next(), false);
    }
  }

  walk(root, true);

  return jsonOut({
    ok: true,
    totalBytes,
    fileCount,
    submissionFolderCount,
    rootFolderUrl: root.getUrl(),
  });
}

function getRootFolder() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty('ROOT_FOLDER_ID');

  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (err) {
      // Folder deleted or stale ID
    }
  }

  const folder = DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty('ROOT_FOLDER_ID', folder.getId());
  return folder;
}

function getOrCreateSubfolder(parentFolder, name) {
  const existing = parentFolder.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parentFolder.createFolder(name);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
