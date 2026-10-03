/**
 * Server-side code for Roy-Hart Portfolio Portal
 * Updated to match the "Roy-Hart Writing Portfolio Portal Master Sheet"
 */

const CONFIG = {
  FOLDER_NAME: 'Roy-Hart Portfolios',
  SHEETS: {
    SETTINGS: 'Settings',
    STUDENTS: 'Students',
    TEACHERS: 'Teachers',
    ARTIFACTS: 'Artifacts',
    ASSIGNMENTS: 'Assignments',
    AUDIT: 'Audit Log'
  }
};

// ==========================================
// 1. WEB APP CORE
// ==========================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Roy-Hart Portfolio Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// 2. SETUP & DATABASE INITIALIZATION
// ==========================================

function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const createSheet = (name, headers) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }
  };

  createSheet(CONFIG.SHEETS.SETTINGS, ['Setting Key', 'Value']);
  createSheet(CONFIG.SHEETS.TEACHERS, ['Teacher Name', 'Teacher Email', 'Role', 'Active']);
  createSheet(CONFIG.SHEETS.STUDENTS, ['Student Name', 'Student Email', 'Grad Year', 'Current Grade', 'Portfolio Folder URL', 'Portfolio Folder ID', 'Active']);
  createSheet(CONFIG.SHEETS.ARTIFACTS, ['ID', 'Student Name', 'Student Email', 'Grad Year', 'School Year', 'Grade Level', 'Title', 'Type', 'Assignment Name', 'File URL', 'Drive File ID', 'Best Work', 'Date Submitted', 'Last Updated', 'Reflection', 'Teacher Notes', 'Tagged Standards', 'Visibility']);
  
  // Updated Assignment Headers to include Creator information
  createSheet(CONFIG.SHEETS.ASSIGNMENTS, ['Assignment Name', 'Grade Level', 'Type', 'Description', 'Required', 'Due Date', 'Creator Name', 'Creator Email']);
  createSheet(CONFIG.SHEETS.AUDIT, ['Timestamp', 'User Email', 'Action', 'Artifact ID', 'Notes']);
  
  Logger.log("Setup complete! Your database is ready and matches the Master Sheet structure.");
}

// ==========================================
// 3. HELPER METHODS
// ==========================================

function getRootFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  
  if (settingsSheet) {
    const data = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'ROOT_FOLDER_ID' && data[i][1]) {
        try {
          return DriveApp.getFolderById(data[i][1].toString().trim());
        } catch(e) {
          Logger.log("Could not find folder by ROOT_FOLDER_ID. Falling back to search by name.");
        }
      }
    }
  }

  const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.FOLDER_NAME);
}

function formatStudentName(fullName) {
  if (!fullName) return null;
  const parts = fullName.trim().split(/\s+/);
  if (parts.length === 1) return parts[0];
  const last = parts.pop();
  const first = parts.join(' ');
  return `${last}_${first}`;
}

/**
 * Escapes HTML special characters in a string so it's safe to interpolate
 * into HTML email bodies, server-rendered HTML, etc.
 */
function escapeHtmlServer(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * SECURITY PATCH: Verifies if the active session belongs to an authorized teacher.
 */
function isTeacher(email) {
  if (!email) return false;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teacherSheet = ss.getSheetByName(CONFIG.SHEETS.TEACHERS);
  if (!teacherSheet) return false;
  
  const data = teacherSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().toLowerCase() === email.toLowerCase()) {
      return true;
    }
  }
  return false;
}

// ==========================================
// 4. API METHODS (Called from Frontend)
// ==========================================

function getUserState() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return { status: 'error', message: 'Could not detect user email.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Teachers are authorized automatically by checking the teacher sheet
  if (isTeacher(email)) {
    const teacherSheet = ss.getSheetByName(CONFIG.SHEETS.TEACHERS);
    const data = teacherSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().toLowerCase() === email.toLowerCase()) {
         return { status: 'success', role: 'teacher', email: email, name: data[i][0] };
      }
    }
  }

  const studentSheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
  if (studentSheet) {
    const studentData = studentSheet.getDataRange().getValues();
    for (let i = 1; i < studentData.length; i++) {
      if (studentData[i][1] && studentData[i][1].toString().toLowerCase() === email.toLowerCase()) {
        return { 
          status: 'success', 
          role: 'student', 
          email: email, 
          name: studentData[i][0],
          gradYear: studentData[i][2]
        };
      }
    }
  }

  return { status: 'unauthorized', email: email };
}

/**
 * Gets a list of all teachers for the "Notify Teacher" dropdown
 */
function getActiveTeachers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TEACHERS);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const teachers = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      // Assuming column 1 is Name, column 2 is Email
      teachers.push({ name: data[i][0], email: data[i][1] });
    }
  }
  
  return teachers.sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Fetches Settings and Assignments to lock down the form for students
 */
function getFormConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let currentYear = new Date().getFullYear() + '/' + String(new Date().getFullYear() + 1).slice(-2);
  
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (settingsSheet) {
    const data = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'CURRENT_SCHOOL_YEAR' && data[i][1]) {
        currentYear = data[i][1].toString().trim();
        break;
      }
    }
  }

  const assignments = [];
  const assignSheet = ss.getSheetByName(CONFIG.SHEETS.ASSIGNMENTS);
  if (assignSheet) {
    const data = assignSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      assignments.push({
        name: data[i][0],
        gradeLevel: data[i][1],
        type: data[i][2],
        description: data[i][3] || '',
        required: data[i][4] || false,
        dueDate: data[i][5] instanceof Date ? data[i][5].toISOString() : (data[i][5] || ''),
        creatorName: data[i][6] || 'Unknown',
        creatorEmail: data[i][7] || ''
      });
    }
  }

  return { currentYear, assignments };
}

/**
 * Creates a new assignment from the Teacher Dashboard
 */
function addAssignment(details) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!isTeacher(userEmail)) {
      throw new Error("Unauthorized access attempt logged.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Look up the creator's real name from the Teacher sheet
    let creatorName = "Teacher";
    const teacherSheet = ss.getSheetByName(CONFIG.SHEETS.TEACHERS);
    if (teacherSheet) {
      const tData = teacherSheet.getDataRange().getValues();
      for (let i = 1; i < tData.length; i++) {
        if (tData[i][1] && tData[i][1].toString().toLowerCase() === userEmail.toLowerCase()) {
          creatorName = tData[i][0];
          break;
        }
      }
    }

    const sheet = ss.getSheetByName(CONFIG.SHEETS.ASSIGNMENTS);
    if (!sheet) throw new Error("Assignments sheet not found.");

    sheet.appendRow([
      details.name,
      details.gradeLevel || '',
      details.type || '',
      details.description || '',
      details.required || false,
      details.dueDate || '',
      creatorName,
      userEmail
    ]);
    
    return { status: 'success', message: 'Assignment created successfully!' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

/**
 * Deletes an assignment by name
 */
function deleteAssignment(assignmentName) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!isTeacher(userEmail)) {
      throw new Error("Unauthorized: Only teachers may delete assignments.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ASSIGNMENTS);
    if (!sheet) throw new Error("Assignments sheet not found.");
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === assignmentName.toLowerCase()) {
        sheet.deleteRow(i + 1);
        
        const auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
        if (auditSheet) {
          const now = new Date().toISOString().replace('T', ' ').substring(0, 19);
          auditSheet.appendRow([now, userEmail, 'Delete Assignment', assignmentName, '']);
        }
        
        return { status: 'success', message: 'Assignment deleted successfully' };
      }
    }
    throw new Error('Assignment not found.');
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function getStudentArtifacts(studentEmail) {
  const activeEmail = Session.getActiveUser().getEmail();
  
  // SECURITY PATCH: Only allow the student themselves OR a teacher to view these artifacts.
  if (activeEmail.toLowerCase() !== studentEmail.toLowerCase() && !isTeacher(activeEmail)) {
    throw new Error("Unauthorized access attempt logged.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ARTIFACTS);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const artifacts = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().toLowerCase() === studentEmail.toLowerCase()) {
      artifacts.push({
        id: data[i][0],
        studentName: data[i][1],
        studentEmail: data[i][2],
        gradYear: data[i][3],
        schoolYear: data[i][4],
        gradeLevel: data[i][5],
        title: data[i][6],
        type: data[i][7],
        assignmentName: data[i][8],
        url: data[i][9],
        driveFileId: data[i][10],
        bestWork: data[i][11],
        dateSubmitted: data[i][12] instanceof Date ? data[i][12].toISOString() : data[i][12],
        lastUpdated: data[i][13] instanceof Date ? data[i][13].toISOString() : data[i][13],
        reflection: data[i][14] || '',
        teacherNotes: data[i][15] || '',
        taggedStandards: data[i][16] || '',
        visibility: data[i][17] || ''
      });
    }
  }
  return artifacts.reverse();
}

function getAllArtifacts() {
  const userEmail = Session.getActiveUser().getEmail();
  if (!isTeacher(userEmail)) {
    throw new Error("Unauthorized access attempt logged.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ARTIFACTS);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const artifacts = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; 
    
    artifacts.push({
      id: data[i][0],
      studentName: data[i][1],
      studentEmail: data[i][2],
      gradYear: data[i][3],
      schoolYear: data[i][4],
      gradeLevel: data[i][5],
      title: data[i][6],
      type: data[i][7],
      assignmentName: data[i][8],
      url: data[i][9],
      driveFileId: data[i][10],
      bestWork: data[i][11],
      dateSubmitted: data[i][12] instanceof Date ? data[i][12].toISOString() : data[i][12],
      lastUpdated: data[i][13] instanceof Date ? data[i][13].toISOString() : data[i][13],
      reflection: data[i][14] || '',
      teacherNotes: data[i][15] || '',
      taggedStandards: data[i][16] || '',
      visibility: data[i][17] || ''
    });
  }
  return artifacts.reverse();
}

function getAllStudents() {
  const userEmail = Session.getActiveUser().getEmail();
  if (!isTeacher(userEmail)) {
    throw new Error("Unauthorized access attempt logged.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const students = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; 

    students.push({
      name: data[i][0],
      email: data[i][1],
      gradYear: data[i][2],
      currentGrade: data[i][3],
      folderUrl: data[i][4],
      folderId: data[i][5],
      active: data[i][6]
    });
  }
  
  // Sort students alphabetically by Last Name, then First Name
  students.sort((a, b) => {
    const splitName = (fullName) => {
      if (!fullName) return { first: "", last: "" };
      const parts = fullName.trim().split(/\s+/);
      if (parts.length === 1) return { first: "", last: parts[0].toLowerCase() };
      const last = parts.pop().toLowerCase();
      const first = parts.join(" ").toLowerCase();
      return { first, last };
    };
    
    const nameA = splitName(a.name);
    const nameB = splitName(b.name);
    
    if (nameA.last < nameB.last) return -1;
    if (nameA.last > nameB.last) return 1;
    if (nameA.first < nameB.first) return -1;
    if (nameA.first > nameB.first) return 1;
    return 0;
  });

  return students;
}

/**
 * Extracts a Google Drive file ID from any common Drive/Docs URL format.
 * Returns null if no ID can be found.
 */
function extractDriveFileId(url) {
  if (!url) return null;
  // Matches /d/FILE_ID/ (Docs, Sheets, Slides, Drive file viewer)
  let m = url.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  // Matches ?id=FILE_ID or &id=FILE_ID (older Drive open links)
  m = url.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  return null;
}

/**
 * Sets a file to "anyone with the link can view". Falls back to domain-only
 * if the Workspace admin blocks public link sharing, so artifacts always end
 * up viewable by at least the school community rather than silently private.
 */
function setArtifactSharing(file) {
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e1) {
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      Logger.log("ANYONE_WITH_LINK blocked by domain policy; used DOMAIN_WITH_LINK for " + file.getId());
    } catch (e2) {
      Logger.log("Could not set sharing for file " + file.getId() + ": " + e2);
    }
  }
}

/**
 * Pre-flight check called from the upload form. Verifies the script can
 * actually read the linked file before the student fills out the rest of
 * the form. Returns { status: 'success' | 'error', message: string }.
 */
function checkLinkAccess(url) {
  try {
    if (!url) {
      return { status: 'error', message: 'Please paste a link first.' };
    }
    const trimmed = url.toLowerCase().trim();
    if (!trimmed.startsWith("https://docs.google.com/") && !trimmed.startsWith("https://drive.google.com/")) {
      return { status: 'error', message: 'Only Google Docs or Drive links are allowed.' };
    }
    const fileId = extractDriveFileId(url);
    if (!fileId) {
      return { status: 'error', message: 'Could not read that link. Make sure you copied the full URL.' };
    }
    // Touching the file forces a permission check. If we can't read it,
    // this throws and we report a friendly fix-it message.
    const file = DriveApp.getFileById(fileId);
    file.getName();
    return { status: 'success', message: 'Access confirmed.' };
  } catch (err) {
    Logger.log("checkLinkAccess failed for " + url + ": " + err);
    return {
      status: 'error',
      message: 'Sharing is not set correctly. Open the doc → Share → set "Anyone with the link" to Viewer.'
    };
  }
}

/**
 * Handles both traditional file uploads AND Google Doc Links
 */
function uploadArtifact(base64Data, filename, mimeType, artifactDetails) {
  try {
    const activeEmail = Session.getActiveUser().getEmail();
    
    // SECURITY PATCH 1: Identity Protection
    // Force the active user's email into the payload, preventing parameter tampering.
    artifactDetails.studentEmail = activeEmail;

    // SECURITY PATCH 2: Unvalidated Inputs / Drive-By Links
    if (artifactDetails.linkUrl) {
      const url = artifactDetails.linkUrl.toLowerCase().trim();
      if (!url.startsWith("https://docs.google.com/") && !url.startsWith("https://drive.google.com/")) {
         throw new Error("Security block: Only secure Google Workspace links are allowed.");
      }
    }

    // --- Process Folder Creation ---
    const rootFolder = getRootFolder();
    
    // Grad Folder
    const gradYearName = artifactDetails.gradYear ? `Class of ${artifactDetails.gradYear}` : 'Unknown Class Year';
    let gradFolder;
    const gradFolders = rootFolder.getFoldersByName(gradYearName);
    if (gradFolders.hasNext()) {
      gradFolder = gradFolders.next();
    } else {
      gradFolder = rootFolder.createFolder(gradYearName);
    }

    // Student Folder
    const formattedName = formatStudentName(artifactDetails.studentName);
    const studentFolderName = formattedName || artifactDetails.studentEmail;
    
    let studentFolder;
    const studentFolders = gradFolder.getFoldersByName(studentFolderName);
    
    if (studentFolders.hasNext()) {
      studentFolder = studentFolders.next();
    } else {
      studentFolder = gradFolder.createFolder(studentFolderName);
      
      // Share with student silently
      if (artifactDetails.studentEmail) {
        try {
          Drive.Permissions.insert({ role: 'writer', type: 'user', value: artifactDetails.studentEmail }, studentFolder.getId(), { sendNotificationEmails: false });
        } catch (errV2) {
          try {
            Drive.Permissions.create({ role: 'writer', type: 'user', emailAddress: artifactDetails.studentEmail }, studentFolder.getId(), { sendNotificationEmail: false });
          } catch (errV3) {
            studentFolder.addEditor(artifactDetails.studentEmail);
          }
        }
        
        // Record back to Students tab
        try {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          const studentSheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
          if (studentSheet) {
            const studentData = studentSheet.getDataRange().getValues();
            for (let i = 1; i < studentData.length; i++) {
              if (studentData[i][1] && studentData[i][1].toString().toLowerCase() === artifactDetails.studentEmail.toLowerCase()) {
                studentSheet.getRange(i + 1, 5).setValue(studentFolder.getUrl()); 
                studentSheet.getRange(i + 1, 6).setValue(studentFolder.getId());
                break;
              }
            }
          }
        } catch (sheetErr) {
          Logger.log("Could not update student sheet: " + sheetErr);
        }
      }
    }

    // Year Folder
    const schoolYearFolderStr = (artifactDetails.schoolYear || 'Unknown Year').replace(/\//g, '-');
    let yearFolder;
    const yearFolders = studentFolder.getFoldersByName(schoolYearFolderStr);
    
    if (yearFolders.hasNext()) {
      yearFolder = yearFolders.next();
    } else {
      yearFolder = studentFolder.createFolder(schoolYearFolderStr);
    }

    // --- Determine if we are uploading a file or saving a link ---
    let fileUrl = '';
    let fileId = '';

    if (base64Data) {
      // SECURITY PATCH 3: Storage Bomb Prevention (File Size Cap)
      // Reject files larger than ~25 MB to prevent storage abuse.
      // base64 inflates size by ~33%, so a 25 MB file is ~33.3 MB encoded.
      const MAX_FILE_BYTES = 25 * 1024 * 1024;
      const approxBytes = Math.floor(base64Data.length * 0.75);
      if (approxBytes > MAX_FILE_BYTES) {
        throw new Error("File is too large. Please upload a file under 25 MB, or share it as a Google Doc link instead.");
      }

      const data = Utilities.base64Decode(base64Data);
      const blob = Utilities.newBlob(data, mimeType, filename);
      const file = yearFolder.createFile(blob);
      setArtifactSharing(file);
      fileUrl = file.getUrl();
      fileId = file.getId();
    } else if (artifactDetails.linkUrl) {
      // Copy the linked Google file into the shared Drive folder so the school
      // owns a permanent snapshot with controlled sharing. The student's
      // original stays untouched in their own Drive.
      const sourceId = extractDriveFileId(artifactDetails.linkUrl);
      if (!sourceId) {
        throw new Error("Could not read that Google Drive link. Please paste a valid Google Doc, Sheet, Slides, or Drive file URL.");
      }
      try {
        const sourceFile = DriveApp.getFileById(sourceId);
        const copyName = artifactDetails.title || sourceFile.getName();
        const copiedFile = sourceFile.makeCopy(copyName, yearFolder);
        setArtifactSharing(copiedFile);
        fileUrl = copiedFile.getUrl();
        fileId = copiedFile.getId();
      } catch (copyErr) {
        Logger.log("Link copy failed for source " + sourceId + ": " + copyErr);
        throw new Error(
          "We couldn't access your linked document. Please open it, click Share, " +
          "and set 'Anyone with the link' to Viewer (or share it with your teacher), " +
          "then try submitting again. You can also use 'Upload Local File' instead."
        );
      }
    }

    // Save to Artifacts Sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ARTIFACTS);
    // Prefix with 'A' so Sheets always stores the ID as text. Pure numeric
    // strings get auto-parsed as numbers, which silently loses precision past
    // ~15 digits and can produce floats like 1.77729e+14 in the cell.
    const id = 'A' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
    const now = new Date().toISOString().replace('T', ' ').substring(0, 19);
    
    sheet.appendRow([
      id,
      artifactDetails.studentName || '',
      artifactDetails.studentEmail, // Will map perfectly to active session identity
      artifactDetails.gradYear || '',
      artifactDetails.schoolYear || '',
      artifactDetails.gradeLevel || '',
      artifactDetails.title,
      artifactDetails.type || '',
      artifactDetails.assignmentName || '',
      fileUrl, 
      fileId, 
      artifactDetails.bestWork || false,
      now, 
      now, 
      artifactDetails.reflection || '',
      '', 
      '', 
      '' 
    ]);

    const auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
    if (auditSheet) {
      auditSheet.appendRow([now, artifactDetails.studentEmail, 'Upload Artifact', id, '']);
    }

    // --- FEATURE: Send Email Notification to Selected Teacher ---
    if (artifactDetails.notifyTeacherEmail) {
      try {
        // Escape every interpolated value so a malicious roster name, title,
        // or reflection can't smuggle HTML/JS into the teacher's inbox.
        const safeStudentName = escapeHtmlServer(artifactDetails.studentName);
        const safeGradYear = escapeHtmlServer(artifactDetails.gradYear);
        const safeAssignment = escapeHtmlServer(artifactDetails.assignmentName || 'N/A');
        const safeTitle = escapeHtmlServer(artifactDetails.title);
        const safeType = escapeHtmlServer(artifactDetails.type || 'N/A');
        const safeFileUrl = encodeURI(fileUrl || '');

        const subject = `New Portfolio Submission: ${artifactDetails.studentName}`;
        const htmlBody = `
          <div style="font-family: sans-serif; color: #333; max-width: 600px;">
            <h2 style="color: #5b2c83; border-bottom: 2px solid #d4a200; padding-bottom: 5px;">New Portfolio Artifact Submitted</h2>
            <p><strong>Student:</strong> ${safeStudentName} (Class of ${safeGradYear})</p>
            <p><strong>Assignment:</strong> ${safeAssignment}</p>
            <p><strong>Title:</strong> ${safeTitle}</p>
            <p><strong>Type:</strong> ${safeType}</p>
            ${artifactDetails.bestWork ? '<p><strong style="color:#d4a200;">★ Marked as Best Work</strong></p>' : ''}
            <br>
            <a href="${safeFileUrl}" style="background-color: #5b2c83; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">View Artifact Document</a>
            <br><br>
            <p style="font-size: 12px; color: #666; border-top: 1px solid #eee; padding-top: 10px;">
              You can log into the Roy-Hart Portfolio Portal to view their full history and leave official feedback.
            </p>
          </div>
        `;
        MailApp.sendEmail({
          to: artifactDetails.notifyTeacherEmail,
          subject: subject,
          htmlBody: htmlBody,
          replyTo: artifactDetails.studentEmail, // Forces replies to go to the student
          name: `Portfolio Portal (${artifactDetails.studentName})` // Masks your name in the inbox
        });
      } catch (err) {
        Logger.log("Failed to send notification email: " + err);
      }
    }

    return { status: 'success', message: 'Artifact saved successfully!' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

/**
 * Toggles the "Best Work" status of an artifact.
 * Students can toggle their own, and Teachers can toggle any.
 */
function toggleBestWork(artifactId, isBest) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const isUserTeacher = isTeacher(userEmail);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ARTIFACTS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === artifactId.toString()) {
        
        // Security: Must be the owning student OR a teacher
        if (data[i][2].toString().toLowerCase() !== userEmail.toLowerCase() && !isUserTeacher) {
          throw new Error("Unauthorized: You can only edit your own artifacts.");
        }
        
        // Column L is index 11 (0-based) -> index 12 (1-based getRange)
        sheet.getRange(i + 1, 12).setValue(isBest);
        
        // Update Last Updated Timestamp
        const now = new Date().toISOString().replace('T', ' ').substring(0, 19);
        sheet.getRange(i + 1, 14).setValue(now);
        
        const auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
        if (auditSheet) {
          auditSheet.appendRow([now, userEmail, `Toggle Best Work to ${isBest}`, artifactId, '']);
        }
        
        return { status: 'success' };
      }
    }
    throw new Error('Artifact not found.');
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function saveTeacherNote(artifactId, note) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    
    // SECURITY PATCH: Feedback Eraser Protection
    if (!isTeacher(userEmail)) {
      throw new Error("Unauthorized access attempt logged. Only teachers may edit feedback.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ARTIFACTS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === artifactId.toString()) {
        sheet.getRange(i + 1, 16).setValue(note);
        
        const now = new Date().toISOString().replace('T', ' ').substring(0, 19);
        sheet.getRange(i + 1, 14).setValue(now);
        
        const auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
        if (auditSheet) {
          auditSheet.appendRow([now, userEmail, 'Update Teacher Note', artifactId, '']);
        }
        
        return { status: 'success' };
      }
    }
    throw new Error('Artifact not found.');
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function bulkImportStudents(csvText) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!isTeacher(userEmail)) {
      throw new Error("Unauthorized access attempt logged.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
    if (!sheet) throw new Error("Students sheet not found.");

    const rows = Utilities.parseCsv(csvText);
    if (rows.length < 2) throw new Error("The CSV data must contain headers and at least one student row.");

    const headers = rows[0].map(h => h.toString().trim().toLowerCase());
    const nameIdx = headers.indexOf('student name');
    const emailIdx = headers.indexOf('student email');
    const gradYearIdx = headers.indexOf('grad year');
    const gradeIdx = headers.indexOf('current grade');

    if (nameIdx === -1 || emailIdx === -1) {
      throw new Error("Your CSV must contain exact headers: 'Student Name' and 'Student Email'.");
    }

    const existingData = sheet.getDataRange().getValues();
    const existingEmails = new Set();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][1]) {
        existingEmails.add(existingData[i][1].toString().trim().toLowerCase());
      }
    }

    let addedCount = 0;
    let skippedCount = 0;
    const newRows = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0 || !row[emailIdx]) continue;

      const email = row[emailIdx].toString().trim();
      const name = row[nameIdx] ? row[nameIdx].toString().trim() : '';
      const gradYear = gradYearIdx !== -1 && row[gradYearIdx] ? row[gradYearIdx].toString().trim() : '';
      const currentGrade = gradeIdx !== -1 && row[gradeIdx] ? row[gradeIdx].toString().trim() : '';

      if (!email || existingEmails.has(email.toLowerCase())) {
        if (email) skippedCount++;
        continue;
      }

      newRows.push([name, email, gradYear, currentGrade, '', '', true]);
      existingEmails.add(email.toLowerCase());
      addedCount++;
    }

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    return { 
      status: 'success', 
      added: addedCount, 
      skipped: skippedCount, 
      message: `Successfully added ${addedCount} student(s). Skipped ${skippedCount} duplicate(s).` 
    };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function deleteArtifact(artifactId) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    
    // SECURITY PATCH: Only teachers can delete artifacts
    if (!isTeacher(userEmail)) {
      throw new Error("Unauthorized access attempt logged. Only teachers may delete artifacts.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ARTIFACTS);
    const data = sheet.getDataRange().getValues();
    
    // Loop to find the artifact by ID
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (data[i][0].toString() === artifactId.toString()) {
        // Capture the Drive file ID (column K, index 10) BEFORE deleting the row
        // so we can also trash the underlying file and avoid orphaned storage.
        const driveFileId = data[i][10];
        let trashNote = '';
        
        if (driveFileId) {
          try {
            DriveApp.getFileById(driveFileId.toString()).setTrashed(true);
            trashNote = 'Drive file trashed: ' + driveFileId;
          } catch (fileErr) {
            // File may already be gone, or permissions may have changed.
            // We still proceed with the row deletion.
            Logger.log("Could not trash Drive file " + driveFileId + ": " + fileErr);
            trashNote = 'Drive file could not be trashed: ' + driveFileId;
          }
        } else {
          trashNote = 'No Drive file ID stored for this artifact';
        }
        
        sheet.deleteRow(i + 1); // +1 because array is 0-indexed but sheet rows are 1-indexed
        
        // Log the deletion
        const auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
        if (auditSheet) {
          const now = new Date().toISOString().replace('T', ' ').substring(0, 19);
          auditSheet.appendRow([now, userEmail, 'Delete Artifact', artifactId, trashNote]);
        }
        
        return { status: 'success', message: 'Artifact removed successfully' };
      }
    }
    throw new Error('Artifact not found.');
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

/**
 * ==============================================================
 * PORTFOLIO COMPILER (chunked)
 *
 * The original implementation did everything in one server call:
 * fetch every artifact, convert each to PDF, base64-encode them all,
 * and return a giant array. That's at risk of two limits:
 *   1. Apps Script's 6-minute execution cap (a 30+ artifact senior
 *      portfolio with mixed Docs/Slides can take longer than that).
 *   2. The ~50MB google.script.run response payload cap.
 *
 * The new approach splits the work into many small calls:
 *   - getPortfolioManifest(email) returns student info, the cover-page
 *     PDF, and a list of artifact descriptors (with stable indexes).
 *   - getPortfolioArtifactPdf(email, idx) returns a single base64 PDF
 *     (or a placeholder for un-convertible items).
 *
 * The client loops the per-artifact call, can show real progress, and
 * each individual call comfortably fits inside both Apps Script limits.
 * ==============================================================
 */

/**
 * Internal: pulls the artifact rows for a student in the same order
 * the export should use them. Used by both manifest and per-artifact calls.
 */
function _portfolioArtifactsForStudent(studentEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ARTIFACTS);
  if (!sheet) throw new Error("Artifacts sheet not found.");

  const data = sheet.getDataRange().getValues();
  const artifacts = [];
  let studentName = "Student";
  let gradYear = "Unknown";

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().toLowerCase() === studentEmail.toLowerCase()) {
      studentName = data[i][1] || studentName;
      gradYear = data[i][3] || gradYear;
      artifacts.push({
        title: data[i][6],
        type: data[i][7],
        assignmentName: data[i][8],
        url: data[i][9],
        driveFileId: data[i][10],
        bestWork: String(data[i][11]).toLowerCase() === 'true',
        schoolYear: data[i][4],
        gradeLevel: data[i][5],
        reflection: data[i][14] || '',
        teacherNotes: data[i][15] || ''
      });
    }
  }

  // Group by school year for a more readable cover page and PDF order.
  // Newest year first so a senior's most recent work leads the document.
  artifacts.sort((a, b) => {
    const ay = (a.schoolYear || '').toString();
    const by = (b.schoolYear || '').toString();
    if (ay !== by) return by.localeCompare(ay);
    return 0;
  });

  return { studentName, gradYear, artifacts };
}

/**
 * Returns metadata + the cover-page PDF for the export. The client uses
 * the artifact list to drive subsequent per-artifact fetches.
 */
function getPortfolioManifest(studentEmail) {
  const activeEmail = Session.getActiveUser().getEmail();
  if (activeEmail.toLowerCase() !== studentEmail.toLowerCase() && !isTeacher(activeEmail)) {
    throw new Error("Unauthorized access attempt.");
  }

  const { studentName, gradYear, artifacts } = _portfolioArtifactsForStudent(studentEmail);
  if (artifacts.length === 0) {
    throw new Error("No artifacts found to export.");
  }

  // Build the cover/TOC page.
  let summaryHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Helvetica', sans-serif; padding: 40px; color: #333; line-height: 1.5; }
        h1 { color: #5b2c83; text-align: center; font-size: 28px; text-transform: uppercase; border-bottom: 2px solid #d4a200; padding-bottom: 10px; margin-bottom: 5px; }
        .subtitle { text-align: center; color: #666; font-size: 14px; margin-bottom: 40px; }
        .year-header { color: #5b2c83; font-size: 18px; margin-top: 30px; margin-bottom: 10px; padding-bottom: 4px; border-bottom: 1px solid #d4a200; text-transform: uppercase; letter-spacing: 1px; }
        .artifact { margin-bottom: 18px; padding: 12px 15px; border-left: 4px solid #5b2c83; background: #f9f9f9; }
        .art-title { margin: 0 0 5px 0; color: #452064; font-size: 16px; }
        .meta { font-size: 12px; color: #555; margin-bottom: 8px; }
        .box { margin-top: 8px; padding: 8px 10px; font-size: 12px; border-radius: 4px; }
        .reflection { background: #eef2ff; border-left: 3px solid #6366f1; }
        .feedback { background: #fffbeb; border-left: 3px solid #d4a200; }
        .badge { background: #d4a200; color: white; padding: 3px 8px; border-radius: 12px; font-size: 10px; font-weight: bold; text-transform: uppercase; vertical-align: middle; margin-left: 10px; }
      </style>
    </head>
    <body>
      <h1>${escapeHtmlServer(studentName)}</h1>
      <div class="subtitle">Writing Portfolio &bull; Class of ${escapeHtmlServer(gradYear)} &bull; Exported ${new Date().toLocaleDateString()}</div>
      <p style="font-size: 14px; margin-bottom: 30px;">
        This document serves as the cover letter and table of contents for the student's writing portfolio.
        The original converted documents for each artifact follow this summary.
      </p>
  `;

  let lastYear = null;
  artifacts.forEach(art => {
    if (art.schoolYear !== lastYear) {
      summaryHtml += `<div class="year-header">${escapeHtmlServer(art.schoolYear || 'Unknown Year')}</div>`;
      lastYear = art.schoolYear;
    }
    summaryHtml += `
      <div class="artifact">
        <h2 class="art-title">${escapeHtmlServer(art.title)} ${art.bestWork ? '<span class="badge">★ Best Work</span>' : ''}</h2>
        <div class="meta">
          <strong>Assignment:</strong> ${escapeHtmlServer(art.assignmentName)} | <strong>Type:</strong> ${escapeHtmlServer(art.type)} | <strong>Grade:</strong> ${escapeHtmlServer(art.gradeLevel)}
        </div>
        ${art.reflection ? `<div class="box reflection"><strong>Student Reflection:</strong><br>${escapeHtmlServer(art.reflection)}</div>` : ''}
        ${art.teacherNotes ? `<div class="box feedback"><strong>Teacher Feedback:</strong><br>${escapeHtmlServer(art.teacherNotes)}</div>` : ''}
      </div>
    `;
  });

  summaryHtml += `</body></html>`;
  const coverBlob = HtmlService.createHtmlOutput(summaryHtml).getAs('application/pdf');

  // Lightweight descriptors for the client to drive per-artifact fetches.
  const descriptors = artifacts.map((a, idx) => ({
    index: idx,
    title: a.title,
    schoolYear: a.schoolYear
  }));

  return {
    studentName: studentName,
    gradYear: gradYear,
    coverPdf: Utilities.base64Encode(coverBlob.getBytes()),
    artifacts: descriptors
  };
}

/**
 * Returns the base64 PDF for a single artifact, or a base64 placeholder PDF
 * if the source can't be converted. Each call is lightweight and fits well
 * inside Apps Script execution + payload limits.
 */
function getPortfolioArtifactPdf(studentEmail, artifactIndex) {
  const activeEmail = Session.getActiveUser().getEmail();
  if (activeEmail.toLowerCase() !== studentEmail.toLowerCase() && !isTeacher(activeEmail)) {
    throw new Error("Unauthorized access attempt.");
  }

  const { artifacts } = _portfolioArtifactsForStudent(studentEmail);
  if (artifactIndex < 0 || artifactIndex >= artifacts.length) {
    throw new Error("Artifact index out of range.");
  }
  const art = artifacts[artifactIndex];

  try {
    let fileId = art.driveFileId;
    if (!fileId && art.url) {
      const match = art.url.toString().match(/[-\w]{25,}/);
      if (match) fileId = match[0];
    }

    if (fileId) {
      const file = DriveApp.getFileById(fileId.toString());
      const pdfBlob = file.getAs('application/pdf');
      return Utilities.base64Encode(pdfBlob.getBytes());
    }
    throw new Error("External link or missing Drive ID.");
  } catch (e) {
    Logger.log("Could not natively convert file for " + art.title + ": " + e.toString());
    const placeholderHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: 'Helvetica', sans-serif; padding: 60px; color: #333; text-align: center; }
          .container { margin-top: 150px; padding: 40px; border: 2px dashed #d4a200; border-radius: 15px; background-color: #f9f9f9; }
          h2 { color: #5b2c83; font-size: 24px; margin-bottom: 15px; text-transform: uppercase; letter-spacing: 1px;}
          p { font-size: 16px; color: #555; margin-bottom: 25px; line-height: 1.6; }
          .url-box { font-family: monospace; background: #fff; padding: 15px; border: 1px solid #ddd; word-break: break-all; color: #5b2c83; font-size: 14px;}
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Digital Artifact Placeholder</h2>
          <h3>${escapeHtmlServer(art.title)}</h3>
          <p>This artifact could not be automatically converted into a printed document. It may be a video, a collaborative website, a Canva presentation, or an external link.</p>
          <p><strong>To review this artifact, please access the original digital version here:</strong></p>
          <div class="url-box">${escapeHtmlServer(art.url || 'No URL provided')}</div>
        </div>
      </body>
      </html>
    `;
    const placeholderBlob = HtmlService.createHtmlOutput(placeholderHtml).getAs('application/pdf');
    return Utilities.base64Encode(placeholderBlob.getBytes());
  }
}
