// ===============================================================
// FILE: Code.gs
// ===============================================================

const SPREADSHEET_ID = '15r5QIYPTfPem1qJEltcZw1mu2LO_N4N9pFAPkQDR5sg'; 
const LESSONS_SHEET_NAME = 'Lessons';
const SECTIONS_SHEET_NAME = 'Sections';
const PROGRESS_SHEET_NAME = 'Progress';
const QUIZZES_SHEET_NAME = 'Quizzes';
const ADMINS_SHEET_NAME = 'Admins';
const TIMESTAMPS_SHEET_NAME = 'VideoTimestamps';
const QUIZ_SCORES_SHEET_NAME = 'QuizScores';
const USER_PROFILES_SHEET_NAME = 'UserProfiles';

// ----------------------------------------------------
// ระบบ ROUTING (แสดงผลหน้าจอหลัก)
// ----------------------------------------------------
function doGet(e){
  const userEmail = Session.getActiveUser().getEmail().trim().toLowerCase();
  const permissions = getCurrentUserPermissions(userEmail);
  const page = e.parameter.page;
  
  let userProfile = { Email: userEmail, FullName: '', Nickname: '', Department: 'General', Position: '', ProfileImage: '' };
  const profiles = getSheetData(USER_PROFILES_SHEET_NAME);
  if (profiles && profiles.length > 0) {
     const found = profiles.find(p => String(p.Email).trim().toLowerCase() === userEmail);
     if (found) userProfile = found;
  }

  const templateData = {
    permissions: permissions, 
    baseUrl: ScriptApp.getService().getUrl(),
    userEmail: userEmail,
    userProfile: userProfile
  };

  if(page){
    switch(page){
      case 'superadmin':
        if(!permissions.isSuperAdmin) return HtmlService.createHtmlOutput('Access Denied.');
        return renderPage('SuperAdmin_View','Super Admin',templateData);
      case 'admin':
        if(!permissions.canEditLessons && !permissions.canEditQuizzes) return HtmlService.createHtmlOutput('Access Denied.');
        return renderPage('Admin_View','Admin Dashboard',templateData);
      case 'profile': 
        templateData.studentData = getStudentData(userEmail);
        return renderPage('Profile_View', 'My Profile', templateData);
      default:
        templateData.studentData = getStudentData(userEmail);
        return renderPage('Student_View','E-Learning Platform',templateData);
    }
  }

  templateData.studentData = getStudentData(userEmail);
  return renderPage('Student_View','E-Learning Platform',templateData);
}

function renderPage(templateFile, title, data){
  const navTemplate = HtmlService.createTemplateFromFile('Navigation_Menu');
  navTemplate.baseUrl = data.baseUrl;
  navTemplate.permissions = data.permissions;
  navTemplate.userEmail = data.userEmail;
  navTemplate.userProfile = data.userProfile; 
  
  const navHtml = navTemplate.evaluate().getContent();
  const mainTemplate = HtmlService.createTemplateFromFile(templateFile);
  mainTemplate.navigationMenu = navHtml;
  
  Object.keys(data).forEach(key => {
    mainTemplate[key] = data[key];
  });
  
  return mainTemplate.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ----------------------------------------------------
// ระบบ HELPER (ฟังก์ชันช่วยเหลือต่างๆ)
// ----------------------------------------------------
function getSheetData(sheetName) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }
    const rows = sheet.getDataRange().getValues();
    const headers = rows.shift();
    return rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
  } catch (e) {
    console.error(`Error getting data from sheet "${sheetName}": ${e.toString()}`);
    return [];
  }
}

function getCurrentUserPermissions(email) {
  const safeEmail = String(email).trim().toLowerCase();
  const admins = getSheetData(ADMINS_SHEET_NAME);
  const userAdmin = admins.find(admin => String(admin.Email).trim().toLowerCase() === safeEmail);
  if (userAdmin) {
    return {
      canEditLessons: userAdmin.CanEditLessons === true,
      canEditQuizzes: userAdmin.CanEditQuizzes === true,
      isSuperAdmin: userAdmin.IsSuperAdmin === true
    };
  }
  return { canEditLessons: false, canEditQuizzes: false, isSuperAdmin: false };
}

// ----------------------------------------------------
// ระบบฝั่ง STUDENT (ผู้เรียน)
// ----------------------------------------------------
function updateLatestTimestamp(data) {
  const { sectionId, timestamp } = data;
  const userEmail = Session.getActiveUser().getEmail().trim().toLowerCase();
  if (!userEmail || !sectionId) return;
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(TIMESTAMPS_SHEET_NAME);
    const sheetData = sheet.getDataRange().getValues();
    const headers = sheetData[0];
    const emailCol = headers.indexOf('Email');
    const sectionCol = headers.indexOf('SectionID');
    const timestampCol = headers.indexOf('LatestTimestamp');
    for (let i = 1; i < sheetData.length; i++) {
      if (String(sheetData[i][emailCol]).trim().toLowerCase() === userEmail && sheetData[i][sectionCol] === sectionId) {
        sheet.getRange(i + 1, timestampCol + 1).setValue(timestamp);
        return { status: 'updated' };
      }
    }
    sheet.appendRow([Session.getActiveUser().getEmail(), sectionId, timestamp]);
    return { status: 'created' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function getStudentData(userEmail) {
  userEmail = (userEmail || Session.getActiveUser().getEmail() || 'anonymous').trim().toLowerCase();
  const lessons = getSheetData(LESSONS_SHEET_NAME).filter(l => l.IsActive === true);
  const sections = getSheetData(SECTIONS_SHEET_NAME);
  
  const progress = getSheetData(PROGRESS_SHEET_NAME)
      .filter(p => String(p.Email).trim().toLowerCase() === userEmail)
      .map(p => p.SectionID);
  
  const timestamps = getSheetData(TIMESTAMPS_SHEET_NAME)
      .filter(t => String(t.Email).trim().toLowerCase() === userEmail);
      
  const scores = getSheetData(QUIZ_SCORES_SHEET_NAME)
      .filter(s => String(s.Email).trim().toLowerCase() === userEmail);
      
  const allQuizzes = getSheetData(QUIZZES_SHEET_NAME);

  let userProfile = { Email: userEmail, FullName: '', Nickname: '', Department: 'General', Position: 'พนักงานทั่วไป', ProfileImage: '' };
  const profiles = getSheetData(USER_PROFILES_SHEET_NAME);
  if (profiles && profiles.length > 0) {
     const found = profiles.find(p => String(p.Email).trim().toLowerCase() === userEmail);
     if (found) userProfile = found;
  }

  let totalBN = 0;

  lessons.forEach(lesson => {
    lesson.sections = sections
      .filter(s => s.LessonID === lesson.LessonID)
      .sort((a, b) => a.Order - b.Order)
      .map(section => {
        const timestampData = timestamps.find(t => t.SectionID === section.SectionID);
        section.latestTimestamp = timestampData ? timestampData.LatestTimestamp : 0;
        section.hasQuiz = allQuizzes.some(q => q.SectionID === section.SectionID);
        const latestScore = scores.filter(s => s.SectionID === section.SectionID)
                                  .sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp))[0];
        section.quizScore = latestScore || null;
        
        // JAVIS UPDATE: จัดการ ContentType (ถ้าไม่มีให้มองว่าเป็น youtube)
        section.ContentType = section.ContentType || 'youtube';

        if (progress.includes(section.SectionID)) {
            let vp = parseInt(section.VideoPoints);
            totalBN += isNaN(vp) ? 10 : vp;
        }
        if (section.quizScore && (section.quizScore.Score / section.quizScore.TotalQuestions) >= 0.5) {
            let qp = parseInt(section.QuizPoints);
            totalBN += isNaN(qp) ? 50 : qp;
        }

        return section;
      });
    
    lesson.isCompleted = lesson.sections.length > 0 && lesson.sections.every(s => progress.includes(s.SectionID));
  });

  const userDept = userProfile.Department ? userProfile.Department.toString().trim() : '';
  const userPos = userProfile.Position ? userProfile.Position.toString().trim() : '';

  const accessibleLessons = lessons.filter(lesson => {
      if (lesson.TargetEmails && lesson.TargetEmails.trim() !== '') {
          const targetEmails = lesson.TargetEmails.split(',').map(e => e.trim().toLowerCase());
          return targetEmails.includes(userEmail);
      }
      let deptMatch = true; 
      if (lesson.TargetDepartments) {
          const targetDepts = lesson.TargetDepartments.split(',').map(d => d.trim());
          deptMatch = targetDepts.includes('ทั้งหมด (All)') || targetDepts.includes(userDept);
      }
      let posMatch = true; 
      if (lesson.TargetPositions) {
          const targetPos = lesson.TargetPositions.split(',').map(p => p.trim());
          posMatch = targetPos.includes('ทั้งหมด (All)') || targetPos.includes(userPos);
      }
      return deptMatch && posMatch;
  });

  const personalizedLessons = accessibleLessons.filter(lesson => {
      if (lesson.TargetEmails && lesson.TargetEmails.trim() !== '') return true;
      if (lesson.TargetDepartments) {
          const targetDepts = lesson.TargetDepartments.split(',').map(d => d.trim());
          return targetDepts.includes(userDept) && !targetDepts.includes('ทั้งหมด (All)');
      }
      return false;
  });

  const personalizedIds = personalizedLessons.map(l => l.LessonID);
  const featuredLessons = accessibleLessons.filter(lesson => !personalizedIds.includes(lesson.LessonID));

  return {
    userInfo: { email: userEmail, progress: progress, profile: userProfile, totalBN: totalBN },
    lessons: lessons,
    featuredLessons: featuredLessons,
    personalizedLessons: personalizedLessons
  };
}

function getPublicLeaderboard() {
  const usersProgress = getSheetData(PROGRESS_SHEET_NAME);
  const scores = getSheetData(QUIZ_SCORES_SHEET_NAME);
  const sections = getSheetData(SECTIONS_SHEET_NAME);
  const profiles = getSheetData(USER_PROFILES_SHEET_NAME);

  const sectionMap = {};
  sections.forEach(s => {
     sectionMap[s.SectionID] = {
         vp: parseInt(s.VideoPoints) || 10,
         qp: parseInt(s.QuizPoints) || 50
     };
  });

  const studentMap = {};

  profiles.forEach(p => {
      const email = String(p.Email).trim().toLowerCase();
      let defaultImg = 'https://ui-avatars.com/api/?name=' + encodeURIComponent(p.Nickname || p.FullName || email) + '&background=31c1d7&color=fff';
      studentMap[email] = {
          email: email,
          name: p.Nickname || p.FullName || email.split('@')[0],
          dept: p.Department || 'General',
          img: p.ProfileImage || defaultImg,
          completedSections: [],
          scores: {},
          totalBN: 0
      };
  });

  usersProgress.forEach(p => {
      const email = String(p.Email).trim().toLowerCase();
      if(!studentMap[email]) { 
          studentMap[email] = { email: email, name: email.split('@')[0], dept: 'General', img: 'https://ui-avatars.com/api/?name='+email+'&background=31c1d7&color=fff', completedSections: [], scores: {}, totalBN: 0 };
      }
      if (sectionMap[p.SectionID] && !studentMap[email].completedSections.includes(p.SectionID)) {
          studentMap[email].completedSections.push(p.SectionID);
          studentMap[email].totalBN += sectionMap[p.SectionID].vp; 
      }
  });

  scores.forEach(s => {
      const email = String(s.Email).trim().toLowerCase();
      if(!studentMap[email]) {
          studentMap[email] = { email: email, name: email.split('@')[0], dept: 'General', img: 'https://ui-avatars.com/api/?name='+email+'&background=31c1d7&color=fff', completedSections: [], scores: {}, totalBN: 0 };
      }
      if (sectionMap[s.SectionID]) {
          const percent = s.TotalQuestions > 0 ? (s.Score / s.TotalQuestions) : 0;
          if (percent >= 0.5) { 
             if (!studentMap[email].scores[s.SectionID]) { 
                 studentMap[email].scores[s.SectionID] = percent;
                 studentMap[email].totalBN += sectionMap[s.SectionID].qp; 
             }
          }
      }
  });

  const activeStudents = Object.values(studentMap).filter(std => std.totalBN > 0);
  return activeStudents.sort((a,b) => b.totalBN - a.totalBN);
}

function saveQuizScore(scoreData) {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return { status: 'error', message: 'User not logged in.' };

    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(QUIZ_SCORES_SHEET_NAME);
        sheet.appendRow([
            userEmail,
            scoreData.sectionId,
            scoreData.score,
            scoreData.totalQuestions,
            new Date()
        ]);
        
        const passed = (scoreData.score / scoreData.totalQuestions) >= 0.5;
        if (passed) {
            recordProgress(scoreData.sectionId);
        }

        return { status: 'success', passed: passed };
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}

function getQuiz(sectionId) {
    try {
        const questions = getSheetData(QUIZZES_SHEET_NAME)
            .filter(q => q.SectionID === sectionId)
            .map(q => ({
                question: q.QuestionText,
                options: JSON.parse(q.Options),
                answer: q.CorrectAnswer
            }));
        return questions;
    } catch (e) {
        return { error: e.toString() };
    }
}

function recordProgress(sectionId) {
  const userEmail = Session.getActiveUser().getEmail();
  const safeEmail = userEmail.trim().toLowerCase();
  if (!userEmail) return { status: 'error', message: 'User not logged in.' };
  
  const progressSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PROGRESS_SHEET_NAME);
  const data = progressSheet.getDataRange().getValues();
  const alreadyExists = data.some(row => String(row[0]).trim().toLowerCase() === safeEmail && row[1] === sectionId);
  
  if (!alreadyExists) {
    progressSheet.appendRow([userEmail, sectionId, new Date()]);
    return { status: 'success' };
  }
  return { status: 'info', message: 'Already recorded.' };
}

// JAVIS ADDED: ฟังก์ชันอัปโหลดหลักฐานเรียนภายนอก และให้คะแนน
function submitExternalProof(data) {
  const userEmail = Session.getActiveUser().getEmail();
  const safeEmail = userEmail.trim().toLowerCase();
  if (!userEmail) return { status: 'error', message: 'User not logged in' };

  try {
    let fileUrl = '';
    // เซฟรูปภาพลง Google Drive ถ้ามีการแนบไฟล์มา
    if (data.imageData) {
      const folderId = "1uYLnZNVG7-Plmf2K_HO0qijLq4l4qwYQ"; // ใช้โฟลเดอร์เดียวกับโปรไฟล์
      const folder = DriveApp.getFolderById(folderId);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const blob = Utilities.newBlob(Utilities.base64Decode(data.imageData), data.imageMimeType, userEmail + "_proof_" + data.sectionId);
      const file = folder.createFile(blob);
      fileUrl = file.getUrl();
    }
    
    // อัปเดตความคืบหน้า (เหมือนดูวิดีโอจบ)
    recordProgress(data.sectionId);
    
    return { status: 'success', fileUrl: fileUrl };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function updateUserProfile(data) {
  const userEmail = Session.getActiveUser().getEmail();
  const safeEmail = userEmail.trim().toLowerCase();
  if (!userEmail) return { status: 'error', message: 'User not logged in' };

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('UserProfiles');
    if (!sheet) {
      return { status: 'error', message: 'หาชีตชื่อ UserProfiles ไม่เจอจริงๆ ครับ โปรดเช็ค SPREADSHEET_ID' };
    }
    const sheetData = sheet.getDataRange().getValues();
    const headers = sheetData[0];
    
    const emailCol = headers.indexOf('Email');
    const nameCol = headers.indexOf('FullName');
    const nickCol = headers.indexOf('Nickname');
    const deptCol = headers.indexOf('Department');
    const posCol = headers.indexOf('Position');
    const imgCol = headers.indexOf('ProfileImage');
    
    let profileImageUrl = null;

    if (data.imageData) {
      const folderId = "1uYLnZNVG7-Plmf2K_HO0qijLq4l4qwYQ"; 
      const folder = DriveApp.getFolderById(folderId);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const blob = Utilities.newBlob(Utilities.base64Decode(data.imageData), data.imageMimeType, userEmail + "_profile");
      const file = folder.createFile(blob);
      profileImageUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w500";
    }

    let rowIndex = -1;
    for (let i = 1; i < sheetData.length; i++) {
      if (String(sheetData[i][emailCol]).trim().toLowerCase() === safeEmail) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, nameCol + 1).setValue(data.fullName);
      sheet.getRange(rowIndex, nickCol + 1).setValue(data.nickname);
      if (profileImageUrl) {
        sheet.getRange(rowIndex, imgCol + 1).setValue(profileImageUrl);
      }
    } else {
      const newRow = new Array(headers.length).fill('');
      newRow[emailCol] = userEmail;
      newRow[nameCol] = data.fullName;
      newRow[nickCol] = data.nickname;
      newRow[deptCol] = 'General';
      newRow[posCol] = 'พนักงาน';
      if (profileImageUrl) newRow[imgCol] = profileImageUrl;
      sheet.appendRow(newRow);
    }
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// ----------------------------------------------------
// ระบบ ADMIN - บันทึกและลบข้อมูลเนื้อหา
// ----------------------------------------------------
function getAdminData() {
    const permissions = getCurrentUserPermissions(Session.getActiveUser().getEmail());
    if (!permissions.canEditLessons && !permissions.canEditQuizzes) return { error: 'Permission Denied' };
    
    const lessons = getSheetData(LESSONS_SHEET_NAME);
    const sections = getSheetData(SECTIONS_SHEET_NAME);
    const quizzes = getSheetData(QUIZZES_SHEET_NAME);

    lessons.forEach(lesson => {
        lesson.sections = sections.filter(s => s.LessonID === lesson.LessonID).sort((a, b) => a.Order - b.Order);
        lesson.sections.forEach(s => {
            s.questionCount = quizzes.filter(q => q.SectionID === s.SectionID).length;
            s.ContentType = s.ContentType || 'youtube'; // Default 
        });
    });

    let settings = { categories: [], departments: [], positions: [] };
    try {
        const settingsData = getSheetData('Settings');
        if (settingsData && settingsData.length > 0) {
            settings.categories = settingsData.map(r => r.Categories).filter(Boolean);
            settings.departments = settingsData.map(r => r.Departments).filter(Boolean);
            settings.positions = settingsData.map(r => r.Positions).filter(Boolean);
        }
    } catch(e) {}

    let users = [];
    try {
        users = getSheetData(USER_PROFILES_SHEET_NAME);
    } catch(e) {}

    return {
      lessons: lessons,
      allQuizzes: quizzes,
      sections: sections,
      settings: settings,
      users: users
    };
}

function saveLesson(data) {
  try {
    const sheetName = typeof LESSONS_SHEET_NAME !== 'undefined' ? LESSONS_SHEET_NAME : 'Lessons';
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    
    if (!sheet) {
      return { status: 'error', message: 'หาแผ่นงาน (Sheet) บทเรียนไม่พบ' };
    }

    const sheetData = sheet.getDataRange().getValues();
    const headers = sheetData[0];
    
    const idCol = headers.indexOf('LessonID');
    const titleCol = headers.indexOf('Title');
    const descCol = headers.indexOf('Description');
    const activeCol = headers.indexOf('IsActive');
    const categoryCol = headers.indexOf('Category');
    const targetDeptCol = headers.indexOf('TargetDepartments'); 
    const targetPosCol = headers.indexOf('TargetPositions'); 
    const targetEmailCol = headers.indexOf('TargetEmails'); 
    
    if (idCol === -1 || titleCol === -1) {
      return { status: 'error', message: 'โครงสร้างคอลัมน์ใน Sheet ไม่ถูกต้อง' };
    }

    if (data.LessonID) {
      let rowIndex = -1;
      for (let i = 1; i < sheetData.length; i++) {
        if (sheetData[i][idCol] == data.LessonID) { 
          rowIndex = i + 1; 
          break; 
        }
      }
      
      if (rowIndex > -1) {
        sheet.getRange(rowIndex, titleCol + 1).setValue(data.Title);
        if (descCol > -1) sheet.getRange(rowIndex, descCol + 1).setValue(data.Description);
        if (activeCol > -1) sheet.getRange(rowIndex, activeCol + 1).setValue(data.IsActive);
        if (categoryCol > -1) sheet.getRange(rowIndex, categoryCol + 1).setValue(data.Category);
        if (targetDeptCol > -1) sheet.getRange(rowIndex, targetDeptCol + 1).setValue(data.TargetDepartments);
        if (targetPosCol > -1) sheet.getRange(rowIndex, targetPosCol + 1).setValue(data.TargetPositions);
        if (targetEmailCol > -1) sheet.getRange(rowIndex, targetEmailCol + 1).setValue(data.TargetEmails);
      } else {
        return { status: 'error', message: 'ไม่พบคอร์สเรียนนี้ในฐานข้อมูล' };
      }
      return { status: 'success' };
    } 
    else {
      const newLessonId = 'LSN-' + Utilities.getUuid().substring(0, 8).toUpperCase();
      const newRow = new Array(headers.length).fill('');
      
      newRow[idCol] = newLessonId;
      newRow[titleCol] = data.Title;
      if (descCol > -1) newRow[descCol] = data.Description;
      if (activeCol > -1) newRow[activeCol] = data.IsActive;
      if (categoryCol > -1) newRow[categoryCol] = data.Category;
      if (targetDeptCol > -1) newRow[targetDeptCol] = data.TargetDepartments;
      if (targetPosCol > -1) newRow[targetPosCol] = data.TargetPositions;
      if (targetEmailCol > -1) newRow[targetEmailCol] = data.TargetEmails;
      
      sheet.appendRow(newRow);
      return { status: 'success', newId: newLessonId };
    }
  } catch (e) { 
    return { status: 'error', message: e.toString() }; 
  }
}

function deleteLesson(lessonId) {
  try {
    const sectionsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SECTIONS_SHEET_NAME);
    const sectionsData = sectionsSheet.getDataRange().getValues();
    const sectionsIdColIndex = sectionsData[0].indexOf('LessonID');
    for (let i = sectionsData.length - 1; i > 0; i--) {
      if (sectionsData[i][sectionsIdColIndex] === lessonId) {
        sectionsSheet.deleteRow(i + 1);
      }
    }
    
    const lessonsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LESSONS_SHEET_NAME);
    const lessonsData = lessonsSheet.getDataRange().getValues();
    const lessonsIdColIndex = lessonsData[0].indexOf('LessonID');
    for (let i = 1; i < lessonsData.length; i++) {
      if (lessonsData[i][lessonsIdColIndex] === lessonId) { 
        lessonsSheet.deleteRow(i + 1); 
        break; 
      }
    }
    return { status: 'success' };
  } catch (e) { 
    return { status: 'error', message: e.toString() }; 
  }
}

function saveSection(data) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SECTIONS_SHEET_NAME);
    const sheetData = sheet.getDataRange().getValues();
    const headers = sheetData[0];
    
    let vpCol = headers.indexOf('VideoPoints');
    if (vpCol === -1) { 
      vpCol = headers.length; 
      sheet.getRange(1, vpCol + 1).setValue('VideoPoints'); 
      headers.push('VideoPoints'); 
    }
    
    let qpCol = headers.indexOf('QuizPoints');
    if (qpCol === -1) { 
      qpCol = headers.length; 
      sheet.getRange(1, qpCol + 1).setValue('QuizPoints'); 
      headers.push('QuizPoints'); 
    }

    // JAVIS UPDATE: สร้างคอลัมน์ ContentType อัตโนมัติถ้าไม่มี
    let typeCol = headers.indexOf('ContentType');
    if (typeCol === -1) {
      typeCol = headers.length;
      sheet.getRange(1, typeCol + 1).setValue('ContentType');
      headers.push('ContentType');
    }

    const idCol = headers.indexOf('SectionID');
    const lessonIdCol = headers.indexOf('LessonID');
    const titleCol = headers.indexOf('Title');
    // ใช้ YouTubeVideoID เดิมเก็บ Data URL ของ Content ไปเลย เพื่อความเข้ากันได้ย้อนหลัง
    const ytCol = headers.indexOf('YouTubeVideoID');
    const orderCol = headers.indexOf('Order');

    if (data.SectionID) {
      let rowIndex = -1;
      for (let i = 1; i < sheetData.length; i++) {
        if (sheetData[i][idCol] == data.SectionID) { 
          rowIndex = i + 1; 
          break; 
        }
      }
      if (rowIndex > -1) {
        sheet.getRange(rowIndex, titleCol + 1).setValue(data.Title);
        sheet.getRange(rowIndex, ytCol + 1).setValue(data.YouTubeVideoID);
        sheet.getRange(rowIndex, orderCol + 1).setValue(data.Order);
        sheet.getRange(rowIndex, vpCol + 1).setValue(data.VideoPoints !== undefined ? data.VideoPoints : 10);
        sheet.getRange(rowIndex, qpCol + 1).setValue(data.QuizPoints !== undefined ? data.QuizPoints : 50);
        sheet.getRange(rowIndex, typeCol + 1).setValue(data.ContentType || 'youtube');
      } else { 
        return { status: 'error', message: 'Section not found' }; 
      }
      return { status: 'success' };
    } 
    else {
      const newId = 'SEC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
      const newRow = new Array(headers.length).fill('');
      newRow[idCol] = newId;
      newRow[lessonIdCol] = data.LessonID;
      newRow[titleCol] = data.Title;
      newRow[ytCol] = data.YouTubeVideoID;
      newRow[orderCol] = data.Order;
      newRow[vpCol] = data.VideoPoints !== undefined ? data.VideoPoints : 10;
      newRow[qpCol] = data.QuizPoints !== undefined ? data.QuizPoints : 50;
      newRow[typeCol] = data.ContentType || 'youtube';
      
      sheet.appendRow(newRow);
      return { status: 'success', newId: newId };
    }
  } catch (e) { 
    return { status: 'error', message: e.toString() }; 
  }
}

function deleteSection(sectionId) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SECTIONS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const idColIndex = data[0].indexOf('SectionID');
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === sectionId) { 
        sheet.deleteRow(i + 1); 
        return { status: 'success' }; 
      }
    }
    return { status: 'error', message: 'Section not found.' };
  } catch (e) { 
    return { status: 'error', message: e.toString() }; 
  }
}

function saveQuestion(questionData) {
    const permissions = getCurrentUserPermissions(Session.getActiveUser().getEmail());
    if (!permissions.canEditQuizzes) return { error: 'Permission Denied' };
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(QUIZZES_SHEET_NAME);
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        if (questionData.id) {
            const data = sheet.getDataRange().getValues();
            const idColIndex = headers.indexOf('QuestionID');
            for (let i = 1; i < data.length; i++) {
                if (data[i][idColIndex] === questionData.id) {
                    const rowToUpdate = i + 1;
                    sheet.getRange(rowToUpdate, headers.indexOf('QuestionText') + 1).setValue(questionData.questionText);
                    sheet.getRange(rowToUpdate, headers.indexOf('Options') + 1).setValue(JSON.stringify(questionData.options));
                    sheet.getRange(rowToUpdate, headers.indexOf('CorrectAnswer') + 1).setValue(questionData.correctAnswer);
                    return { status: 'success', message: 'Question updated.' };
                }
            }
            return { status: 'error', message: 'Question ID not found.' };
        } else {
            const newId = 'QZ-' + Utilities.getUuid().substring(0, 8).toUpperCase();
            sheet.appendRow([newId, questionData.sectionId, questionData.questionText, JSON.stringify(questionData.options), questionData.correctAnswer]);
            return { status: 'success', message: 'Question added.', newId: newId };
        }
    } catch (e) {
        return { status: 'error', message: e.toString() };
    }
}

function deleteQuestion(questionId) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(QUIZZES_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const idColIndex = data[0].indexOf('QuestionID');
    for(let i = data.length - 1; i > 0; i--) {
      if(data[i][idColIndex] === questionId) { 
        sheet.deleteRow(i + 1); 
        return { status: 'success' }; 
      }
    }
    return { status: 'error', message: 'Question not found.' };
  } catch (e) { 
    return { status: 'error', message: e.toString() }; 
  }
}

function saveUserAdmin(userData) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USER_PROFILES_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailCol = headers.indexOf('Email');
    const nameCol = headers.indexOf('FullName');
    const nickCol = headers.indexOf('Nickname');
    const deptCol = headers.indexOf('Department');
    const posCol = headers.indexOf('Position');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][emailCol]).trim().toLowerCase() === String(userData.Email).trim().toLowerCase()) {
        if (nameCol > -1) sheet.getRange(i + 1, nameCol + 1).setValue(userData.FullName);
        if (nickCol > -1) sheet.getRange(i + 1, nickCol + 1).setValue(userData.Nickname);
        if (deptCol > -1) sheet.getRange(i + 1, deptCol + 1).setValue(userData.Department);
        if (posCol > -1) sheet.getRange(i + 1, posCol + 1).setValue(userData.Position);
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'ไม่พบผู้ใช้งานนี้ในระบบ' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ----------------------------------------------------
// ระบบ SUPER ADMIN (จัดการสิทธิ์ และ ดูสถิติ)
// ----------------------------------------------------
function getAdmins() {
  const permissions = getCurrentUserPermissions(Session.getActiveUser().getEmail());
  if (!permissions.isSuperAdmin) return { error: 'Permission Denied' };
  
  return getSheetData(ADMINS_SHEET_NAME);
}

function updateAdminPermissions(email, newPermissions) {
  const currentUserPermissions = getCurrentUserPermissions(Session.getActiveUser().getEmail());
  if (!currentUserPermissions.isSuperAdmin) return { status: 'error', message: 'Permission Denied.' };
  
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ADMINS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailColIndex = headers.indexOf('Email');
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][emailColIndex]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
        const rowToUpdate = i + 1;
        sheet.getRange(rowToUpdate, headers.indexOf('CanEditLessons') + 1).setValue(newPermissions.canEditLessons);
        sheet.getRange(rowToUpdate, headers.indexOf('CanEditQuizzes') + 1).setValue(newPermissions.canEditQuizzes);
        sheet.getRange(rowToUpdate, headers.indexOf('IsSuperAdmin') + 1).setValue(newPermissions.isSuperAdmin);
        return { status: 'success' };
      }
    }
    
    sheet.appendRow([email, newPermissions.canEditLessons, newPermissions.canEditQuizzes, newPermissions.isSuperAdmin]);
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// ==============================================================
// 📊 ฟังก์ชันใหม่: ดึงสถิติ Dashboard และ Leaderboard ฉบับสมบูรณ์
// ==============================================================
function getDashboardStats() {
  const permissions = getCurrentUserPermissions(Session.getActiveUser().getEmail());
  if (!permissions.canEditLessons && !permissions.canEditQuizzes && !permissions.isSuperAdmin) return { error: 'Permission Denied' };

  const usersProgress = getSheetData(PROGRESS_SHEET_NAME);
  const scores = getSheetData(QUIZ_SCORES_SHEET_NAME);
  const sections = getSheetData(SECTIONS_SHEET_NAME);
  const profiles = getSheetData(USER_PROFILES_SHEET_NAME);
  const lessons = getSheetData(LESSONS_SHEET_NAME).filter(l => l.IsActive);

  const activeLessonIds = lessons.map(l => l.LessonID);
  const activeSections = sections.filter(s => activeLessonIds.includes(s.LessonID));
  const totalSections = activeSections.length;

  const sectionMap = {};
  activeSections.forEach(s => {
     sectionMap[s.SectionID] = {
         vp: parseInt(s.VideoPoints) || 10,
         qp: parseInt(s.QuizPoints) || 50
     };
  });

  const studentMap = {};

  profiles.forEach(p => {
      const email = String(p.Email).trim().toLowerCase();
      studentMap[email] = {
          email: email,
          fullName: p.FullName || email,
          nickname: p.Nickname || '-',
          department: p.Department || 'General',
          position: p.Position || '-',
          profileImage: p.ProfileImage || '',
          completedSections: [],
          scores: {},
          totalBN: 0,
          lastActive: null
      };
  });

  usersProgress.forEach(p => {
      const email = String(p.Email).trim().toLowerCase();
      if(!studentMap[email]) {
          studentMap[email] = { email: email, fullName: email, nickname: '-', department: '-', position: '-', profileImage: '', completedSections: [], scores: {}, totalBN: 0, lastActive: null };
      }
      
      if (sectionMap[p.SectionID] && !studentMap[email].completedSections.includes(p.SectionID)) {
          studentMap[email].completedSections.push(p.SectionID);
          studentMap[email].totalBN += sectionMap[p.SectionID].vp; 
      }
      if (!studentMap[email].lastActive || new Date(p.Timestamp) > new Date(studentMap[email].lastActive)) {
          studentMap[email].lastActive = p.Timestamp;
      }
  });

  scores.forEach(s => {
      const email = String(s.Email).trim().toLowerCase();
      if(!studentMap[email]) {
          studentMap[email] = { email: email, fullName: email, nickname: '-', department: '-', position: '-', profileImage: '', completedSections: [], scores: {}, totalBN: 0, lastActive: null };
      }
      
      if (sectionMap[s.SectionID]) {
          const percent = s.TotalQuestions > 0 ? (s.Score / s.TotalQuestions) : 0;
          
          if (percent >= 0.5) {
             if (!studentMap[email].scores[s.SectionID]) {
                 studentMap[email].scores[s.SectionID] = percent;
                 studentMap[email].totalBN += sectionMap[s.SectionID].qp; 
             }
          } else {
             if (!studentMap[email].scores[s.SectionID]) {
                 studentMap[email].scores[s.SectionID] = percent; 
             }
          }
      }

      if (!studentMap[email].lastActive || new Date(s.Timestamp) > new Date(studentMap[email].lastActive)) {
          studentMap[email].lastActive = s.Timestamp;
      }
  });

  const students = Object.values(studentMap).map(std => {
      std.progressPercent = totalSections > 0 ? Math.round((std.completedSections.length / totalSections) * 100) : 0;
      
      const scoreValues = Object.values(std.scores);
      if (scoreValues.length > 0) {
          const sumPercent = scoreValues.reduce((sum, val) => sum + (val * 100), 0);
          std.averageScore = Math.round(sumPercent / scoreValues.length);
      } else {
          std.averageScore = 0;
      }
      
      if (std.lastActive) {
          const d = new Date(std.lastActive);
          std.lastActiveStr = d.toLocaleDateString('th-TH') + ' ' + d.toLocaleTimeString('th-TH', {hour: '2-digit', minute:'2-digit'});
      } else {
          std.lastActiveStr = '-';
      }
      
      return std;
  }).filter(std => std.completedSections.length > 0 || Object.keys(std.scores).length > 0); 

  return JSON.stringify({
      totalStudents: students.length,
      totalSections: totalSections,
      totalLessons: lessons.length,
      students: students.sort((a,b) => b.totalBN - a.totalBN) 
  });
}

function saveMasterData(masterData) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Settings');
    if (!sheet) return { status: 'error', message: 'หาชีต Settings ไม่พบ' };

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
    }

    const maxLen = Math.max(
      masterData.categories ? masterData.categories.length : 0,
      masterData.departments ? masterData.departments.length : 0,
      masterData.positions ? masterData.positions.length : 0
    );

    const writeData = [];
    for (let i = 0; i < maxLen; i++) {
      writeData.push([
        masterData.categories && masterData.categories[i] ? masterData.categories[i] : '',
        masterData.departments && masterData.departments[i] ? masterData.departments[i] : '',
        masterData.positions && masterData.positions[i] ? masterData.positions[i] : ''
      ]);
    }

    if (writeData.length > 0) {
      sheet.getRange(2, 1, writeData.length, 3).setValues(writeData);
    }

    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}
