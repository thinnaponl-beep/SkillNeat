// ===============================================================
// FILE: Code.gs (Full Complete Version)
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
const EXTERNAL_PROOFS_SHEET_NAME = 'ExternalProofs';
const MANUAL_SCORES_SHEET_NAME = 'ManualScoresLog';

// ----------------------------------------------------
// ระบบ ROUTING (แสดงผลหน้าจอหลัก)
// ----------------------------------------------------
function doGet(e){
  const userEmail = Session.getActiveUser().getEmail().trim().toLowerCase();
  const permissions = getCurrentUserPermissions(userEmail);
  const page = e.parameter.page;
  
  let userProfile = { Email: userEmail, FullName: '', Nickname: '', Department: 'General', Position: 'พนักงานทั่วไป', JobTitle: '', ProfileImage: '', ManualBN: 0 };
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
  const cache = CacheService.getScriptCache();
  const cacheKey = 'CACHE_SHEET_' + sheetName;
  
  const cacheableSheets = [LESSONS_SHEET_NAME, SECTIONS_SHEET_NAME, QUIZZES_SHEET_NAME, 'Settings'];
  
  if (cacheableSheets.includes(sheetName)) {
      const cachedData = cache.get(cacheKey);
      if (cachedData) {
          try { return JSON.parse(cachedData); } catch(e) {}
      }
  }

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) {
      if (sheetName === EXTERNAL_PROOFS_SHEET_NAME || sheetName === MANUAL_SCORES_SHEET_NAME) return []; 
      throw new Error(`Sheet "${sheetName}" not found.`);
    }
    
    const rows = sheet.getDataRange().getValues();
    if (rows.length === 0) return [];
    const headers = rows.shift();
    const data = rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    if (cacheableSheets.includes(sheetName)) {
        const jsonString = JSON.stringify(data);
        if (jsonString.length < 100000) {
            cache.put(cacheKey, jsonString, 60); 
        }
    }

    return data;
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

function ensureExternalProofsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(EXTERNAL_PROOFS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(EXTERNAL_PROOFS_SHEET_NAME);
    sheet.appendRow(['Email', 'SectionID', 'FileUrl', 'Status', 'Timestamp']);
  }
  return sheet;
}

function ensureManualScoresSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(MANUAL_SCORES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MANUAL_SCORES_SHEET_NAME);
    sheet.appendRow(['Email', 'Points', 'Reason', 'Timestamp', 'AdminEmail']);
  }
  return sheet;
}

function isDateInPeriod(dateObj, period) {
  if (!period || period === 'all') return true;
  if (!dateObj) return false;
  try {
    let d;
    if (dateObj instanceof Date) {
       d = new Date(dateObj.getTime());
    } else {
       let str = String(dateObj).trim();
       if (/^\d{1,4}[/-]\d{1,2}[/-]\d{1,4}/.test(str)) {
           const parts = str.split(/[/\sT:-]/);
           let p0 = parseInt(parts[0], 10);
           let p1 = parseInt(parts[1], 10);
           let p2 = parseInt(parts[2], 10);

           let year, month, day;
           if (p0 > 1000) { year = p0; month = p1 - 1; day = p2; } 
           else { day = p0; month = p1 - 1; year = p2; }

           if (year < 100) year += 2000;
           if (year > 2500) year -= 543; 

           d = new Date(year, month, day);
       } else {
           d = new Date(str);
       }
    }
    if (!d || isNaN(d.getTime())) return false;
    if (d.getFullYear() > 2500) d.setFullYear(d.getFullYear() - 543);
    
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    const recordPeriod = `${year}-${month}`;
    return recordPeriod === period;
  } catch(e) {
    return false;
  }
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
  
  const rawProgress = getSheetData(PROGRESS_SHEET_NAME)
      .filter(p => String(p.Email).trim().toLowerCase() === userEmail);
  const progress = rawProgress.map(p => p.SectionID);
  
  const timestamps = getSheetData(TIMESTAMPS_SHEET_NAME)
      .filter(t => String(t.Email).trim().toLowerCase() === userEmail);
      
  const scores = getSheetData(QUIZ_SCORES_SHEET_NAME)
      .filter(s => String(s.Email).trim().toLowerCase() === userEmail);
      
  const allQuizzes = getSheetData(QUIZZES_SHEET_NAME);
  
  const externalProofs = getSheetData(EXTERNAL_PROOFS_SHEET_NAME)
      .filter(p => String(p.Email).trim().toLowerCase() === userEmail);

  const manualLogs = getSheetData(MANUAL_SCORES_SHEET_NAME)
      .filter(m => String(m.Email).trim().toLowerCase() === userEmail);

  let mainCats = [];
  try {
      const settingsData = getSheetData('Settings');
      if (settingsData && settingsData.length > 0) {
          mainCats = settingsData.map(r => r.MainCategories).filter(Boolean);
      }
  } catch(e) {}

  let userProfile = { Email: userEmail, FullName: '', Nickname: '', Department: 'General', Position: 'พนักงานทั่วไป', JobTitle: '', ProfileImage: '', ManualBN: 0 };
  const profiles = getSheetData(USER_PROFILES_SHEET_NAME);
  if (profiles && profiles.length > 0) {
     const found = profiles.find(p => String(p.Email).trim().toLowerCase() === userEmail);
     if (found) userProfile = found;
  }

  let totalBN = parseInt(userProfile.ManualBN) || 0;

  const historyMap = {};
  function addHistory(dateStr, points) {
      if (!dateStr || isNaN(points) || points === 0) return;
      let d;
      if (dateStr instanceof Date) d = dateStr;
      else {
          const str = String(dateStr);
          if (str.includes('/')) {
               const parts = str.split(/[/\sT:-]/);
               if (parts.length >= 3) {
                   let p0 = parseInt(parts[0], 10);
                   let p1 = parseInt(parts[1], 10);
                   let p2 = parseInt(parts[2], 10);
                   let day, month, year;
                   if (p0 > 1000) { year = p0; month = p1 - 1; day = p2; } 
                   else { day = p0; month = p1 - 1; year = p2; }
                   if (year < 100) year += 2000;
                   if (year > 2500) year -= 543;
                   d = new Date(year, month, day);
               }
          } else { d = new Date(dateStr); }
      }
      if (!d || isNaN(d.getTime())) return;
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      if (!historyMap[key]) historyMap[key] = { total: 0 };
      historyMap[key].total += points;
  }

  let studentEarnedSections = {};
  rawProgress.forEach(p => {
      if (!studentEarnedSections[p.SectionID]) {
          studentEarnedSections[p.SectionID] = true;
          const sec = sections.find(s => s.SectionID === p.SectionID);
          let t = p.Timestamp;
          if (!t && p.EarnedBN instanceof Date) t = p.EarnedBN; 
          let earned = (p.EarnedBN !== undefined && p.EarnedBN !== '' && !(p.EarnedBN instanceof Date)) ? parseInt(p.EarnedBN) : (sec && sec.VideoPoints !== undefined ? parseInt(sec.VideoPoints) : 10);
          addHistory(t, earned);
      }
  });

  let studentPassedQuizzes = {};
  const sortedScores = [...scores].sort((a, b) => new Date(a.Timestamp) - new Date(b.Timestamp));
  sortedScores.forEach(s => {
      if (!studentPassedQuizzes[s.SectionID]) {
          const percent = s.TotalQuestions > 0 ? (s.Score / s.TotalQuestions) : 0;
          if (percent >= 0.5) {
              studentPassedQuizzes[s.SectionID] = true;
              const sec = sections.find(sec => sec.SectionID === s.SectionID);
              let t = s.Timestamp;
              if (!t && s.EarnedBN instanceof Date) t = s.EarnedBN;
              let earned = (s.EarnedBN !== undefined && s.EarnedBN !== '' && !(s.EarnedBN instanceof Date)) ? parseInt(s.EarnedBN) : (sec && sec.QuizPoints !== undefined ? parseInt(sec.QuizPoints) : 50);
              addHistory(t, earned);
          }
      }
  });

  manualLogs.forEach(m => { addHistory(m.Timestamp, parseInt(m.Points) || 0); });

  const monthlyHistory = Object.keys(historyMap).map(k => {
      return { month: k, score: historyMap[k].total };
  }).sort((a, b) => b.month.localeCompare(a.month)); 

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
        section.ContentType = section.ContentType || 'youtube';
        const proof = externalProofs.filter(p => p.SectionID === section.SectionID)
                                    .sort((a,b) => new Date(b.Timestamp) - new Date(a.Timestamp))[0];
        section.externalProof = proof || null;

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
    
    lesson.isCompleted = lesson.sections.length > 0 && lesson.sections.every(s => {
        const isVideoDone = progress.includes(s.SectionID);
        const isQuizDone = !s.hasQuiz || (s.quizScore && (s.quizScore.Score / s.quizScore.TotalQuestions) >= 0.5);
        return isVideoDone && isQuizDone;
    });

    lesson.isStarted = lesson.sections.some(s => {
        return progress.includes(s.SectionID) || s.latestTimestamp > 0 || s.quizScore || s.externalProof;
    });
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
      
      let isPersonalized = false;
      if (lesson.TargetDepartments) {
          const targetDepts = lesson.TargetDepartments.split(',').map(d => d.trim());
          if (targetDepts.includes(userDept) && !targetDepts.includes('ทั้งหมด (All)')) {
              isPersonalized = true;
          }
      }
      if (lesson.TargetPositions) {
          const targetPos = lesson.TargetPositions.split(',').map(p => p.trim());
          if (targetPos.includes(userPos) && !targetPos.includes('ทั้งหมด (All)')) {
              isPersonalized = true;
          }
      }
      return isPersonalized;
  });

  const personalizedIds = personalizedLessons.map(l => l.LessonID);
  const featuredLessons = accessibleLessons.filter(lesson => !personalizedIds.includes(lesson.LessonID));

  return {
    userInfo: { email: userEmail, progress: progress, profile: userProfile, totalBN: totalBN, monthlyHistory: monthlyHistory },
    lessons: lessons,
    featuredLessons: featuredLessons,
    personalizedLessons: personalizedLessons,
    mainCategories: mainCats
  };
}

function getPublicLeaderboard(period) {
  period = period || 'current_month';

  let targetPeriod = period;
  if (period === 'current_month') {
      const d = new Date();
      targetPeriod = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  }

  const allUsersProgress = getSheetData(PROGRESS_SHEET_NAME);
  const allScores = getSheetData(QUIZ_SCORES_SHEET_NAME);
  const sections = getSheetData(SECTIONS_SHEET_NAME);
  const profiles = getSheetData(USER_PROFILES_SHEET_NAME);
  const allManualLogs = getSheetData(MANUAL_SCORES_SHEET_NAME);

  const usersProgress = allUsersProgress.filter(p => {
      let t = p.Timestamp;
      if (!t && p.EarnedBN instanceof Date) t = p.EarnedBN;
      return isDateInPeriod(t, targetPeriod);
  });
  
  const scores = allScores.filter(s => {
      let t = s.Timestamp;
      if (!t && s.EarnedBN instanceof Date) t = s.EarnedBN;
      return isDateInPeriod(t, targetPeriod);
  });
  
  const manualLogs = allManualLogs.filter(m => isDateInPeriod(m.Timestamp, targetPeriod));

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
      const baseManualBN = period === 'all' ? (parseInt(p.ManualBN) || 0) : 0;

      studentMap[email] = {
          email: email,
          name: p.Nickname || p.FullName || email.split('@')[0],
          dept: p.Department || 'General',
          img: p.ProfileImage || defaultImg,
          completedSections: [],
          scores: {},
          totalBN: baseManualBN 
      };
  });

  if (period !== 'all') {
      manualLogs.forEach(m => {
          const email = String(m.Email).trim().toLowerCase();
          if (studentMap[email]) {
              studentMap[email].totalBN += (parseInt(m.Points) || 0);
          }
      });
  }

  usersProgress.forEach(p => {
      const email = String(p.Email).trim().toLowerCase();
      if(!studentMap[email]) { 
          studentMap[email] = { email: email, name: email.split('@')[0], dept: 'General', img: 'https://ui-avatars.com/api/?name='+email+'&background=31c1d7&color=fff', completedSections: [], scores: {}, totalBN: 0 };
      }
      if (sectionMap[p.SectionID] && !studentMap[email].completedSections.includes(p.SectionID)) {
          studentMap[email].completedSections.push(p.SectionID);
          let earned = (p.EarnedBN !== undefined && p.EarnedBN !== '' && !(p.EarnedBN instanceof Date)) ? parseInt(p.EarnedBN) : (sectionMap[p.SectionID] ? sectionMap[p.SectionID].vp : 0);
          studentMap[email].totalBN += earned; 
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
                 let earned = (s.EarnedBN !== undefined && s.EarnedBN !== '' && !(s.EarnedBN instanceof Date)) ? parseInt(s.EarnedBN) : (sectionMap[s.SectionID] ? sectionMap[s.SectionID].qp : 0);
                 studentMap[email].totalBN += earned; 
             }
          }
      }
  });

  const activeStudents = Object.values(studentMap).filter(std => std.totalBN > 0 || std.totalBN < 0);
  return activeStudents.sort((a,b) => b.totalBN - a.totalBN);
}

function saveQuizScore(scoreData) {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return { status: 'error', message: 'User not logged in.' };

    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(QUIZ_SCORES_SHEET_NAME);
        const data = sheet.getDataRange().getValues();
        let headers = data[0] || [];
        
        if (headers.length === 0) {
            headers = ['Email', 'SectionID', 'Score', 'TotalQuestions', 'Timestamp', 'EarnedBN'];
            sheet.appendRow(headers);
        } else {
            if (headers.indexOf('Timestamp') === -1) {
                headers.push('Timestamp');
                sheet.getRange(1, headers.length).setValue('Timestamp');
            }
            if (headers.indexOf('EarnedBN') === -1) {
                headers.push('EarnedBN');
                sheet.getRange(1, headers.length).setValue('EarnedBN');
            }
        }

        const passed = (scoreData.score / scoreData.totalQuestions) >= 0.5;
        let earnedPoints = 0;
        
        if (passed) {
            const sections = getSheetData(SECTIONS_SHEET_NAME);
            const sec = sections.find(s => s.SectionID === scoreData.sectionId);
            earnedPoints = sec && sec.QuizPoints ? parseInt(sec.QuizPoints) : 50;
        }

        const newRow = new Array(headers.length).fill('');
        newRow[headers.indexOf('Email')] = userEmail;
        newRow[headers.indexOf('SectionID')] = scoreData.sectionId;
        
        let scoreIdx = headers.indexOf('Score');
        if(scoreIdx > -1) newRow[scoreIdx] = scoreData.score;
        else newRow[2] = scoreData.score; 
        
        let tqIdx = headers.indexOf('TotalQuestions');
        if(tqIdx > -1) newRow[tqIdx] = scoreData.totalQuestions;
        else newRow[3] = scoreData.totalQuestions;

        newRow[headers.indexOf('Timestamp')] = new Date();
        newRow[headers.indexOf('EarnedBN')] = earnedPoints;

        sheet.appendRow(newRow);
        
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
  let headers = data[0] || [];
  
  if (headers.length === 0) {
      headers = ['Email', 'SectionID', 'Timestamp', 'EarnedBN'];
      progressSheet.appendRow(headers);
  } else {
      if (headers.indexOf('Timestamp') === -1) {
          headers.push('Timestamp');
          progressSheet.getRange(1, headers.length).setValue('Timestamp');
      }
      if (headers.indexOf('EarnedBN') === -1) {
          headers.push('EarnedBN');
          progressSheet.getRange(1, headers.length).setValue('EarnedBN');
      }
  }

  const alreadyExists = data.some(row => String(row[0]).trim().toLowerCase() === safeEmail && row[1] === sectionId);
  
  if (!alreadyExists) {
    const sections = getSheetData(SECTIONS_SHEET_NAME);
    const sec = sections.find(s => s.SectionID === sectionId);
    const videoPoints = sec && sec.VideoPoints !== undefined ? parseInt(sec.VideoPoints) : 10;

    const newRow = new Array(headers.length).fill('');
    newRow[headers.indexOf('Email')] = userEmail;
    newRow[headers.indexOf('SectionID')] = sectionId;
    newRow[headers.indexOf('Timestamp')] = new Date();
    newRow[headers.indexOf('EarnedBN')] = videoPoints;

    progressSheet.appendRow(newRow);
    return { status: 'success' };
  }
  return { status: 'info', message: 'Already recorded.' };
}

function recordProgressAdmin(email, sectionId) {
  const progressSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PROGRESS_SHEET_NAME);
  const data = progressSheet.getDataRange().getValues();
  let headers = data[0] || [];
  
  if (headers.length === 0) {
      headers = ['Email', 'SectionID', 'Timestamp', 'EarnedBN'];
      progressSheet.appendRow(headers);
  } else {
      if (headers.indexOf('Timestamp') === -1) {
          headers.push('Timestamp');
          progressSheet.getRange(1, headers.length).setValue('Timestamp');
      }
      if (headers.indexOf('EarnedBN') === -1) {
          headers.push('EarnedBN');
          progressSheet.getRange(1, headers.length).setValue('EarnedBN');
      }
  }

  const alreadyExists = data.some(row => String(row[0]).trim().toLowerCase() === email.trim().toLowerCase() && row[1] === sectionId);
  
  if (!alreadyExists) {
    const sections = getSheetData(SECTIONS_SHEET_NAME);
    const sec = sections.find(s => s.SectionID === sectionId);
    const videoPoints = sec && sec.VideoPoints !== undefined ? parseInt(sec.VideoPoints) : 10;

    const newRow = new Array(headers.length).fill('');
    newRow[headers.indexOf('Email')] = email.trim().toLowerCase();
    newRow[headers.indexOf('SectionID')] = sectionId;
    newRow[headers.indexOf('Timestamp')] = new Date();
    newRow[headers.indexOf('EarnedBN')] = videoPoints;

    progressSheet.appendRow(newRow);
    return { status: 'success' };
  }
  return { status: 'info', message: 'Already recorded.' };
}

function removeProgressAdmin(email, sectionId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PROGRESS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
      if (String(data[i][0]).trim().toLowerCase() === email && data[i][1] === sectionId) {
          sheet.deleteRow(i + 1);
      }
  }
}

// ----------------------------------------------------
// ระบบรับไฟล์รูปภาพและการอนุมัติ
// ----------------------------------------------------
function submitExternalProof(data) {
  const userEmail = Session.getActiveUser().getEmail();
  const safeEmail = userEmail.trim().toLowerCase();
  if (!userEmail) return { status: 'error', message: 'User not logged in' };

  try {
    let fileUrl = '';
    if (data.imageData) {
      const folderId = "1uYLnZNVG7-Plmf2K_HO0qijLq4l4qwYQ"; 
      const folder = DriveApp.getFolderById(folderId);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const blob = Utilities.newBlob(Utilities.base64Decode(data.imageData), data.imageMimeType, userEmail + "_proof_" + data.sectionId);
      const file = folder.createFile(blob);
      fileUrl = file.getUrl();
    }
    
    let sheet = ensureExternalProofsSheet();
    let dataRange = sheet.getDataRange().getValues();
    let foundIndex = -1;
    
    for(let i = 1; i < dataRange.length; i++) {
        if(String(dataRange[i][0]).trim().toLowerCase() === safeEmail && dataRange[i][1] === data.sectionId) {
            foundIndex = i + 1; 
            break;
        }
    }
    
    if (foundIndex > -1) {
        sheet.getRange(foundIndex, 3).setValue(fileUrl);
        sheet.getRange(foundIndex, 4).setValue('Pending');
        sheet.getRange(foundIndex, 5).setValue(new Date());
    } else {
        sheet.appendRow([safeEmail, data.sectionId, fileUrl, 'Pending', new Date()]);
    }
    
    return { status: 'success', fileUrl: fileUrl };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function adminUpdateProofStatus(email, sectionId, status) {
  const permissions = getCurrentUserPermissions(Session.getActiveUser().getEmail());
  if (!permissions.canEditLessons && !permissions.isSuperAdmin) return { status: 'error', message: 'Permission Denied' };
  
  try {
      let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXTERNAL_PROOFS_SHEET_NAME);
      if(!sheet) return {status: 'error', message:'Sheet not found'};
      
      let data = sheet.getDataRange().getValues();
      for(let i = 1; i < data.length; i++) {
          if(String(data[i][0]).trim().toLowerCase() === String(email).trim().toLowerCase() && data[i][1] === sectionId) {
              sheet.getRange(i+1, 4).setValue(status);
              
              if (status === 'Approved') {
                  recordProgressAdmin(email, sectionId);
              } else if (status === 'Rejected') {
                  removeProgressAdmin(email, sectionId);
              }
              return {status: 'success'};
          }
      }
      return {status: 'error', message: 'ไม่พบประวัติการส่งหลักฐาน'};
  } catch (e) {
      return {status: 'error', message: e.toString()};
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
    
    let manualCol = headers.indexOf('ManualBN');
    if (manualCol === -1) {
        manualCol = headers.length;
        sheet.getRange(1, manualCol + 1).setValue('ManualBN');
        headers.push('ManualBN');
    }

    let jobCol = headers.indexOf('JobTitle');
    if (jobCol === -1) {
       jobCol = headers.length;
       sheet.getRange(1, jobCol + 1).setValue('JobTitle');
       headers.push('JobTitle');
    }
    
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
      if (nameCol > -1) sheet.getRange(rowIndex, nameCol + 1).setValue(data.fullName);
      if (nickCol > -1) sheet.getRange(rowIndex, nickCol + 1).setValue(data.nickname);
      if (jobCol > -1 && data.jobTitle !== undefined) sheet.getRange(rowIndex, jobCol + 1).setValue(data.jobTitle);
      if (profileImageUrl && imgCol > -1) {
        sheet.getRange(rowIndex, imgCol + 1).setValue(profileImageUrl);
      }
    } else {
      const newRow = new Array(headers.length).fill('');
      newRow[emailCol] = userEmail;
      newRow[nameCol] = data.fullName;
      newRow[nickCol] = data.nickname;
      newRow[deptCol] = 'General';
      newRow[posCol] = 'พนักงาน';
      newRow[manualCol] = 0;
      newRow[jobCol] = data.jobTitle || '';
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
        });
    });

    let settings = { mainCategories: [], categories: [], departments: [], positions: [] };
    try {
        const settingsData = getSheetData('Settings');
        if (settingsData && settingsData.length > 0) {
            settings.mainCategories = settingsData.map(r => r.MainCategories).filter(Boolean);
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
    
    let mainCatCol = headers.indexOf('MainCategory');
    if (mainCatCol === -1) {
      mainCatCol = headers.length;
      sheet.getRange(1, mainCatCol + 1).setValue('MainCategory');
      headers.push('MainCategory');
    }
    
    let featCol = headers.indexOf('IsFeatured');
    if (featCol === -1) {
      featCol = headers.length;
      sheet.getRange(1, featCol + 1).setValue('IsFeatured');
      headers.push('IsFeatured');
    }

    let catFeatCol = headers.indexOf('IsCategoryFeatured');
    if (catFeatCol === -1) {
      catFeatCol = headers.length;
      sheet.getRange(1, catFeatCol + 1).setValue('IsCategoryFeatured');
      headers.push('IsCategoryFeatured');
    }

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
        if (featCol > -1) sheet.getRange(rowIndex, featCol + 1).setValue(data.IsFeatured); 
        if (catFeatCol > -1) sheet.getRange(rowIndex, catFeatCol + 1).setValue(data.IsCategoryFeatured); 
        if (categoryCol > -1) sheet.getRange(rowIndex, categoryCol + 1).setValue(data.Category);
        if (mainCatCol > -1) sheet.getRange(rowIndex, mainCatCol + 1).setValue(data.MainCategory); 
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
      if (featCol > -1) newRow[featCol] = data.IsFeatured; 
      if (catFeatCol > -1) newRow[catFeatCol] = data.IsCategoryFeatured;
      if (categoryCol > -1) newRow[categoryCol] = data.Category;
      if (mainCatCol > -1) newRow[mainCatCol] = data.MainCategory; 
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

    let typeCol = headers.indexOf('ContentType');
    if (typeCol === -1) {
      typeCol = headers.length;
      sheet.getRange(1, typeCol + 1).setValue('ContentType');
      headers.push('ContentType');
    }

    let durCol = headers.indexOf('Duration');
    if (durCol === -1) {
      durCol = headers.length;
      sheet.getRange(1, durCol + 1).setValue('Duration');
      headers.push('Duration');
    }

    const idCol = headers.indexOf('SectionID');
    const lessonIdCol = headers.indexOf('LessonID');
    const titleCol = headers.indexOf('Title');
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
        sheet.getRange(rowIndex, durCol + 1).setValue(data.Duration !== undefined ? data.Duration : 15);
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
      newRow[durCol] = data.Duration !== undefined ? data.Duration : 15;
      
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
    
    let manualCol = headers.indexOf('ManualBN');
    if (manualCol === -1) {
       manualCol = headers.length;
       sheet.getRange(1, manualCol + 1).setValue('ManualBN');
       headers.push('ManualBN');
    }

    let jobCol = headers.indexOf('JobTitle');
    if (jobCol === -1) {
       jobCol = headers.length;
       sheet.getRange(1, jobCol + 1).setValue('JobTitle');
       headers.push('JobTitle');
    }

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
        
        if (jobCol > -1 && userData.JobTitle !== undefined) {
            sheet.getRange(i + 1, jobCol + 1).setValue(userData.JobTitle);
        }
        
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'ไม่พบผู้ใช้งานนี้ในระบบ' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function adminAddManualScore(data) {
  const adminEmail = Session.getActiveUser().getEmail();
  const permissions = getCurrentUserPermissions(adminEmail);
  if (!permissions.isSuperAdmin && !permissions.canEditLessons) return { status: 'error', message: 'Permission Denied' };

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USER_PROFILES_SHEET_NAME);
    const sheetData = sheet.getDataRange().getValues();
    const headers = sheetData[0];

    let manualCol = headers.indexOf('ManualBN');
    if (manualCol === -1) {
      manualCol = headers.length;
      sheet.getRange(1, manualCol + 1).setValue('ManualBN');
      headers.push('ManualBN');
    }

    const emailCol = headers.indexOf('Email');
    let targetEmail = String(data.email).trim().toLowerCase();
    let pointsToAdd = parseInt(data.points) || 0;
    
    let foundUser = false;

    for (let i = 1; i < sheetData.length; i++) {
      if (String(sheetData[i][emailCol]).trim().toLowerCase() === targetEmail) {
        let currentPoints = parseInt(sheetData[i][manualCol]) || 0;
        sheet.getRange(i + 1, manualCol + 1).setValue(currentPoints + pointsToAdd);
        foundUser = true;
        break;
      }
    }
    
    if (!foundUser) return { status: 'error', message: 'ไม่พบอีเมลพนักงานนี้ในระบบ' };

    const logSheet = ensureManualScoresSheet();
    logSheet.appendRow([targetEmail, pointsToAdd, data.reason, new Date(), adminEmail]);

    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function getUserManualScoreLogs(email) {
  try {
    const permissions = getCurrentUserPermissions(Session.getActiveUser().getEmail());
    if (!permissions.isSuperAdmin && !permissions.canEditLessons) return { status: 'error', message: 'Permission Denied' };

    const logs = getSheetData(MANUAL_SCORES_SHEET_NAME);
    const userLogs = logs.filter(log => String(log.Email).trim().toLowerCase() === String(email).trim().toLowerCase());

    userLogs.sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));

    const formattedLogs = userLogs.map(log => {
      let dateStr = '-';
      if (log.Timestamp) {
         const d = new Date(log.Timestamp);
         dateStr = d.toLocaleDateString('th-TH') + ' ' + d.toLocaleTimeString('th-TH', {hour: '2-digit', minute:'2-digit'});
      }
      return {
        points: log.Points,
        reason: log.Reason,
        timestamp: dateStr,
        adminEmail: log.AdminEmail
      };
    });

    return { status: 'success', logs: formattedLogs };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

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

function getDashboardStats(period) {
  period = period || 'all'; 

  const permissions = getCurrentUserPermissions(Session.getActiveUser().getEmail());
  if (!permissions.canEditLessons && !permissions.canEditQuizzes && !permissions.isSuperAdmin) return { error: 'Permission Denied' };

  const allUsersProgress = getSheetData(PROGRESS_SHEET_NAME);
  const allScores = getSheetData(QUIZ_SCORES_SHEET_NAME);
  const sections = getSheetData(SECTIONS_SHEET_NAME);
  const profiles = getSheetData(USER_PROFILES_SHEET_NAME);
  const lessons = getSheetData(LESSONS_SHEET_NAME).filter(l => l.IsActive);
  const allExternalProofs = getSheetData(EXTERNAL_PROOFS_SHEET_NAME);
  const allManualLogs = getSheetData(MANUAL_SCORES_SHEET_NAME); 

  const usersProgress = allUsersProgress.filter(p => {
      let t = p.Timestamp;
      if (!t && p.EarnedBN instanceof Date) t = p.EarnedBN;
      return isDateInPeriod(t, period);
  });
  
  const scores = allScores.filter(s => {
      let t = s.Timestamp;
      if (!t && s.EarnedBN instanceof Date) t = s.EarnedBN;
      return isDateInPeriod(t, period);
  });
  
  const externalProofs = allExternalProofs.filter(p => isDateInPeriod(p.Timestamp, period));
  const manualLogs = allManualLogs.filter(m => isDateInPeriod(m.Timestamp, period)); 

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
      const baseManualBN = period === 'all' ? (parseInt(p.ManualBN) || 0) : 0;
      
      studentMap[email] = {
          email: email,
          fullName: p.FullName || email,
          nickname: p.Nickname || '-',
          department: p.Department || 'General',
          position: p.Position || '-',
          jobTitle: p.JobTitle || '', 
          profileImage: p.ProfileImage || '',
          completedSections: [],
          scores: {},
          manualBN: baseManualBN,
          totalBN: baseManualBN,
          lastActive: null,
          externalProofs: externalProofs.filter(proof => String(proof.Email).trim().toLowerCase() === email)
      };
  });

  if (period !== 'all') {
      manualLogs.forEach(m => {
          const email = String(m.Email).trim().toLowerCase();
          if (studentMap[email]) {
              const pts = parseInt(m.Points) || 0;
              studentMap[email].manualBN += pts;
              studentMap[email].totalBN += pts;
          }
      });
  }

  usersProgress.forEach(p => {
      const email = String(p.Email).trim().toLowerCase();
      if(!studentMap[email]) {
          studentMap[email] = { email: email, fullName: email, nickname: '-', department: '-', position: '-', jobTitle: '', profileImage: '', completedSections: [], scores: {}, manualBN: 0, totalBN: 0, lastActive: null, externalProofs: externalProofs.filter(proof => String(proof.Email).trim().toLowerCase() === email) };
      }
      
      let t = p.Timestamp;
      if (!t && p.EarnedBN instanceof Date) t = p.EarnedBN;

      if (sectionMap[p.SectionID] && !studentMap[email].completedSections.includes(p.SectionID)) {
          studentMap[email].completedSections.push(p.SectionID);
          let earned = (p.EarnedBN !== undefined && p.EarnedBN !== '' && !(p.EarnedBN instanceof Date)) ? parseInt(p.EarnedBN) : (sectionMap[p.SectionID] ? sectionMap[p.SectionID].vp : 0);
          studentMap[email].totalBN += earned; 
      }
      if (!studentMap[email].lastActive || new Date(t) > new Date(studentMap[email].lastActive)) {
          studentMap[email].lastActive = t;
      }
  });

  scores.forEach(s => {
      const email = String(s.Email).trim().toLowerCase();
      if(!studentMap[email]) {
          studentMap[email] = { email: email, fullName: email, nickname: '-', department: '-', position: '-', jobTitle: '', profileImage: '', completedSections: [], scores: {}, manualBN: 0, totalBN: 0, lastActive: null, externalProofs: externalProofs.filter(proof => String(proof.Email).trim().toLowerCase() === email) };
      }
      
      let t = s.Timestamp;
      if (!t && s.EarnedBN instanceof Date) t = s.EarnedBN;

      if (sectionMap[s.SectionID]) {
          const percent = s.TotalQuestions > 0 ? (s.Score / s.TotalQuestions) : 0;
          
          if (percent >= 0.5) {
             if (!studentMap[email].scores[s.SectionID]) {
                 studentMap[email].scores[s.SectionID] = percent;
                 let earned = (s.EarnedBN !== undefined && s.EarnedBN !== '' && !(s.EarnedBN instanceof Date)) ? parseInt(s.EarnedBN) : (sectionMap[s.SectionID] ? sectionMap[s.SectionID].qp : 0);
                 studentMap[email].totalBN += earned; 
             }
          } else {
             if (!studentMap[email].scores[s.SectionID]) {
                 studentMap[email].scores[s.SectionID] = percent; 
             }
          }
      }

      if (!studentMap[email].lastActive || new Date(t) > new Date(studentMap[email].lastActive)) {
          studentMap[email].lastActive = t;
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
  }).filter(std => {
      if (period !== 'all') {
         return std.completedSections.length > 0 || Object.keys(std.scores).length > 0 || std.externalProofs.length > 0 || std.manualBN !== 0;
      }
      return std.completedSections.length > 0 || Object.keys(std.scores).length > 0 || std.externalProofs.length > 0 || std.manualBN !== 0;
  }); 

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

    const headers = ["MainCategories", "Categories", "Departments", "Positions"];
    const lastRow = sheet.getLastRow();
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
    }

    const maxLen = Math.max(
      masterData.mainCategories ? masterData.mainCategories.length : 0,
      masterData.categories ? masterData.categories.length : 0,
      masterData.departments ? masterData.departments.length : 0,
      masterData.positions ? masterData.positions.length : 0
    );

    const writeData = [];
    for (let i = 0; i < maxLen; i++) {
      writeData.push([
        masterData.mainCategories && masterData.mainCategories[i] ? masterData.mainCategories[i] : '',
        masterData.categories && masterData.categories[i] ? masterData.categories[i] : '',
        masterData.departments && masterData.departments[i] ? masterData.departments[i] : '',
        masterData.positions && masterData.positions[i] ? masterData.positions[i] : ''
      ]);
    }

    if (writeData.length > 0) {
      sheet.getRange(2, 1, writeData.length, headers.length).setValues(writeData);
    }

    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function deleteUserAdmin(email) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    const permissions = getCurrentUserPermissions(adminEmail);
    if (!permissions.isSuperAdmin && !permissions.canEditLessons) return { status: 'error', message: 'Permission Denied' };

    const targetEmail = String(email).trim().toLowerCase();
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USER_PROFILES_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const emailCol = data[0].indexOf('Email');
    let deleted = false;
    for (let i = data.length - 1; i > 0; i--) {
      if (String(data[i][emailCol]).trim().toLowerCase() === targetEmail) {
        sheet.deleteRow(i + 1);
        deleted = true;
        break;
      }
    }

    const progSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PROGRESS_SHEET_NAME);
    const progData = progSheet.getDataRange().getValues();
    for (let i = progData.length - 1; i > 0; i--) {
      if (String(progData[i][0]).trim().toLowerCase() === targetEmail) progSheet.deleteRow(i + 1);
    }

    const quizSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(QUIZ_SCORES_SHEET_NAME);
    const quizData = quizSheet.getDataRange().getValues();
    for (let i = quizData.length - 1; i > 0; i--) {
      if (String(quizData[i][0]).trim().toLowerCase() === targetEmail) quizSheet.deleteRow(i + 1);
    }

    if(deleted) return { status: 'success' };
    return { status: 'error', message: 'ไม่พบอีเมลนี้ในระบบ' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}
