// ====================================================
// SCIENTIFIC CENTER - Course Management System
// ERP Lite Full Database Schema
// Google Apps Script Backend with doPost/doGet Support
// ====================================================

// --------------------------------------------
// 1. MAIN FUNCTION TO SETUP THE ENTIRE SYSTEM
// --------------------------------------------
function setupEntireSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Keep only one sheet, delete others
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    if (sheet.getSheetName() !== "Instructions" && sheet.getSheetName() !== "README" && sheet.getSheetName() !== "Sheet1") {
      try {
        ss.deleteSheet(sheet);
      } catch(e) { /* ignore */ }
    }
  });
  
  // 2. Create all sheets
  const sheetNames = [
    "Users", "Roles", "Departments", 
    "Trainers", "Halls", "Courses", 
    "Groups", "Students", "Bookings", 
    "Payments", "Levels", "AddOns"
  ];
  
  const sheetsCreated = {};
  sheetNames.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    sheet = ss.insertSheet(name);
    sheetsCreated[name] = sheet;
  });
  
  // 3. Create schemas
  createTableSchemas(sheetsCreated);
  
  // 4. Insert seed data
  insertSeedData(sheetsCreated);
  
  // 5. Setup relationships
  setupRelationships(sheetsCreated);
  
  // 6. Create custom menu
  createCustomMenu();
  
  SpreadsheetApp.getUi().alert('✅ System Ready!\n\nAll 12 tables created with sample data.\nUse the custom menu: "🏫 Course System"');
}

// --------------------------------------------
// 2. DEFINE ALL TABLE SCHEMAS
// --------------------------------------------
function createTableSchemas(sheets) {
  // Roles
  sheets.Roles.getRange(1,1,1,3).setValues([["role_id", "role_name", "description"]]);
  
  // Users
  sheets.Users.getRange(1,1,1,8).setValues([["user_id", "username", "password", "full_name", "phone", "role_id", "is_active", "created_at"]]);
  
  // Departments
  sheets.Departments.getRange(1,1,1,4).setValues([["dept_id", "dept_name", "dept_code", "created_by"]]);
  
  // Trainers
  sheets.Trainers.getRange(1,1,1,7).setValues([["trainer_id", "name", "phone", "dept_id", "specialization", "status", "dept_name_auto"]]);
  
  // Halls
  sheets.Halls.getRange(1,1,1,6).setValues([["hall_id", "hall_name", "floor_number", "hall_type", "capacity", "status"]]);
  
  // Courses
  sheets.Courses.getRange(1,1,1,6).setValues([["course_id", "course_name", "dept_id", "price_per_level", "duration_levels", "dept_name_auto"]]);
  
  // Groups
  sheets.Groups.getRange(1,1,1,6).setValues([["group_id", "course_id", "group_name", "level_count", "start_date", "course_name_auto"]]);
  
  // Students
  sheets.Students.getRange(1,1,1,10).setValues([["student_id", "student_code", "name", "phone", "school", "age", "dept_id", "group_id", "created_at", "group_name_auto"]]);
  
  // Bookings
  sheets.Bookings.getRange(1,1,1,9).setValues([["booking_id", "hall_id", "trainer_id", "group_id", "day", "start_time", "end_time", "created_by", "conflict_check"]]);
  
  // Payments
  sheets.Payments.getRange(1,1,1,8).setValues([["payment_id", "student_id", "level_number", "amount_paid", "total_level_fee", "remaining_balance", "payment_date", "created_by"]]);
  
  // Levels
  sheets.Levels.getRange(1,1,1,6).setValues([["level_id", "student_id", "level_number", "level_fee", "status", "student_name_auto"]]);
  
  // AddOns
  sheets.AddOns.getRange(1,1,1,5).setValues([["addon_id", "name", "price", "course_id", "level_applicable"]]);
  
  // Format headers
  Object.values(sheets).forEach(sheet => {
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,sheet.getLastColumn()).setFontWeight("bold").setBackground("#1a2a6c").setFontColor("#ffffff");
  });
}

// --------------------------------------------
// 3. SEED DATA (CORRECTED - NO ROW COUNT MISMATCH)
// --------------------------------------------
function insertSeedData(sheets) {
  // --- Roles (3 columns) ---
  sheets.Roles.getRange(2,1,5,3).setValues([
    [1, "Admin", "Full system access"],
    [2, "Accountant", "Payments and students only"],
    [3, "Floor Manager", "Floor-specific data"],
    [4, "Booking Officer", "Schedule halls & trainers"],
    [5, "Viewer", "Read-only access"]
  ]);
  
  // --- Users (8 columns) ---
  sheets.Users.getRange(2,1,5,8).setValues([
    [1, "admin", "admin123", "Ahmed Mansour", "0100112233", 1, true, new Date()],
    [2, "accountant", "acc123", "Mona El Sayed", "0100445566", 2, true, new Date()],
    [3, "floor2", "floor123", "Khaled Hassan", "0100778899", 3, true, new Date()],
    [4, "booking", "book123", "Sara Waleed", "0100998877", 4, true, new Date()],
    [5, "viewer1", "view123", "Nour Ali", "0100554433", 5, true, new Date()]
  ]);
  
  // --- Departments (4 columns) ---
  sheets.Departments.getRange(2,1,4,4).setValues([
    [1, "Information Technology", "IT", 1],
    [2, "Languages", "LANG", 1],
    [3, "Business Administration", "BUS", 1],
    [4, "Graphic Design", "GD", 1]
  ]);
  
  // --- Trainers (7 columns) ---
  sheets.Trainers.getRange(2,1,5,7).setValues([
    [1, "Dr. Mohamed Fathy", "0123456789", 1, "Web Development", "Active", ""],
    [2, "Mr. Ali Zaki", "0123456790", 1, "Python", "Active", ""],
    [3, "Ms. Hend Sabry", "0123456791", 2, "English", "Active", ""],
    [4, "Dr. Nour Kamel", "0123456792", 3, "Project Management", "Active", ""],
    [5, "Mr. Tamer Roshdy", "0123456793", 4, "Photoshop", "Active", ""]
  ]);
  
  // --- Halls (6 columns) ---
  sheets.Halls.getRange(2,1,6,6).setValues([
    [1, "Hall A (Theory)", 1, "theory", 30, "Active"],
    [2, "Hall B (Practical)", 1, "practical", 20, "Active"],
    [3, "Hall C (Theory)", 2, "theory", 35, "Active"],
    [4, "Lab 1", 2, "practical", 15, "Active"],
    [5, "Conference Hall", 3, "theory", 50, "Active"],
    [6, "Meeting Room", 3, "theory", 10, "Active"]
  ]);
  
  // --- Courses (6 columns) ---
  sheets.Courses.getRange(2,1,5,6).setValues([
    [1, "Full Stack Web", 1, 500, 6, ""],
    [2, "Data Science", 1, 600, 8, ""],
    [3, "English Conversation", 2, 300, 4, ""],
    [4, "PMP Preparation", 3, 700, 2, ""],
    [5, "Motion Graphics", 4, 450, 5, ""]
  ]);
  
  // --- Groups (6 columns) ---
  sheets.Groups.getRange(2,1,5,6).setValues([
    [1, 1, "FSW-01", 6, new Date(2025,0,15), ""],
    [2, 1, "FSW-02", 6, new Date(2025,1,10), ""],
    [3, 3, "ENG-01", 4, new Date(2025,0,20), ""],
    [4, 4, "PMP-01", 2, new Date(2025,1,5), ""],
    [5, 2, "DS-01", 8, new Date(2025,0,25), ""]
  ]);
  
  // --- Students (10 columns) ---
  sheets.Students.getRange(2,1,8,10).setValues([
    [1, "IT-1001", "Youssef Ahmed", "0101112223", "Cairo University", 22, 1, 1, new Date(), ""],
    [2, "IT-1002", "Laila Mohamed", "0101112224", "Ain Shams", 21, 1, 2, new Date(), ""],
    [3, "LANG-2001", "Omar Khaled", "0101112225", "American University", 23, 2, 3, new Date(), ""],
    [4, "BUS-3001", "Salma Hany", "0101112226", "Cairo University", 25, 3, 4, new Date(), ""],
    [5, "IT-1003", "Mariam Tarek", "0101112227", "GUC", 20, 1, 1, new Date(), ""],
    [6, "GD-4001", "Hossam Gamal", "0101112228", "MSA", 24, 4, null, new Date(), ""],
    [7, "LANG-2002", "Nada Essam", "0101112229", "Ain Shams", 22, 2, 3, new Date(), ""],
    [8, "IT-1004", "Ziad Walid", "0101112230", "Cairo University", 21, 1, 2, new Date(), ""]
  ]);
  
  // --- Bookings (9 columns) ---
  sheets.Bookings.getRange(2,1,6,9).setValues([
    [1, 1, 1, 1, "Sunday", "10:00", "12:00", 1, "OK"],
    [2, 2, 2, 2, "Sunday", "13:00", "15:00", 1, "OK"],
    [3, 3, 3, 3, "Monday", "11:00", "13:00", 1, "OK"],
    [4, 1, 4, 4, "Tuesday", "09:00", "11:00", 1, "OK"],
    [5, 4, 5, 5, "Wednesday", "14:00", "16:00", 1, "OK"],
    [6, 5, 1, 1, "Thursday", "15:00", "17:00", 1, "OK"]
  ]);
  
  // --- Payments (8 columns) ---
  sheets.Payments.getRange(2,1,10,8).setValues([
    [1, 1, 1, 500, 500, 0, new Date(2025,0,10), 2],
    [2, 1, 2, 500, 500, 0, new Date(2025,1,10), 2],
    [3, 2, 1, 300, 500, 200, new Date(2025,0,15), 2],
    [4, 3, 1, 300, 300, 0, new Date(2025,0,20), 2],
    [5, 4, 1, 400, 700, 300, new Date(2025,0,5), 2],
    [6, 5, 1, 250, 500, 250, new Date(2025,0,12), 2],
    [7, 6, 1, 450, 450, 0, new Date(2025,0,18), 2],
    [8, 7, 1, 150, 300, 150, new Date(2025,0,22), 2],
    [9, 8, 1, 500, 500, 0, new Date(2025,0,8), 2],
    [10, 1, 3, 500, 500, 0, new Date(2025,2,10), 2]
  ]);
  
  // --- Levels (6 columns) ---
  sheets.Levels.getRange(2,1,10,6).setValues([
    [1, 1, 1, 500, "paid", ""],
    [2, 1, 2, 500, "paid", ""],
    [3, 1, 3, 500, "paid", ""],
    [4, 2, 1, 500, "partial", ""],
    [5, 3, 1, 300, "paid", ""],
    [6, 4, 1, 700, "unpaid", ""],
    [7, 5, 1, 500, "partial", ""],
    [8, 6, 1, 450, "paid", ""],
    [9, 7, 1, 300, "partial", ""],
    [10, 8, 1, 500, "paid", ""]
  ]);
  
  // --- AddOns (5 columns) ---
  sheets.AddOns.getRange(2,1,5,5).setValues([
    [1, "Course Book", 150, 1, 1],
    [2, "Online Exam Access", 100, 1, 1],
    [3, "Placement Test", 50, 2, 1],
    [4, "Certificate", 75, null, 0],
    [5, "Extra Practical Sessions", 200, 3, 1]
  ]);
}

// --------------------------------------------
// 4. SETUP RELATIONSHIPS (VLOOKUP Formulas)
// --------------------------------------------
function setupRelationships(sheets) {
  // Trainers - auto department name
  sheets.Trainers.getRange("G2").setFormula("=ARRAYFORMULA(IF(D2:D=\"\",, VLOOKUP(D2:D, Departments!A:D, 2, FALSE)))");
  
  // Courses - auto department name
  sheets.Courses.getRange("F2").setFormula("=ARRAYFORMULA(IF(C2:C=\"\",, VLOOKUP(C2:C, Departments!A:D, 2, FALSE)))");
  
  // Groups - auto course name
  sheets.Groups.getRange("F2").setFormula("=ARRAYFORMULA(IF(B2:B=\"\",, VLOOKUP(B2:B, Courses!A:F, 2, FALSE)))");
  
  // Students - auto group name
  sheets.Students.getRange("J2").setFormula("=ARRAYFORMULA(IF(H2:H=\"\",, VLOOKUP(H2:H, Groups!A:F, 3, FALSE)))");
  
  // Levels - auto student name
  sheets.Levels.getRange("F2").setFormula("=ARRAYFORMULA(IF(B2:B=\"\",, VLOOKUP(B2:B, Students!A:J, 3, FALSE)))");
  
  // Add dropdown validation
  addDropdownValidation(sheets);
  
  // Auto-resize columns
  Object.values(sheets).forEach(sheet => {
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });
}

// --------------------------------------------
// 5. DROPDOWN VALIDATION
// --------------------------------------------
function addDropdownValidation(sheets) {
  try {
    // Role dropdown in Users
    const rolesRange = sheets.Roles.getRange("B2:B" + sheets.Roles.getLastRow());
    const roleRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rolesRange, true)
      .build();
    sheets.Users.getRange("F2:F").setDataValidation(roleRule);
  } catch(e) {}
  
  try {
    // Status dropdowns
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Active", "Inactive"], true)
      .build();
    sheets.Trainers.getRange("F2:F").setDataValidation(statusRule);
    sheets.Halls.getRange("F2:F").setDataValidation(statusRule);
  } catch(e) {}
  
  try {
    // Level status
    const levelStatusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["paid", "partial", "unpaid"], true)
      .build();
    sheets.Levels.getRange("E2:E").setDataValidation(levelStatusRule);
  } catch(e) {}
  
  try {
    // Hall type
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["theory", "practical"], true)
      .build();
    sheets.Halls.getRange("D2:D").setDataValidation(typeRule);
  } catch(e) {}
}

// --------------------------------------------
// 6. CUSTOM MENU
// --------------------------------------------
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏫 Course System')
    .addItem('🔄 Full System Reset (Wipe & Rebuild)', 'setupEntireSystem')
    .addSeparator()
    .addItem('➕ Add Sample Student', 'addSampleStudent')
    .addItem('💰 Add Sample Payment', 'addSamplePayment')
    .addSeparator()
    .addItem('📊 Show Dashboard Summary', 'showDashboardSummary')
    .addToUi();
}

// --------------------------------------------
// 7. UTILITY FUNCTIONS
// --------------------------------------------
function addSampleStudent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentsSheet = ss.getSheetByName("Students");
  if (!studentsSheet) return;
  
  const lastRow = studentsSheet.getLastRow();
  const newId = lastRow;
  const deptCode = "IT";
  const randomNum = Math.floor(Math.random() * 9000 + 1000);
  const studentCode = `${deptCode}-${randomNum}`;
  
  studentsSheet.getRange(lastRow + 1, 1, 1, 10).setValues([[
    newId, studentCode, "New Student", "0100000000", "Test School", 20, 1, 1, new Date(), ""
  ]]);
  SpreadsheetApp.getUi().alert(`✅ Student added with Code: ${studentCode}`);
}

function addSamplePayment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentsSheet = ss.getSheetByName("Payments");
  if (!paymentsSheet) return;
  
  const lastRow = paymentsSheet.getLastRow();
  const newId = lastRow;
  
  paymentsSheet.getRange(lastRow + 1, 1, 1, 8).setValues([[
    newId, 1, 1, 100, 500, 400, new Date(), 2
  ]]);
  SpreadsheetApp.getUi().alert(`✅ Partial payment of 100 recorded for Student ID 1. Remaining: 400`);
}

function showDashboardSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const students = ss.getSheetByName("Students")?.getLastRow() - 1 || 0;
  const trainers = ss.getSheetByName("Trainers")?.getLastRow() - 1 || 0;
  const payments = ss.getSheetByName("Payments")?.getLastRow() - 1 || 0;
  
  SpreadsheetApp.getUi().alert(
    `📊 SYSTEM DASHBOARD\n\n` +
    `👨‍🎓 Students: ${students}\n` +
    `👨‍🏫 Trainers: ${trainers}\n` +
    `💰 Payments: ${payments}\n` +
    `✅ System ready for daily use.`
  );
}

// --------------------------------------------
// 8. ON OPEN TRIGGER
// --------------------------------------------
function onOpen() {
  createCustomMenu();
}

// ====================================================
// 9. WEB APP ENDPOINTS (doGet & doPost) - MODIFIED FOR CORS
// ====================================================

/**
 * doGet - Handles GET requests (for web app serving and API calls)
 */
function doGet(e) {
  // If there's a parameter 'action', handle API request (for CORS support)
  if (e && e.parameter && e.parameter.action) {
    const result = handleApiRequestGet(e.parameter);
    return createJsonResponse(result);
  }
  
  // Otherwise serve the HTML page
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Scientific Center - Course Management')
    .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3135/3135715.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * doPost - Handles POST requests (for API calls from GitHub frontend)
 */
function doPost(e) {
  try {
    // Parse the request data
    let requestData = {};
    if (e && e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      requestData = e.parameter;
    }
    
    const action = requestData.action;
    
    if (!action) {
      return createJsonResponse({ success: false, message: "Missing 'action' parameter" }, 400);
    }
    
    const result = handleApiRequestPost(action, requestData);
    return createJsonResponse(result);
    
  } catch (error) {
    return createJsonResponse({ 
      success: false, 
      message: "Server error: " + error.toString() 
    }, 500);
  }
}

/**
 * Handle GET API requests (for CORS support from GitHub Pages)
 */
function handleApiRequestGet(params) {
  const action = params.action;
  
  switch (action) {
    case 'login':
      return verifyLogin(params.username, params.password);
    
    case 'getAllStudents':
      return { success: true, data: getAllStudents() };
    
    case 'getAllPayments':
      return { success: true, data: getAllPayments() };
    
    case 'getFinancialSummary':
      return { success: true, data: getFinancialSummary() };
    
    case 'getAllTrainers':
      return { success: true, data: getAllTrainers() };
    
    case 'getAllHalls':
      return { success: true, data: getAllHalls() };
    
    case 'getAllGroups':
      return { success: true, data: getAllGroups() };
    
    case 'getAllBookings':
      return { success: true, data: getAllBookings() };
    
    case 'getDashboardStats':
      return { success: true, data: getDashboardStats() };
    
    default:
      return { success: false, message: "Unknown action: " + action };
  }
}

/**
 * Handle POST API requests
 */
function handleApiRequestPost(action, params) {
  switch (action) {
    case 'login':
      return verifyLogin(params.username, params.password);
      
    case 'getAllUsers':
      return { success: true, data: getAllUsers() };
      
    case 'saveUser':
      return saveUser(params.data);
      
    case 'deleteUser':
      return deleteUser(params.id);
    
    case 'getAllStudents':
      return { success: true, data: getAllStudents() };
    
    case 'getStudentById':
      return { success: true, data: getStudentById(params.studentId) };
    
    case 'saveStudent':
      return saveStudent(params);
    
    case 'deleteStudent':
      return deleteStudent(params.studentId);
    
    case 'getAllPayments':
      return { success: true, data: getAllPayments() };
    
    case 'savePayment':
      return savePayment(params);
    
    case 'deletePayment':
      return deletePayment(params.paymentId);
    
    case 'getFinancialSummary':
      return { success: true, data: getFinancialSummary() };
    
    case 'getAllTrainers':
      return { success: true, data: getAllTrainers() };
      
    case 'saveTrainer':
      return saveTrainer(params.data);
      
    case 'deleteTrainer':
      return deleteTrainer(params.id);
    
    case 'getAllHalls':
      return { success: true, data: getAllHalls() };
    
    case 'getAllGroups':
      return { success: true, data: getAllGroups() };
    
    case 'getAllBookings':
      return { success: true, data: getAllBookings() };
    
    case 'saveBooking':
      return saveBooking(params);
    
    case 'deleteBooking':
      return deleteBooking(params.bookingId);
    
    case 'getDashboardStats':
      return { success: true, data: getDashboardStats() };
    
    default:
      return { success: false, message: "Unknown action: " + action };
  }
}

/**
 * Create JSON response with proper headers for CORS
 */
function createJsonResponse(data, statusCode = 200) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ====================================================
// 10. API FUNCTIONS (DATABASE OPERATIONS)
// ====================================================

// ---------- AUTH ----------
function verifyLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "System not initialized" };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username && data[i][2] === password) {
      return {
        success: true,
        user: {
          id: data[i][0],
          username: data[i][1],
          fullName: data[i][3],
          phone: data[i][4],
          roleId: data[i][5],
          isActive: data[i][6]
        }
      };
    }
  }
  return { success: false, message: "Invalid username or password" };
}

// ---------- USERS ----------
function getAllUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    users.push({
      id: data[i][0],
      username: data[i][1],
      password: data[i][2],
      fullName: data[i][3],
      phone: data[i][4],
      roleId: data[i][5],
      isActive: data[i][6],
      createdAt: data[i][7]
    });
  }
  return users;
}

function saveUser(userData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "Sheet not found" };
  
  const data = sheet.getDataRange().getValues();
  
  // Check for unique username
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userData.username && data[i][0] != userData.id) {
      return { success: false, message: "اسم المستخدم موجود مسبقاً!" };
    }
  }
  
  if (userData.id) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == userData.id) {
        sheet.getRange(i+1, 2).setValue(userData.username);
        if (userData.password) sheet.getRange(i+1, 3).setValue(userData.password);
        sheet.getRange(i+1, 4).setValue(userData.fullName);
        sheet.getRange(i+1, 5).setValue(userData.phone);
        sheet.getRange(i+1, 6).setValue(parseInt(userData.roleId));
        sheet.getRange(i+1, 7).setValue(userData.isActive);
        return { success: true, message: "تم تحديث المستخدم بنجاح" };
      }
    }
  } else {
    const lastRow = sheet.getLastRow();
    const newId = lastRow;
    
    sheet.getRange(lastRow + 1, 1).setValue(newId);
    sheet.getRange(lastRow + 1, 2).setValue(userData.username);
    sheet.getRange(lastRow + 1, 3).setValue(userData.password);
    sheet.getRange(lastRow + 1, 4).setValue(userData.fullName);
    sheet.getRange(lastRow + 1, 5).setValue(userData.phone);
    sheet.getRange(lastRow + 1, 6).setValue(parseInt(userData.roleId));
    sheet.getRange(lastRow + 1, 7).setValue(userData.isActive);
    sheet.getRange(lastRow + 1, 8).setValue(new Date());
    
    return { success: true, message: "تم إضافة المستخدم بنجاح" };
  }
  return { success: false, message: "حدث خطأ غير معروف" };
}

function deleteUser(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "Sheet not found" };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == userId) {
      sheet.deleteRow(i+1);
      return { success: true, message: "تم حذف المستخدم بنجاح" };
    }
  }
  return { success: false, message: "المستخدم غير موجود" };
}

// ---------- STUDENTS ----------
function getAllStudents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < data.length; i++) {
    students.push({
      id: data[i][0],
      code: data[i][1],
      name: data[i][2],
      phone: data[i][3],
      school: data[i][4],
      age: data[i][5],
      deptId: data[i][6],
      groupId: data[i][7],
      createdAt: data[i][8],
      groupName: data[i][9]
    });
  }
  return students;
}

function getStudentById(studentId) {
  const students = getAllStudents();
  return students.find(s => s.id == studentId) || null;
}

function saveStudent(studentData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  if (!sheet) return { success: false, message: "Sheet not found" };
  
  const data = sheet.getDataRange().getValues();
  
  if (studentData.id) {
    // Update existing student
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == studentData.id) {
        let studentCode = data[i][1];
        if (!studentCode) {
          const deptCode = {1: "IT", 2: "LANG", 3: "BUS", 4: "GD"}[studentData.deptId] || "STD";
          const randomNum = Math.floor(Math.random() * 9000 + 1000);
          studentCode = `${deptCode}-${randomNum}`;
        }
        
        sheet.getRange(i+1, 2).setValue(studentCode);
        sheet.getRange(i+1, 3).setValue(studentData.name);
        sheet.getRange(i+1, 4).setValue(studentData.phone);
        sheet.getRange(i+1, 5).setValue(studentData.school);
        sheet.getRange(i+1, 6).setValue(studentData.age);
        sheet.getRange(i+1, 7).setValue(parseInt(studentData.deptId));
        sheet.getRange(i+1, 8).setValue(studentData.groupId ? parseInt(studentData.groupId) : null);
        
        return { success: true, message: "تم تحديث بيانات الطالب بنجاح" };
      }
    }
  } else {
    // Add new student
    const lastRow = sheet.getLastRow();
    const newId = lastRow;
    
    const deptCode = {1: "IT", 2: "LANG", 3: "BUS", 4: "GD"}[studentData.deptId] || "STD";
    const randomNum = Math.floor(Math.random() * 9000 + 1000);
    const studentCode = `${deptCode}-${randomNum}`;
    
    sheet.getRange(lastRow + 1, 1).setValue(newId);
    sheet.getRange(lastRow + 1, 2).setValue(studentCode);
    sheet.getRange(lastRow + 1, 3).setValue(studentData.name);
    sheet.getRange(lastRow + 1, 4).setValue(studentData.phone);
    sheet.getRange(lastRow + 1, 5).setValue(studentData.school);
    sheet.getRange(lastRow + 1, 6).setValue(studentData.age);
    sheet.getRange(lastRow + 1, 7).setValue(parseInt(studentData.deptId));
    sheet.getRange(lastRow + 1, 8).setValue(studentData.groupId ? parseInt(studentData.groupId) : null);
    sheet.getRange(lastRow + 1, 9).setValue(new Date());
    
    return { success: true, message: `تم إضافة الطالب بنجاح - الكود: ${studentCode}` };
  }
  
  return { success: false, message: "حدث خطأ" };
}

function deleteStudent(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  if (!sheet) return { success: false, message: "Sheet not found" };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == studentId) {
      sheet.deleteRow(i+1);
      return { success: true, message: "تم حذف الطالب بنجاح" };
    }
  }
  
  return { success: false, message: "الطالب غير موجود" };
}

// ---------- PAYMENTS ----------
function getAllPayments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Payments");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const payments = [];
  for (let i = 1; i < data.length; i++) {
    payments.push({
      id: data[i][0],
      studentId: data[i][1],
      studentName: getStudentNameById(data[i][1]),
      levelNumber: data[i][2],
      amountPaid: data[i][3],
      totalFee: data[i][4],
      remainingBalance: data[i][5],
      paymentDate: data[i][6],
      createdBy: data[i][7]
    });
  }
  return payments;
}

function getStudentNameById(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == studentId) return data[i][2];
  }
  return "";
}

function savePayment(paymentData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Payments");
  if (!sheet) return { success: false, message: "Sheet not found" };
  
  const lastRow = sheet.getLastRow();
  const newId = lastRow;
  
  const remainingBalance = paymentData.totalLevelFee - paymentData.amountPaid;
  
  sheet.getRange(lastRow + 1, 1).setValue(newId);
  sheet.getRange(lastRow + 1, 2).setValue(parseInt(paymentData.studentId));
  sheet.getRange(lastRow + 1, 3).setValue(paymentData.levelNumber);
  sheet.getRange(lastRow + 1, 4).setValue(paymentData.amountPaid);
  sheet.getRange(lastRow + 1, 5).setValue(paymentData.totalLevelFee);
  sheet.getRange(lastRow + 1, 6).setValue(remainingBalance);
  sheet.getRange(lastRow + 1, 7).setValue(paymentData.paymentDate);
  sheet.getRange(lastRow + 1, 8).setValue(paymentData.createdBy);
  
  // Update levels table
  updateLevelsTable(parseInt(paymentData.studentId), paymentData.levelNumber, paymentData.totalLevelFee, remainingBalance);
  
  return { success: true, message: "تم تسجيل الدفعة بنجاح" };
}

function updateLevelsTable(studentId, levelNumber, levelFee, remainingBalance) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const levelsSheet = ss.getSheetByName("Levels");
  if (!levelsSheet) return;
  
  const data = levelsSheet.getDataRange().getValues();
  let found = false;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == studentId && data[i][2] == levelNumber) {
      const status = remainingBalance === 0 ? "paid" : "partial";
      levelsSheet.getRange(i+1, 4).setValue(levelFee);
      levelsSheet.getRange(i+1, 5).setValue(status);
      found = true;
      break;
    }
  }
  
  if (!found) {
    const lastRow = levelsSheet.getLastRow();
    const status = remainingBalance === 0 ? "paid" : "partial";
    levelsSheet.getRange(lastRow + 1, 1).setValue(lastRow);
    levelsSheet.getRange(lastRow + 1, 2).setValue(studentId);
    levelsSheet.getRange(lastRow + 1, 3).setValue(levelNumber);
    levelsSheet.getRange(lastRow + 1, 4).setValue(levelFee);
    levelsSheet.getRange(lastRow + 1, 5).setValue(status);
  }
}

function deletePayment(paymentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Payments");
  if (!sheet) return { success: false, message: "Sheet not found" };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == paymentId) {
      sheet.deleteRow(i+1);
      return { success: true, message: "تم حذف الدفعة بنجاح" };
    }
  }
  return { success: false, message: "الدفعة غير موجودة" };
}

function getFinancialSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentsSheet = ss.getSheetByName("Payments");
  if (!paymentsSheet) return { totalPaid: 0, totalRemaining: 0, studentCount: 0 };
  
  const payments = paymentsSheet.getDataRange().getValues();
  let totalPaid = 0;
  let totalRemaining = 0;
  for (let i = 1; i < payments.length; i++) {
    totalPaid += payments[i][3] || 0;
    totalRemaining += payments[i][5] || 0;
  }
  
  const studentsSheet = ss.getSheetByName("Students");
  const studentCount = studentsSheet ? studentsSheet.getLastRow() - 1 : 0;
  
  return {
    totalPaid: totalPaid,
    totalRemaining: totalRemaining,
    studentCount: studentCount,
    trainerCount: ss.getSheetByName("Trainers")?.getLastRow() - 1 || 0,
    bookingCount: ss.getSheetByName("Bookings")?.getLastRow() - 1 || 0
  };
}

// ---------- TRAINERS ----------
function getAllTrainers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Trainers");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const trainers = [];
  for (let i = 1; i < data.length; i++) {
    trainers.push({
      id: data[i][0],
      name: data[i][1],
      phone: data[i][2],
      deptId: data[i][3],
      specialization: data[i][4],
      status: data[i][5],
      deptName: data[i][6]
    });
  }
  return trainers;
}

function saveTrainer(trainerData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Trainers");
  if (!sheet) return { success: false, message: "Sheet not found" };
  
  const data = sheet.getDataRange().getValues();
  
  if (trainerData.id) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == trainerData.id) {
        sheet.getRange(i+1, 2).setValue(trainerData.name);
        sheet.getRange(i+1, 3).setValue(trainerData.phone);
        sheet.getRange(i+1, 4).setValue(parseInt(trainerData.deptId));
        sheet.getRange(i+1, 5).setValue(trainerData.specialization);
        sheet.getRange(i+1, 6).setValue(trainerData.status || "Active");
        return { success: true, message: "تم تحديث بيانات المدرب بنجاح" };
      }
    }
  } else {
    const lastRow = sheet.getLastRow();
    const newId = lastRow;
    
    sheet.getRange(lastRow + 1, 1).setValue(newId);
    sheet.getRange(lastRow + 1, 2).setValue(trainerData.name);
    sheet.getRange(lastRow + 1, 3).setValue(trainerData.phone);
    sheet.getRange(lastRow + 1, 4).setValue(parseInt(trainerData.deptId));
    sheet.getRange(lastRow + 1, 5).setValue(trainerData.specialization);
    sheet.getRange(lastRow + 1, 6).setValue(trainerData.status || "Active");
    
    return { success: true, message: "تم إضافة المدرب بنجاح" };
  }
  
  return { success: false, message: "حدث خطأ" };
}

function deleteTrainer(trainerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Trainers");
  if (!sheet) return { success: false, message: "Sheet not found" };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == trainerId) {
      sheet.deleteRow(i+1);
      return { success: true, message: "تم حذف المدرب بنجاح" };
    }
  }
  
  return { success: false, message: "المدرب غير موجود" };
}

// ---------- HALLS ----------
function getAllHalls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Halls");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const halls = [];
  for (let i = 1; i < data.length; i++) {
    halls.push({
      id: data[i][0],
      name: data[i][1],
      floorNumber: data[i][2],
      type: data[i][3],
      capacity: data[i][4],
      status: data[i][5]
    });
  }
  return halls;
}

// ---------- GROUPS ----------
function getAllGroups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Groups");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const groups = [];
  for (let i = 1; i < data.length; i++) {
    groups.push({
      id: data[i][0],
      name: data[i][2],
      courseId: data[i][1],
      levelCount: data[i][3],
      startDate: data[i][4],
      courseName: data[i][5]
    });
  }
  return groups;
}

// ---------- BOOKINGS ----------
function getAllBookings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Bookings");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const bookings = [];
  for (let i = 1; i < data.length; i++) {
    bookings.push({
      id: data[i][0],
      hallId: data[i][1],
      trainerId: data[i][2],
      groupId: data[i][3],
      day: data[i][4],
      startTime: data[i][5],
      endTime: data[i][6],
      createdBy: data[i][7],
      conflict: data[i][8]
    });
  }
  return bookings;
}

function saveBooking(bookingData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Bookings");
  if (!sheet) return { success: false, message: "Sheet not found" };
  
  // Check for conflicts
  const allBookings = getAllBookings();
  const conflict = checkBookingConflict(bookingData, allBookings);
  
  if (bookingData.id) {
    // Update existing
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == bookingData.id) {
        sheet.getRange(i+1, 2).setValue(parseInt(bookingData.hallId));
        sheet.getRange(i+1, 3).setValue(parseInt(bookingData.trainerId));
        sheet.getRange(i+1, 4).setValue(parseInt(bookingData.groupId));
        sheet.getRange(i+1, 5).setValue(bookingData.day);
        sheet.getRange(i+1, 6).setValue(bookingData.startTime);
        sheet.getRange(i+1, 7).setValue(bookingData.endTime);
        sheet.getRange(i+1, 9).setValue(conflict ? "CONFLICT" : "OK");
        return { success: true, message: conflict ? "تم الحفظ مع وجود تعارض!" : "تم تحديث الحجز بنجاح" };
      }
    }
  } else {
    // Add new
    const lastRow = sheet.getLastRow();
    const newId = lastRow;
    sheet.getRange(lastRow + 1, 1).setValue(newId);
    sheet.getRange(lastRow + 1, 2).setValue(parseInt(bookingData.hallId));
    sheet.getRange(lastRow + 1, 3).setValue(parseInt(bookingData.trainerId));
    sheet.getRange(lastRow + 1, 4).setValue(parseInt(bookingData.groupId));
    sheet.getRange(lastRow + 1, 5).setValue(bookingData.day);
    sheet.getRange(lastRow + 1, 6).setValue(bookingData.startTime);
    sheet.getRange(lastRow + 1, 7).setValue(bookingData.endTime);
    sheet.getRange(lastRow + 1, 8).setValue(bookingData.createdBy || 1);
    sheet.getRange(lastRow + 1, 9).setValue(conflict ? "CONFLICT" : "OK");
    return { success: true, message: conflict ? "تم الحجز مع وجود تعارض!" : "تم الحجز بنجاح" };
  }
  
  return { success: false, message: "حدث خطأ" };
}

function checkBookingConflict(newBooking, existingBookings) {
  for (let booking of existingBookings) {
    if (booking.id == newBooking.id) continue;
    if (booking.day === newBooking.day) {
      // Check hall conflict
      if (booking.hallId == newBooking.hallId) {
        if ((newBooking.startTime >= booking.startTime && newBooking.startTime < booking.endTime) ||
            (newBooking.endTime > booking.startTime && newBooking.endTime <= booking.endTime) ||
            (newBooking.startTime <= booking.startTime && newBooking.endTime >= booking.endTime)) {
          return true;
        }
      }
      // Check trainer conflict
      if (booking.trainerId == newBooking.trainerId) {
        if ((newBooking.startTime >= booking.startTime && newBooking.startTime < booking.endTime) ||
            (newBooking.endTime > booking.startTime && newBooking.endTime <= booking.endTime) ||
            (newBooking.startTime <= booking.startTime && newBooking.endTime >= booking.endTime)) {
          return true;
        }
      }
    }
  }
  return false;
}

function deleteBooking(bookingId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Bookings");
  if (!sheet) return { success: false, message: "Sheet not found" };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == bookingId) {
      sheet.deleteRow(i+1);
      return { success: true, message: "تم حذف الحجز بنجاح" };
    }
  }
  return { success: false, message: "الحجز غير موجود" };
}

// ---------- DASHBOARD ----------
function getDashboardStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  return {
    studentCount: ss.getSheetByName("Students")?.getLastRow() - 1 || 0,
    trainerCount: ss.getSheetByName("Trainers")?.getLastRow() - 1 || 0,
    bookingCount: ss.getSheetByName("Bookings")?.getLastRow() - 1 || 0,
    paymentCount: ss.getSheetByName("Payments")?.getLastRow() - 1 || 0,
    hallCount: ss.getSheetByName("Halls")?.getLastRow() - 1 || 0,
    groupCount: ss.getSheetByName("Groups")?.getLastRow() - 1 || 0
  };
}