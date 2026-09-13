// ===== EQUIPMENT MAINTENANCE MANAGEMENT SYSTEM =====
// Main Google Apps Script Backend
// ================================

// Spreadsheet Configuration
const SPREADSHEET_ID = PropertiesService.getUserProperties().getProperty('SPREADSHEET_ID') || '';
const SHEET_NAMES = {
  EQUIPMENT: 'Equipment',
  SPARE_PARTS: 'SpareParts',
  APPROVERS: 'Approvers',
  REPAIR_REQUESTS: 'RepairRequests',
  REPAIR_RECORDS: 'RepairRecords',
  PM_PLANS: 'PMPlans',
  PM_RECORDS: 'PMRecords',
  SIGNATURES: 'Signatures',
  EVALUATIONS: 'Evaluations',
  SPARE_PARTS_USAGE: 'SparePartsUsage'
};

let ss = null;

/**
 * Initialize Spreadsheet
 */
function initializeSpreadsheet() {
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = ss.getId();
    PropertiesService.getUserProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
    
    // Create all required sheets if they don't exist
    Object.values(SHEET_NAMES).forEach(sheetName => {
      if (!ss.getSheetByName(sheetName)) {
        ss.insertSheet(sheetName);
      }
    });
    
    // Initialize sheet headers
    initializeSheetHeaders();
    Logger.log('Spreadsheet initialized successfully');
  } catch (e) {
    Logger.log('Error initializing spreadsheet: ' + e.toString());
  }
}

/**
 * Initialize sheet headers
 */
function initializeSheetHeaders() {
  // Equipment Sheet Headers
  const equipmentHeaders = ['ID', 'ຊື່ເຄື່ອງ', 'ຍີ່ຫໍ້', 'ຮຸ່ນ', 'ໝວດອຸປະກອນ', 'ຜູ້ໃຊ້ງານ', 'ສະຖານທີ່ຕິດຕັ້ງ', 'ວັນທີລົງທະບຽນ', 'ຮູບພາບ', 'ວັນທີປະກັນເລີ່ມ', 'ວັນທີປະກັນສິ້ນ', 'ທຶນ', 'ລາຍການ', 'ສະຖານະ'];
  setSheetHeaders(SHEET_NAMES.EQUIPMENT, equipmentHeaders);
  
  // Spare Parts Sheet Headers
  const sparePartsHeaders = ['ID', 'ຊື່ອາໄຫຼ່', 'ປະເພດ', 'ຈຳນວນຕົ້ນຕໍ', 'ຈຳນວນຄົງເຫຼືອ', 'ລະດັບເຕືອນ', 'ລາຄາຕໍ່ຫົວໜ່ວຍ', 'ວັນທີສ້າງ'];
  setSheetHeaders(SHEET_NAMES.SPARE_PARTS, sparePartsHeaders);
  
  // Approvers Sheet Headers
  const approversHeaders = ['ໜ່ວຍງານ/ແຜນກ', 'ອີເມວຜູ້ອະນຸມັດ', 'ຊື່ຜູ້ອະນຸມັດ', 'ວັນທີຕັ້ງຄ່າ'];
  setSheetHeaders(SHEET_NAMES.APPROVERS, approversHeaders);
  
  // Repair Requests Sheet Headers
  const repairRequestsHeaders = ['ID', 'ອຸປະກອນID', 'ຊື່ອຸປະກອນ', 'ປະເພດວຽກ', 'ລະດັບຄວາມດ່ວນ', 'ລາຍລະອຽດບັນຫາ', 'ຮູບພາບ', 'ຜູ້ແຈ້ງ', 'ວັນທີແຈ້ງ', 'ສະຖານະ', 'ໜ່ວຍງານ/ແຜນກ', 'ເຫດຜົນປະຕິເສດ'];
  setSheetHeaders(SHEET_NAMES.REPAIR_REQUESTS, repairRequestsHeaders);
  
  // Repair Records Sheet Headers
  const repairRecordsHeaders = ['ID', 'ໄອດີການແຈ້ງສ້ອມ', 'ສາເຫດ', 'ວິທີການແກ້ໄຂ', 'ອາໄຼ່ທີ່ໃຊ້', 'ຈຳນວນອາໄຼ່', 'ຄ່າອາໄຼ່', 'ຄ່າແຮງງານ', 'ຄ່າອື່ນໆ', 'ຄ່າໃຊ້ຈ່າຍລວມ', 'ວັນທີສ້ອມ', 'ຊ່າງສ້ອມ', 'ສະຖານະ'];
  setSheetHeaders(SHEET_NAMES.REPAIR_RECORDS, repairRecordsHeaders);
  
  // PM Plans Sheet Headers
  const pmPlansHeaders = ['ID', 'ອຸປະກອນID', 'ຊື່ອຸປະກອນ', 'ປະເພດ PM', 'ຄວາມຖີ່', 'ວັນທີເລີ່ມ', 'ວັນທີ PM ຄັ້ງຕໍ່ໄປ', 'ລາຍການກວດສອບ', 'ສະຖານະ'];
  setSheetHeaders(SHEET_NAMES.PM_PLANS, pmPlansHeaders);
  
  // PM Records Sheet Headers
  const pmRecordsHeaders = ['ID', 'ແຜນPMID', 'ວັນທີທຳ', 'ຜົນການກວດສອບ', 'ວຽກທີ່ຕ້ອງເຮັດ', 'ຊ່າງລາຍລະອຽກ', 'ໝາຍເຫດ'];
  setSheetHeaders(SHEET_NAMES.PM_RECORDS, pmRecordsHeaders);
  
  // Signatures Sheet Headers
  const signaturesHeaders = ['ID', 'ໄອດີການແຈ້ງສ້ອມ', 'ລາຍເຊັນ', 'ວັນທີເຊັນ', 'ຊື່ຜູ້ເຊັນ'];
  setSheetHeaders(SHEET_NAMES.SIGNATURES, signaturesHeaders);
  
  // Evaluations Sheet Headers
  const evaluationsHeaders = ['ID', 'ໄອດີການແຈ້ງສ້ອມ', 'ຫົວຂໍ້ທີ 1', 'ຫົວຂໍ້ທີ 2', 'ຫົວຂໍ້ທີ 3', 'ຫົວຂໍ້ທີ 4', 'ຫົວຂໍ້ທີ 5', 'ຄວາມຄິດເຫັນ', 'ວັນທີປະເມີນ'];
  setSheetHeaders(SHEET_NAMES.EVALUATIONS, evaluationsHeaders);
  
  // Spare Parts Usage Sheet Headers
  const sparePartsUsageHeaders = ['ID', 'ອາໄຼ່ID', 'ໄອດີການແຈ້ງສ້ອມ', 'ຈຳນວນທີ່ໃຊ້', 'ວັນທີໃຊ້', 'ລາຄາລວມ'];
  setSheetHeaders(SHEET_NAMES.SPARE_PARTS_USAGE, sparePartsUsageHeaders);
}

/**
 * Set sheet headers
 */
function setSheetHeaders(sheetName, headers) {
  const sheet = ss.getSheetByName(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
}

// ===== EQUIPMENT MANAGEMENT =====

/**
 * Add new equipment
 */
function addEquipment(equipmentData) {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.EQUIPMENT);
    const id = generateId('EQP');
    const row = [
      id,
      equipmentData.name,
      equipmentData.brand,
      equipmentData.model,
      equipmentData.category,
      equipmentData.user,
      equipmentData.location,
      new Date(),
      equipmentData.imageUrl || '',
      equipmentData.warrantyStart || '',
      equipmentData.warrantyEnd || '',
      equipmentData.capital || '',
      equipmentData.remark || '',
      'ໃຊ້ຢູ່'
    ];
    sheet.appendRow(row);
    Logger.log('Equipment added: ' + id);
    return {success: true, id: id};
  } catch (e) {
    Logger.log('Error adding equipment: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

/**
 * Get all equipment
 */
function getAllEquipment() {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.EQUIPMENT);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const equipmentList = [];
    for (let i = 1; i < data.length; i++) {
      equipmentList.push({
        id: data[i][0],
        name: data[i][1],
        brand: data[i][2],
        model: data[i][3],
        category: data[i][4],
        user: data[i][5],
        location: data[i][6],
        dateAdded: data[i][7],
        imageUrl: data[i][8],
        warrantyStart: data[i][9],
        warrantyEnd: data[i][10],
        capital: data[i][11],
        remark: data[i][12],
        status: data[i][13]
      });
    }
    return equipmentList;
  } catch (e) {
    Logger.log('Error getting equipment: ' + e.toString());
    return [];
  }
}

/**
 * Get equipment by ID
 */
function getEquipmentById(equipmentId) {
  try {
    const equipment = getAllEquipment();
    return equipment.find(eq => eq.id === equipmentId) || null;
  } catch (e) {
    Logger.log('Error getting equipment by ID: ' + e.toString());
    return null;
  }
}

/**
 * Update equipment
 */
function updateEquipment(equipmentId, equipmentData) {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.EQUIPMENT);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === equipmentId) {
        sheet.getRange(i + 1, 2, 1, 12).setValues([[
          equipmentData.name || data[i][1],
          equipmentData.brand || data[i][2],
          equipmentData.model || data[i][3],
          equipmentData.category || data[i][4],
          equipmentData.user || data[i][5],
          equipmentData.location || data[i][6],
          data[i][7],
          equipmentData.imageUrl || data[i][8],
          equipmentData.warrantyStart || data[i][9],
          equipmentData.warrantyEnd || data[i][10],
          equipmentData.capital || data[i][11],
          equipmentData.remark || data[i][12]
        ]]);
        return {success: true};
      }
    }
    return {success: false, error: 'Equipment not found'};
  } catch (e) {
    Logger.log('Error updating equipment: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

// ===== SPARE PARTS MANAGEMENT =====

/**
 * Add new spare part
 */
function addSparePart(sparePartData) {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.SPARE_PARTS);
    const id = generateId('SP');
    const row = [
      id,
      sparePartData.name,
      sparePartData.type,
      sparePartData.quantity || 0,
      sparePartData.quantity || 0,
      sparePartData.warningLevel || 5,
      sparePartData.price || 0,
      new Date()
    ];
    sheet.appendRow(row);
    return {success: true, id: id};
  } catch (e) {
    Logger.log('Error adding spare part: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

/**
 * Get all spare parts
 */
function getAllSpareParts() {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.SPARE_PARTS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const sparePartsList = [];
    for (let i = 1; i < data.length; i++) {
      sparePartsList.push({
        id: data[i][0],
        name: data[i][1],
        type: data[i][2],
        initialQuantity: data[i][3],
        remainingQuantity: data[i][4],
        warningLevel: data[i][5],
        price: data[i][6],
        dateCreated: data[i][7]
      });
    }
    return sparePartsList;
  } catch (e) {
    Logger.log('Error getting spare parts: ' + e.toString());
    return [];
  }
}

/**
 * Update spare parts quantity
 */
function updateSparePartQuantity(sparePartId, quantityUsed) {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.SPARE_PARTS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sparePartId) {
        const newQuantity = data[i][4] - quantityUsed;
        sheet.getRange(i + 1, 5).setValue(newQuantity);
        return {success: true, newQuantity: newQuantity};
      }
    }
    return {success: false, error: 'Spare part not found'};
  } catch (e) {
    Logger.log('Error updating spare part quantity: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

// ===== REPAIR REQUEST MANAGEMENT =====

/**
 * Create repair request
 */
function createRepairRequest(requestData) {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.REPAIR_REQUESTS);
    const id = generateId('RR');
    
    // Get department/section from user
    const approverInfo = getApproverByEquipment(requestData.equipmentId);
    
    const row = [
      id,
      requestData.equipmentId || '',
      requestData.equipmentName || '',
      requestData.workType || '',
      requestData.urgency || 'ធម្មតា',
      requestData.description || '',
      requestData.imageUrl || '',
      Session.getActiveUser().getEmail(),
      new Date(),
      'ແຈ້ງສ້ອມ',
      approverInfo.department || '',
      ''
    ];
    sheet.appendRow(row);
    
    // Send email to approver
    if (approverInfo.email) {
      sendApprovalNotification(id, approverInfo.email, requestData);
    }
    
    return {success: true, id: id};
  } catch (e) {
    Logger.log('Error creating repair request: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

/**
 * Get all repair requests
 */
function getAllRepairRequests() {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.REPAIR_REQUESTS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const requestList = [];
    for (let i = 1; i < data.length; i++) {
      requestList.push({
        id: data[i][0],
        equipmentId: data[i][1],
        equipmentName: data[i][2],
        workType: data[i][3],
        urgency: data[i][4],
        description: data[i][5],
        imageUrl: data[i][6],
        reportedBy: data[i][7],
        dateReported: data[i][8],
        status: data[i][9],
        department: data[i][10],
        rejectionReason: data[i][11]
      });
    }
    return requestList;
  } catch (e) {
    Logger.log('Error getting repair requests: ' + e.toString());
    return [];
  }
}

/**
 * Approve or reject repair request
 */
function approveRepairRequest(requestId, approved, reason = '') {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.REPAIR_REQUESTS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === requestId) {
        const newStatus = approved ? 'ລໍຖ້າກວດສອບ' : 'ປະຕິເສດ';
        sheet.getRange(i + 1, 10).setValue(newStatus);
        
        if (!approved) {
          sheet.getRange(i + 1, 12).setValue(reason);
        }
        
        // Send notification to technician or reporter
        const email = data[i][7];
        if (approved) {
          sendTechnicianNotification(requestId, email);
        } else {
          sendRejectionNotification(requestId, email, reason);
        }
        
        return {success: true};
      }
    }
    return {success: false, error: 'Request not found'};
  } catch (e) {
    Logger.log('Error approving/rejecting request: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

// ===== REPAIR RECORD MANAGEMENT =====

/**
 * Add repair record
 */
function addRepairRecord(repairData) {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.REPAIR_RECORDS);
    const id = generateId('RP');
    
    const totalCost = (repairData.spareCost || 0) + (repairData.laborCost || 0) + (repairData.otherCost || 0);
    
    const row = [
      id,
      repairData.requestId,
      repairData.cause || '',
      repairData.solution || '',
      repairData.sparePart || '',
      repairData.quantity || 0,
      repairData.spareCost || 0,
      repairData.laborCost || 0,
      repairData.otherCost || 0,
      totalCost,
      new Date(),
      Session.getActiveUser().getEmail(),
      'ສ້ອມສຳເລັດ'
    ];
    sheet.appendRow(row);
    
    // Update spare part quantity if used
    if (repairData.sparePartId && repairData.quantity) {
      updateSparePartQuantity(repairData.sparePartId, repairData.quantity);
    }
    
    // Update repair request status
    updateRepairRequestStatus(repairData.requestId, 'ສ້ອມສຳເລັດ');
    
    return {success: true, id: id};
  } catch (e) {
    Logger.log('Error adding repair record: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

/**
 * Update repair request status
 */
function updateRepairRequestStatus(requestId, status) {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.REPAIR_REQUESTS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === requestId) {
        sheet.getRange(i + 1, 10).setValue(status);
        return {success: true};
      }
    }
    return {success: false, error: 'Request not found'};
  } catch (e) {
    Logger.log('Error updating repair request status: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

// ===== PM PLAN MANAGEMENT =====

/**
 * Create PM plan
 */
function createPMPlan(pmData) {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.PM_PLANS);
    const id = generateId('PM');
    
    const row = [
      id,
      pmData.equipmentId,
      pmData.equipmentName,
      pmData.type,
      pmData.frequency,
      new Date(),
      calculateNextPMDate(new Date(), pmData.frequency),
      pmData.checklist || '',
      'ກຳລັງລໍຖ້າ'
    ];
    sheet.appendRow(row);
    
    return {success: true, id: id};
  } catch (e) {
    Logger.log('Error creating PM plan: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

/**
 * Calculate next PM date
 */
function calculateNextPMDate(currentDate, frequency) {
  const date = new Date(currentDate);
  
  switch(frequency) {
    case 'ລາຍວັນ':
      date.setDate(date.getDate() + 1);
      break;
    case 'ລາຍສັບດາ':
      date.setDate(date.getDate() + 7);
      break;
    case 'ລາຍເດືອນ':
      date.setMonth(date.getMonth() + 1);
      break;
    case 'ລາຍປີ':
      date.setFullYear(date.getFullYear() + 1);
      break;
  }
  
  return date;
}

/**
 * Get all PM plans
 */
function getAllPMPlans() {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.PM_PLANS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const pmList = [];
    for (let i = 1; i < data.length; i++) {
      pmList.push({
        id: data[i][0],
        equipmentId: data[i][1],
        equipmentName: data[i][2],
        type: data[i][3],
        frequency: data[i][4],
        startDate: data[i][5],
        nextPMDate: data[i][6],
        checklist: data[i][7],
        status: data[i][8]
      });
    }
    return pmList;
  } catch (e) {
    Logger.log('Error getting PM plans: ' + e.toString());
    return [];
  }
}

// ===== APPROVER MANAGEMENT =====

/**
 * Get approver by equipment ID
 */
function getApproverByEquipment(equipmentId) {
  try {
    const equipment = getEquipmentById(equipmentId);
    if (!equipment) return {email: '', department: ''};
    
    const sheet = ss.getSheetByName(SHEET_NAMES.APPROVERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === equipment.location || data[i][0] === equipment.user) {
        return {
          email: data[i][1],
          name: data[i][2],
          department: data[i][0]
        };
      }
    }
    return {email: '', department: ''};
  } catch (e) {
    Logger.log('Error getting approver: ' + e.toString());
    return {email: '', department: ''};
  }
}

/**
 * Set approver
 */
function setApprover(department, email, name) {
  try {
    const sheet = ss.getSheetByName(SHEET_NAMES.APPROVERS);
    const data = sheet.getDataRange().getValues();
    
    // Check if already exists
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === department) {
        sheet.getRange(i + 1, 2, 1, 2).setValues([[email, name]]);
        return {success: true};
      }
    }
    
    // Add new approver
    sheet.appendRow([department, email, name, new Date()]);
    return {success: true};
  } catch (e) {
    Logger.log('Error setting approver: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

// ===== UTILITY FUNCTIONS =====

/**
 * Generate unique ID
 */
function generateId(prefix) {
  const timestamp = new Date().getTime().toString().slice(-8);
  const random = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return prefix + '-' + timestamp + '-' + random;
}

/**
 * Send approval notification email
 */
function sendApprovalNotification(requestId, recipientEmail, requestData) {
  try {
    const subject = 'ການແຈ້ງສ້ອມອົງກະກອນໃໝ່ - ' + requestId;
    const message = `
      ທ່ານສະບາຍດີ,
      
      ມີການແຈ້ງສ້ອມອຸປະກອນໃໝ່ຕ້ອງໄດ້ຮັບການອະນຸມັດ:
      
      ລະຫັດການແຈ້ງສ້ອມ: ${requestId}
      ອຸປະກອນ: ${requestData.equipmentName}
      ປະເພດວຽກ: ${requestData.workType}
      ລະດັບຄວາມດ່ວນ: ${requestData.urgency}
      ລາຍລະອຽດ: ${requestData.description}
      
      ກະລຸນາລົງທະບຽນເຂົ້າໄປໃນລະບົບເພື່ອອະນຸມັດ
      
      ຂໍ້ໜ້າທີ່ລະບົບ
    `;
    
    GmailApp.sendEmail(recipientEmail, subject, message);
    Logger.log('Approval notification sent to: ' + recipientEmail);
  } catch (e) {
    Logger.log('Error sending email: ' + e.toString());
  }
}

/**
 * Send technician notification
 */
function sendTechnicianNotification(requestId, recipientEmail) {
  try {
    const subject = 'ການແຈ້ງສ້ອມໄດ້ຮັບການອະນຸມັດ - ' + requestId;
    const message = `
      ທ່ານສະບາຍດີ,
      
      ການແຈ້ງສ້ອມ ${requestId} ໄດ້ຮັບການອະນຸມັດແລ້ວ
      ກະລຸນາລົງທະບຽນເຂົ້າໄປໃນລະບົບເພື່ອເບິ່ງລາຍລະອຽດ ແລະ ບັນທຶກຜົນການສ້ອມ
      
      ຂໍ້ໜ້າທີ່ລະບົບ
    `;
    
    GmailApp.sendEmail(recipientEmail, subject, message);
    Logger.log('Technician notification sent to: ' + recipientEmail);
  } catch (e) {
    Logger.log('Error sending email: ' + e.toString());
  }
}

/**
 * Send rejection notification
 */
function sendRejectionNotification(requestId, recipientEmail, reason) {
  try {
    const subject = 'ການແຈ້ງສ້ອມຖືກປະຕິເສດ - ' + requestId;
    const message = `
      ທ່ານສະບາຍດີ,
      
      ການແຈ້ງສ້ອມ ${requestId} ຖືກປະຕິເສດ
      
      ເຫດຜົນ: ${reason}
      
      ກະລຸນາສະແກນເອກະສານ ແລະ ແຈ້ງສ້ອມໃໝ່
      
      ຂໍ້ໜ້າທີ່ລະບົບ
    `;
    
    GmailApp.sendEmail(recipientEmail, subject, message);
    Logger.log('Rejection notification sent to: ' + recipientEmail);
  } catch (e) {
    Logger.log('Error sending email: ' + e.toString());
  }
}

/**
 * Get dashboard statistics
 */
function getDashboardStats() {
  try {
    const allRequests = getAllRepairRequests();
    const allEquipment = getAllEquipment();
    const allSpareParts = getAllSpareParts();
    const allPMPlans = getAllPMPlans();
    
    const stats = {
      totalEquipment: allEquipment.length,
      pendingRequests: allRequests.filter(r => r.status === 'ແຈ້ງສ້ອມ').length,
      inProgressRequests: allRequests.filter(r => r.status === 'ດຳເນີນການ').length,
      completedRequests: allRequests.filter(r => r.status === 'ສ້ອມສຳເລັດ').length,
      lowSpareParts: allSpareParts.filter(sp => sp.remainingQuantity <= sp.warningLevel).length,
      upcomingPMPlans: allPMPlans.filter(pm => {
        const nextDate = new Date(pm.nextPMDate);
        const today = new Date();
        return nextDate <= new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
      }).length
    };
    
    return stats;
  } catch (e) {
    Logger.log('Error getting dashboard stats: ' + e.toString());
    return null;
  }
}

// Initialize when script first loads
if (!PropertiesService.getUserProperties().getProperty('INITIALIZED')) {
  initializeSpreadsheet();
  PropertiesService.getUserProperties().setProperty('INITIALIZED', 'true');
}
