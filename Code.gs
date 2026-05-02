/**
 * @file Code.gs
 * @description نظام إدارة الموارد البشرية والرواتب الذكي المتوافق مع نظام العمل السعودي.
 * @author ahmedaaalshaer-cell
 * @version 2.0.0
 */

/**
 * الوظيفة الرئيسية: توليد 100 سجل موظف مع الحسابات المالية والرقابية.
 */
function generateExpertHRSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  sheet.clear();
  
  // 1. تعريف الهيكل (العناوين) لخدمة الأسئلة العشرة الإحصائية
  const headers = [
    "التاريخ", "الرقم الوظيفي", "اسم الموظف", "الفرع", "المستوى الوظيفي", 
    "الراتب الأساسي", "الأجر الفعلي", "وقت الحضور الفعلي", "وقت الانصراف الفعلي", 
    "نوع العمل", "حالة الحضور", "ساعات الأوفر تايم", "تاريخ الانضمام", 
    "مغادرة مبكرة (دقيقة)", "إجمالي الأوفر تايم (المادة 107)", "مؤشر الانضباط", "احتمالية الاستقالة"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setBackground("#1155cc").setFontColor("white").setFontWeight("bold");

  const branches = ["الرياض", "جدة", "الدمام", "حائل", "المكتب الرئيسي"];
  const rows = [];
  const today = new Date("2026-05-03"); // تاريخ محاكي لليوم لضبط استجابة الوكيل

  // 2. محاكاة بيانات 100 موظف (حالات غياب، أوفر تايم، مستويات قيادية)
  for (let i = 1; i <= 100; i++) {
    let branch = branches[i % branches.length];
    let level = (i <= 5) ? "قيادي" : (i <= 20) ? "مشرف" : "موظف";
    let basic = (level === "قيادي") ? 15000 : (level === "مشرف") ? 10000 : 5000;
    let actual = basic + 2000; 
    
    // متغيرات للتحليل الإحصائي (سؤال 1، 3، 10)
    let isAbsent = (i % 12 === 0); // حالات غياب عشوائية
    let isRemote = (i % 5 === 0); // موظفين يعملون عن بعد
    let earlyExit = (level === "مشرف" && i % 3 === 0) ? 60 : 0; // مغادرة مبكرة للمشرفين
    let joinDate = (i <= 10) ? "2026-01-01" : "2025-06-15"; // موظفون انضموا بداية العام
    let otHours = (i % 4 === 0) ? 12 : 0; // ساعات العمل الإضافي

    rows.push([
      today, 
      "ID-"+(1000+i), 
      "موظف "+i, 
      branch, 
      level, 
      basic, 
      actual, 
      (isAbsent ? "-" : "08:15 AM"), 
      (isAbsent ? "-" : (earlyExit ? "03:00 PM" : "04:00 PM")), 
      (isRemote ? "عن بعد" : "ميداني"), 
      (isAbsent ? "غائب" : "حاضر"), 
      otHours, 
      joinDate, 
      earlyExit
    ]);
  }
  
  sheet.getRange(2, 1, rows.length, 14).setValues(rows);
  
  // 3. حقن المعادلات النظامية (المادة 107 والمادة 84)
  applySaudiLaborLawFormulas(sheet, rows.length + 1);
  
  // تنسيق نهائي
  sheet.setFrozenRows(1);
  sheet.getRange(2, 6, rows.length, 2).setNumberFormat("#,##0");
  sheet.getRange(2, 15, rows.length, 1).setNumberFormat("#,##0.00");
}

/**
 * تطبيق المعادلات بناءً على الأنظمة السعودية والأعمدة المطلوبة للوكيل الذكي.
 */
function applySaudiLaborLawFormulas(sheet, lastRow) {
  const rangeCount = lastRow - 1;

  // أ. إجمالي الأوفر تايم (المادة 107): (الأجر الفعلي / 30 / 8) * 1.5 لكل ساعة
  sheet.getRange(2, 15, rangeCount).setFormulaR1C1("=RC[-3] * ((RC[-8]/30/8) * 1.5)");

  // ب. مؤشر الانضباط (سؤال 4): 100 للحاضر، 60 لمن غادر مبكراً، 0 للغائب
  sheet.getRange(2, 16, rangeCount).setFormulaR1C1("=IF(RC[-5]=\"غائب\", 0, IF(RC[-2]>0, 60, 100))");

  // ج. احتمالية الاستقالة (سؤال 8): مرتفعة إذا كان الانضباط أقل من 70
  sheet.getRange(2, 17, rangeCount).setFormulaR1C1("=IF(RC[-1]<70, \"مرتفعة\", \"منخفضة\")");
}
