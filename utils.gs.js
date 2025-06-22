const VACATION = 'відпустка щорічна';
const HEALTH_VACATION = 'відпустка СЗ';
const FAMILY_VACATION = 'відпустка СО';
const HOSPITAL = 'шпиталь';
const PPD = 'ппд';
const OFFICAL_JOURNEY = 'відрядження';
const VLK = 'ВЛК';

const GROUP_TYPE= [VACATION, HEALTH_VACATION, FAMILY_VACATION, HOSPITAL, PPD, OFFICAL_JOURNEY];

const GROUNDS_RANGE = "I3:L";
const CHECKBOX_RANGE = "G3:H";
const READ_DATA_RANGE = "A3:L";

const FULL_RANK_COL_IDX = 3;
const FULL_NAME_COL_IDX = 4;
const FULL_JOB_COL_IDX = 5;
const IN_COL_IDX = 6;
const OUT_COL_IDX = 7;
const DESTINATION_COL_IDX = 8;
const DINNER_TIME_COL_IDX = 9;
const FACILITY_NAME_COL_IDX = 10;
const DOCUMENT_COL_IDX = 11;


const MONTH_TO = [  "січеня",  "лютого",  "березня",  "квітня",  "травня",  "червня",  "липня",  "серпня",  "вересня",  "жовтня",  "листопада", "грудня"];


function clearCheckbox() {
  const sheed = SpreadsheetApp.getActiveSpreadsheet();
  sheed.getDataRange().uncheck();
}

function getSimpleDate() {
    const today = new Date();

  const day = today.getDate() < 10 ? "0" + today.getDate() : today.getDate();         
  const month = today.getMonth() + 1 < 10 ? "0" + (today.getMonth() + 1) : today.getMonth() + 1; 
  const year = today.getFullYear(); 

   return `${day}.${month}.${year}`;
}

function getDateDuration(){
  const today = new Date();

  const day = today.getDate();         
  const month = today.getMonth(); 
  const year = today.getFullYear(); 

  return `${day} ${MONTH_TO[month]} ${year}`;
}

function setParagraphStyle(paragraph, options = {}) {
  try {
    const text = paragraph.editAsText();
    
    if (options.fontSize) {
      text.setFontSize(options.fontSize);
    }
    
    if (options.bold) {
      text.setBold(options.bold);
    }
    
    if (options.italic) {
      text.setItalic(options.italic);
    }
    
    if (options.underline) {
      text.setUnderline(options.underline);
    }
    
    if (options.color) {
      text.setForegroundColor(options.color);
    }
    
    if (options.fontFamily) {
      text.setFontFamily(options.fontFamily);
    }
    
    if (options.alignment) {
      paragraph.setAlignment(options.alignment);
    }
    
    if (options.lineSpacing) {
      paragraph.setLineSpacing(options.lineSpacing);
    }
    
    if (options.spacingBefore) {
      paragraph.setSpacingBefore(options.spacingBefore);
    }
    
    if (options.spacingAfter) {
      paragraph.setSpacingAfter(options.spacingAfter);
    }
    
  } catch (error) {
    Logger.log('Error setting paragraph style: ' + error.toString());
    throw error;
  }
}





