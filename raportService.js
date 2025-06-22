const resultTemplate = {
  in: {
    [VACATION]: [],
    [FAMILY_VACATION]: [],
    [HEALTH_VACATION]: [],
    [PPD]: [],
    [OFFICAL_JOURNEY]: [],
    [HOSPITAL]: [],
    [VLK]: []
  },
  out: {
    [VACATION]: [],
    [FAMILY_VACATION]: [],
    [HEALTH_VACATION]: [],
    [PPD]: [],
    [OFFICAL_JOURNEY]: [],
    [HOSPITAL]: [],
    [VLK]: []
  },
}

const movementTemplate = {
  fullNamePosition: '',
  jobPosition: '',
  militaryRank: '',
  name:'',
  dinnerTime: '',
  facilityName: '',
  documents: ''
}


function generateRaport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange(READ_DATA_RANGE).getValues(); 
  const result = {...resultTemplate}
  groupData(data, result);
  
  // create word
  createRaport(result);
}

function groupData(data, result) {
  data.map(row => {
    
    if(row[IN_COL_IDX]) {
      const movement = getPersonData(row);
      updateDinnerTime(movement, row, true, getDateDuration());
      movement.documents = row[DOCUMENT_COL_IDX];
      movement.facilityName = row[FACILITY_NAME_COL_IDX];
      
      switch(row[DESTINATION_COL_IDX]) {
        case VACATION:
          movement.documents = `відпускний квиток №${getVacationNumber(movement.name)}`;
          result.in[VACATION].push(movement);
          break;
        case HEALTH_VACATION:
          movement.documents = `відпускний квиток №${getVacationNumber(movement.name)}`;
          result.in[HEALTH_VACATION].push(movement);
          break;
        case FAMILY_VACATION:
          movement.documents = `відпускний квиток №${getVacationNumber(movement.name)}`;
          result.in[FAMILY_VACATION].push(movement);
          break;
        case HOSPITAL:
          result.in[HOSPITAL].push(movement);
          break;
        case PPD:
          result.in[PPD].push(movement);
          break;
        case OFFICAL_JOURNEY:
          result.in[OFFICAL_JOURNEY].push(movement);
          break;
        case VLK:
          result.in[VLK].push(movement);
          break;
      }
      
    } else if(row[OUT_COL_IDX]) {
      const movement = getPersonData(row);
      updateDinnerTime(movement, row, false, getDateDuration());
      movement.documents = row[DOCUMENT_COL_IDX];
      movement.facilityName = row[FACILITY_NAME_COL_IDX];
     
      switch(row[DESTINATION_COL_IDX]) {
        case VACATION:
          movement.documents = `відпускний квиток №${getVacationNumber(movement.name)}`;
          result.out[VACATION].push(movement);
          break;
        case HEALTH_VACATION:
          movement.documents = `відпускний квиток №${getVacationNumber(movement.name)}`;
          result.out[HEALTH_VACATION].push(movement);
          break;
        case FAMILY_VACATION:
          movement.documents = `відпускний квиток №${getVacationNumber(movement.name)}`;
          result.out[FAMILY_VACATION].push(movement);
          break;
        case HOSPITAL:
          result.out[HOSPITAL].push(movement);
          break;
        case PPD:
          result.out[PPD].push(movement);
          break;
        case OFFICAL_JOURNEY:
          result.out[OFFICAL_JOURNEY].push(movement);
          break;
        case VLK:
          result.out[VLK].push(movement);
          break;
      }
    }
  })

} 

function getPersonData(row) {
  const person = {...movementTemplate};
  person.militaryRank = row[FULL_RANK_COL_IDX];
  person.name = row[FULL_NAME_COL_IDX];
  person.jobPosition = row[FULL_JOB_COL_IDX];
  person.fullNamePosition = `${person.militaryRank} ${person.name}, ${person.jobPosition}`;
  return person;
}

function updateDinnerTime(person, row, isIn, todayData) {
  const kitchen = isIn ? 'Прошу зарахувати на продовольче забезпечення' : 'Прошу зняти з продовольчого забезпечення ';
  
  if(row[DINNER_TIME_COL_IDX] === 'сніданок') {
    person.dinnerTime = `${kitchen} (за каталогом, коефіцієнтом 1,2) з сніданку ${todayData}.`
  } else if (row[DINNER_TIME_COL_IDX] === 'обід') {
    person.dinnerTime = `${kitchen} (за каталогом, коефіцієнтом 1,2) з обіду ${todayData}.`
  } else {
    person.dinnerTime = `${kitchen} (за каталогом, коефіцієнтом 1,2) з вечері ${todayData}.`
  }
    
}










