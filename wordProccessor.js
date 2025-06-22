function createRaport(data) {
  createWordDocument(data)
}

function createWordDocument(data) {
  try {
    const doc = DocumentApp.create(` Рапорт рух особового складу в районі виконання завдання за ${getSimpleDate()}`);
    const body = doc.getBody();
    body.clear();

    appendHeader(body);
    appendData(data, body);
    appendFooter(body);

    doc.saveAndClose();
    downloadAsWord(doc.getId(), doc.getName());
    
  } catch (error) {
    Logger.log('Error creating document: ' + error.toString());
    throw error;
  }
}

const outgoingMovements = [
  { key: VACATION, label: 'В ЩОРІЧНУ ВІДПУСТКУ:', handler: appendTo },
  { key: HEALTH_VACATION, label: 'В ВІДПУСТКУ ЗА СТАНОМ ЗДОРОВ`Я:', handler: appendTo },
  { key: FAMILY_VACATION, label: 'В ВІДПУСТКУ ЗА СІМЕЙНИМИ ОБСТАВИНАМИ:', handler: appendTo },
  { key: PPD, label: 'В ПУНКТ ПОСТІЙНОЇ ДИСЛОКАЦІЇ:', handler: appendTo },
  { key: OFFICAL_JOURNEY, label: 'В ВІДРЯДЖЕННЯ:', handler: appendToFacilityWrapper },
  { key: HOSPITAL, label: 'У ЛІКУВАЛЬНИЙ ЗАКЛАД НА СТАЦІОНАРНЕ ЛІКУВАННЯ:', handler: appendToFacilityWrapper },
  { key: VLK, label: 'У ЛІКУВАЛЬНИЙ ЗАКЛАД НА ВЛК:', handler: appendToFacilityWrapper }
];

const incomingMovements = [
  { key: VACATION, label: 'З ЩОРІЧНОЇ ВІДПУСТКИ:', handler: appendVacation },
  { key: HEALTH_VACATION, label: 'З ВІДПУСТКИ ЗА СТАНОМ ЗДОРОВ`Я:', handler: appendVacation },
  { key: FAMILY_VACATION, label: 'З ВІДПУСТКИ ЗА СІМЕЙНИМИ ОБСТАВИНАМИ:', handler: appendVacation },
  { key: PPD, label: 'З ПУНКТУ ПОСТІЙНОЇ ДИСЛОКАЦІЇ:', handler: appendFromPPDWrapper },
  { key: OFFICAL_JOURNEY, label: 'З ВІДРЯДЖЕННЯ:', handler: appendFromFacilityWrapper(appendVacation) },
  { key: HOSPITAL, label: 'З ЛІКУВАЛЬНОГО ЗАКЛАДУ ПІСЛЯ СТАЦІОНАРНОГО ЛІКУВАННЯ:', handler: appendFromFacilityWrapper(appendFromHospital) },
  { key: VLK, label: 'З ЛІКУВАЛЬНОГО ЗАКЛАДУ ПІСЛЯ ВЛК:', handler: appendFromFacilityWrapper(appendFromVLK) }
];

function hasMovements(directionData, directions) {
  return directions.some(({ key }) => directionData[key]?.length > 0);
}

function appendData(data, body) {
  const sections = [
    { type: 'in', label: 'ПРИБУТТЯ:', directions: incomingMovements },
    { type: 'out', label: 'ВИБУТТЯ:', directions: outgoingMovements }
  ];

  for (const section of sections) {
    if (hasMovements(data[section.type], section.directions)) {
      body.appendParagraph('');
      const title = body.appendParagraph(section.label);
      setRowBold(title);

      section.directions.forEach(({ key, label, handler }) => {
        if (data[section.type][key]?.length > 0) {
          body.appendParagraph('');
          const subTitle = body.appendParagraph(label);
          setRowBold(subTitle);

          data[section.type][key].forEach(movement => handler(movement, body));
        }
      });
    }
  }

}

function appendFromPPDWrapper(movement, body) {
  const row = body.appendParagraph(`    ${movement.fullNamePosition}.`);
  row.setBold(false);
  setMovementStyle(row);

  const dinner = body.appendParagraph(`     ${movement.dinner}`);
  const reason = body.appendParagraph('     Підстава: іменний список.');
  [dinner, reason].forEach(p => {
    setMovementStyle(p);
    p.setBold(false);
  });
}

function appendFromFacilityWrapper(handler) {
  return function(movement, body) {
    const main = body.appendParagraph(`З ${movement.facilityName}:`);
    main.setBold(false);
    setMovementStyle(main);
    handler(movement, body);
  };
}

function appendToFacilityWrapper(movement, body) {
  const main = body.appendParagraph(`У ${movement.facilityName}:`);
  main.setBold(false);
  setMovementStyle(main);
  appendTo(movement, body);
}

function appendFromPPD(movement, body) {
  const mainRow = body.appendParagraph(`    ${movement.fullNamePosition}.`); 
  mainRow.setBold(false);
  setMovementStyle(mainRow);
 
}

function appendTo(movement, body) {
  body.appendParagraph('');
  const mainRow = body.appendParagraph(`    ${movement.fullNamePosition}. ${movement.dinnerTime}`); 
  mainRow.setBold(false);
  setMovementStyle(mainRow);
  const reasonRow = body.appendParagraph(`     Підстава: ${movement.documents}.`);
  setMovementStyle(reasonRow);
  reasonRow.setBold(false);
}

function appendVacation(movement, body) {
  body.appendParagraph('');
  const mainRow = body.appendParagraph(`    ${movement.fullNamePosition}. ${movement.dinnerTime}`); 
  mainRow.setBold(false);
  setMovementStyle(mainRow);
  const reasonRow = body.appendParagraph(`     Підстава: ${movement.documents}.`);
  setMovementStyle(reasonRow);
  reasonRow.setBold(false);
}

function appendFromHospital(movement, body) {
  body.appendParagraph('');
  const mainRow = body.appendParagraph(`    ${movement.fullNamePosition}. ${movement.dinnerTime}`); 
  mainRow.setBold(false);
  setMovementStyle(mainRow);
  const reasonRow = body.appendParagraph(`     Підстава: виписка із медичної карти стаціонарного хворого ${movement.documents}.`);
  setMovementStyle(reasonRow);
  reasonRow.setBold(false);
}

function appendFromVLK(movement, body) {
  body.appendParagraph('');
  const mainRow = body.appendParagraph(`    ${movement.fullNamePosition}. ${movement.dinnerTime}`); 
  mainRow.setBold(false);
  setMovementStyle(mainRow);
  const reasonRow = body.appendParagraph(`     Підстава: висновок військово лікарської комісії ${movement.documents}.`);
  setMovementStyle(reasonRow);
  reasonRow.setBold(false);
}

function setMovementStyle(row) {
  setParagraphStyle(row, {
    fontSize: 13,bold: false,alignment: DocumentApp.HorizontalAlignment.JUSTIFY,fontFamily: 'Times New Roman'
  });
}

function setRowBold(row) {
  setParagraphStyle(row, {
    fontSize: 13,bold: true, alignment: DocumentApp.HorizontalAlignment.LEFT, fontFamily: 'Times New Roman'
  });
}

function appendHeader(body) {
  const headerLines = [
    'Командиру військової частини А4350', '', '', '',
    'Рапорт',
    `    Дійсним доповідаю про рух особового складу 2 аеромобільного батальйону військової частини А4350 в район виконання завдання протягом ${getDateDuration()}:`
  ];

  headerLines.forEach((text, i) => {
    const p = body.appendParagraph(text);
    const style = i === 0 ? DocumentApp.HorizontalAlignment.RIGHT :
                 i === 4 ? DocumentApp.HorizontalAlignment.CENTER : DocumentApp.HorizontalAlignment.JUSTIFY;

    setParagraphStyle(p, { fontSize: 13, alignment: style, fontFamily: 'Times New Roman' });
  });
}

function appendFooter(body) {
  const lines = [
    '', '', '',
                                                                            ',
    `${getSimpleDate()}`
  ];

  lines.forEach(text => {
    const p = body.appendParagraph(text);
    setParagraphStyle(p, { fontSize: 13, alignment: DocumentApp.HorizontalAlignment.LEFT, fontFamily: 'Times New Roman' });
  });
}




function downloadAsWord(docId, fileName) {
  try {
    downloadFileFromDrive(docId, fileName)
  } catch (error) {
    Logger.log('Error downloading as Word: ' + error.toString());
    throw error;
  }
}





