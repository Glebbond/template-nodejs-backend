const Excel = require('exceljs/modern.nodejs');
const moment = require('moment');

module.exports = async (data, res) => {
  var workbook = new Excel.Workbook();
  var worksheet =  workbook.addWorksheet('sheet', {
    pageSetup:{paperSize: 9, orientation:'landscape'}
  });
  worksheet.getColumn('A').alignment = { horizontal:'center', vertical: 'middle', wrapText: true};
  worksheet.getColumn('B').alignment = { horizontal:'center', vertical: 'middle', wrapText: true};
  worksheet.getColumn('C').alignment = { horizontal:'center', vertical: 'middle', wrapText: true};
  worksheet.getColumn('D').alignment = { horizontal:'center', vertical: 'middle', wrapText: true};

  worksheet.mergeCells('A1:D1');
  worksheet.getCell('A1').value = `School Year ${moment().format('YYYY')}: School and the Educational Climate (SSEC) Summary Data Collection Form`;
  worksheet.mergeCells('A2:B3');
  worksheet.getCell('A2').value = 'Part 1:Dignity for All Student Act (DASA) and Violent and Disruptive Incident Reporting (VADIR)*';
  worksheet.mergeCells('C2:D2');
  worksheet.getCell('C2').value = 'Material Incidents of Discrimination, Harassment, and Bullying';
  worksheet.getCell('C3').value = 'All Excluding Cyberbullying';
  worksheet.getCell('D3').value = 'Cyberbullying';

  worksheet.mergeCells('A4:D4');
  worksheet.getCell('A4').value = 'Report the total number of incidents. Count each incident only one time regardless of the number of offenders or targets/victims involved. For incidents that fit more than one category, choose the most serious (higher weighted category).';
  worksheet.addRow(['Total Number of Incidents', 'a', data.numberOther, data.numberCyber]);
  worksheet.mergeCells('A6:D6');

  worksheet.getCell('A6').value = `Report if the offense listed in row (a)  was related to a bias. * Note that if appropriate, an incident may be reported for more than one bias (duplicated count). For example, if an  Assault with Physical Injury was related to the Victim/Target's Religion and Gender, it should be reported in both rows.  See directions for additional information.`;
  worksheet.addRow(['Total Number of Biased-Related Incidents', 'b', data.biasedOther, data.biasedCyber]);
  worksheet.addRow(['Race', 'c', data.raceOther, data.raceCyber]);
  worksheet.addRow(['Ethnic Group', 'd', data.ethnicOther, data.ethnicCyber]);
  worksheet.addRow(['National Origin', 'e', data.nationalOther, data.nationalCyber]);
  worksheet.addRow(['Color', 'f', data.colorOther, data.colorCyber]);
  worksheet.addRow(['Religion', 'g', data.religionOther, data.religionCyber]);
  worksheet.addRow(['Religious Practices', 'h', data.religiousOther, data.religiousCyber]);
  worksheet.addRow(['Disability', 'i', data.disabilityOther, data.disabilityCyber]);
  worksheet.addRow(['Gender', 'j', data.genderOther, data.genderCyber]);
  worksheet.addRow(['Sexual Orientation', 'k', data.sexualOther, data.sexualCyber]);
  worksheet.addRow(['Sex', 'l', data.sexOther, data.sexCyber]);
  worksheet.addRow(['Weight', 'm', data.weightOther, data.weightCyber]);
  worksheet.addRow(['Other', 'n', data.otherOther, data.otherCyber]);
  worksheet.mergeCells('A20:D20');
  worksheet.getCell('A20').value = `Report the number of incidents in row (a) that were gang/group related.`;
  worksheet.addRow(['Gang or Group Related', 'o', data.gangOther, data.gangCyber]);
  worksheet.mergeCells('A22:D22');
  worksheet.getCell('A22').value = `Report the number of incidents in row (a) that involved a weapon, alcohol, and/or drugs. The sum of rows (p) and (q) must equal the number reported in row (a)* Note rows (q1-q3) may be duplicated counts if an incident involved more than one weapon.`;
  worksheet.addRow(['Total Number of Incidents Not Involving a Weapon', 'p', data.notWeaponOther, data.notWeaponCyber]);
  worksheet.addRow(['Number of Incidents Involving Alcohol ', 'r', data.alcoholOther, data.alcoholCyber]);
  worksheet.addRow(['Number of Incidents Involving Drugs ', 's', data.drugsOther, data.drugsCyber]);
  worksheet.mergeCells('A26:D26');

  worksheet.mergeCells('A27:B28');
  worksheet.getCell('A27').value = 'Part 1:Dignity for All Student Act (DASA) and Violent and Disruptive Incident Reporting (VADIR)*';
  worksheet.mergeCells('C27:D27');
  worksheet.getCell('C27').value = 'Material Incidents of Discrimination, Harassment, and Bullying';
  worksheet.getCell('C28').value = 'All Excluding Cyberbullying';
  worksheet.getCell('D28').value = 'Cyberbullying';
  worksheet.mergeCells('A29:D29');
  worksheet.getCell('A29').value = 'Report the location where incidents reported in row (a) occurred - report each incident only one time. The sum of rows (t), (u), and (v) must equal the number reported in row (a).';
  worksheet.addRow(['On School Property (including on school transportation)', 't', data.onSchoolOther, data.onSchoolCyber]);
  worksheet.addRow(['At School Function Off Grounds', 'u', data.atSchoolOther, data.atSchoolCyber]);
  worksheet.addRow(['Off School Property (that creates a risk of disruption within the school environment)', 'v', data.offSchoolOther, data.offSchoolCyber]);
  worksheet.addRow(['Of the incidents reported in Row (t) above, report the number that occurred on School Transportation', 'w', data.onTransportationOther, data.onTransportationCyber]);
  worksheet.mergeCells('A34:D34');
  worksheet.getCell('A34').value = 'Report the number of incidents in row (a) that occurred during the regular school day and after school hours. The sum of rows (x) and (y) must equal the number reported in row (a).';
  worksheet.addRow(['During Regular School Hours', 'x', data.duringRegularHoursOther, data.duringRegularHoursCyber]);
  worksheet.addRow([' Outside of Regular School Hours', 'y', data.outsideRegularHoursOther, data.outsideRegularHoursCyber]);
  worksheet.mergeCells('A37:D37');
  worksheet.getCell('A37').value = 'Report the number of Targets/Victims that were students, staff or other involved in incidents in row (a). A target/victim must be counted more than once if he/she is a target/victim of more than one incident (duplicated count).';
  worksheet.addRow(['Number of Student Targets/Victims', 'z', data.victimsOther, data.victimsCyber]);
  worksheet.mergeCells('A39:D39');
  worksheet.getCell('A39').value = 'Report the number of OFFENDERS that were students, staff or other involved in incidents in row (a). An offender must be counted more than once if he/she initiates more than one incident (duplicated count).';
  worksheet.addRow(['Number of Student Offenders', 'cc', data.studentOffendersOther, data.studentOffendersCyber]);
  worksheet.addRow(['Number of Staff Offenders', 'dd', data.staffOffendersOther, data.staffOffendersCyber]);
  worksheet.addRow(['Number of "Other" Offenders', 'ee', data.otherOffendersOther, data.otherOffendersCyber]);
  worksheet.mergeCells('A43:D43');
  worksheet.getCell('A43').value = ' Report the number of STUDENT OFFENDERS that received the following type of discplinary action or referral (Report all that apply).';
  worksheet.addRow(['Counseling or Treatment Programs', 'ff', data.counselingOther, data.counselingCyber]);
  worksheet.addRow(['Teacher Removal (Section 3214)', 'gg', data.teacherOther, data.teacherCyber]);
  worksheet.addRow(['In School Suspension', 'hh', data.inSchoolSuspensionOther, data.inSchoolSuspensionCyber]);
  worksheet.addRow(['Out-of-School Suspension', 'ii', data.outSchoolSuspensionOther, data.outSchoolSuspensionCyber]);
  worksheet.addRow(['Involuntary Transfer to an Alternative Placement', 'jj', data.involuntaryTransferOther, data.involuntaryTransferCyber]);
  worksheet.addRow(['Community Service', 'kk', data.communityServiceOther, data.communityServiceCyber]);
  worksheet.addRow(['Juvenile Justice Or Criminal Justice System', 'll', data.juvenileJusticeOther, data.juvenileJusticeCyber]);
  worksheet.addRow(['Law Enforcement', 'mm', data.lawEnforcementOther, data.lawEnforcementCyber]);
  worksheet.mergeCells('A52:D52');
  worksheet.getCell('A52').value = ' Report the unduplicated count of STUDENT OFFENDERS.';
  worksheet.addRow(['Number of Unduplicated Student Offenders for Serious Incidents', 'nn', data.unduplicatedOther, data.unduplicatedCyber]);
  worksheet.mergeCells('A54:D54');
  worksheet.getCell('A54').value = `*Items collected on this form are required by Education Law §2802 and Commissioner's Regulation 100.2 (gg) as amended in December 2016 (http://www.regents.nysed.gov/common/regents/files/1216p12a2.pdf).			`;
  let col = worksheet.getColumn('A');
  col.eachCell(cell => {
    cell.border = {
      top: {style:'thin'},
      left: {style:'thin'},
      bottom: {style:'thin'},
      right: {style:'thin'}
    }
    cell.font ={
      size: 8, bold: true
    }
  })
  col.width = 30;
  col = worksheet.getColumn('B');
  col.eachCell(cell => {
    cell.border = {
      top: {style:'thin'},
      left: {style:'thin'},
      bottom: {style:'thin'},
      right: {style:'thin'}
    }
    cell.font ={
      size: 8, bold: true
    }
    
  })
  col.width = 5;
  col = worksheet.getColumn('C');
  col.eachCell(cell => {
    cell.border = {
      top: {style:'thin'},
      left: {style:'thin'},
      bottom: {style:'thin'},
      right: {style:'thin'}
    }
    cell.font ={
      size: 8, bold: true
    }
    
  })
  col.width = 25;
  col = worksheet.getColumn('D');
  col.eachCell(cell => {
    cell.border = {
      top: {style:'thin'},
      left: {style:'thin'},
      bottom: {style:'thin'},
      right: {style:'thin'}
    }
    cell.font ={
      size: 8, bold: true
    }
    
  })
  col.width = 25;
  worksheet.getCell('A1').font = {size: 14, bold: true};
  worksheet.getCell('A2').font = {size: 11, bold: true};
  worksheet.getCell('C2').font = {size: 11, bold: true};
  worksheet.getCell('C3').font = {size: 11, bold: true};
  worksheet.getCell('D3').font = {size: 11, bold: true};
  let row;
  for (let i = 7; i <= 19; i++) {
    row = worksheet.getRow(`${i}`);
    row.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{argb:'F7CAAC'},
      };
    })
  }
  worksheet.getRow('21').eachCell(cell => {
    cell.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'DADADA'},
  }});
  for (let i = 23; i <= 25; i++) {
    row = worksheet.getRow(`${i}`);
    row.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{argb:'FFE598'},
      };
    })
  }
  for (let i = 30; i <= 33; i++) {
    row = worksheet.getRow(`${i}`);
    row.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{argb:'BDD6EE'},
      };
    })
  }
  for (let i = 35; i <= 36; i++) {
    row = worksheet.getRow(`${i}`);
    row.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{argb:'BDAED2'},
      };
    })
  }
  worksheet.getRow('38').eachCell(cell => {
    cell.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'C5E0B3'},
  }});
  for (let i = 40; i <= 42; i++) {
    row = worksheet.getRow(`${i}`);
    row.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{argb:'ADB9CA'},
      };
    })
  }
  for (let i = 44; i <= 51; i++) {
    row = worksheet.getRow(`${i}`);
    row.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{argb:'D6DCE4'},
      };
    })
  }
  worksheet.getRow('53').eachCell(cell => {
    cell.fill = {
    type: 'pattern',
    pattern:'solid',
    fgColor:{argb:'E6A8A8'},
  }});
  worksheet.getRow('1').height = 40;
  worksheet.getRow('2').height = 30;
  worksheet.getRow('3').height = 30;
  worksheet.getRow('4').height = 30;
  worksheet.getRow('6').height = 30;
  worksheet.getRow('22').height = 30;
  worksheet.getRow('27').height = 30;
  worksheet.getRow('28').height = 30;
  worksheet.getRow('29').height = 30;
  worksheet.getRow('34').height = 30;
  worksheet.getRow('37').height = 30;
  worksheet.getRow('39').height = 30;
  worksheet.getRow('43').height = 30;
  worksheet.getRow('54').height = 30;

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader("Content-Disposition", "attachment; filename=" + "Report.xlsx");
  await workbook.xlsx.write(res);
  res.end();
}