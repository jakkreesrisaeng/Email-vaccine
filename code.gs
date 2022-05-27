// @ts-nocheck
/**
* Sends non-duplicate emails with data from the current spreadsheet.
*/
var EMAIL_SENT = 'EMAIL_SENT' ;
 
function sendEmails2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getDataRange();
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 1; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[3]; // First column  ที่อยู่ของ E-mail 
    var message = 'เรื่อง ข้อความตอบกลับสำหรับผู้ยื่นเอกสารการขอรับหนังสือรับรอง (COVID-19 Vaccine Passport )(ข้อความตอบกลับอัตโนมัติ)' + '\n\n' + 'เรียน   ' + row[1] + '\n\n' +  '   สำนักงานป้องกันควบคุมโรคที่ 8 จังหวัดอุดรธานี ได้รับเอกสารประกอบการขอรับหนังสือรับรอง COVID-19 Vaccine Passport แล้ว ท่านสามารถติดตามสถานะการออกเอกสารได้ผ่านลิงค์ https://bit.ly/3DKWS2r สอบถามรายละเอียดเพิ่มเติม สามารถติดต่อได้ที่ (099-089-9940)' + '\n\n\n' + 'ขอแสดงความนับถือ' + '\n\n' + 'สำนักงานป้องกันควบคุมโรคที่ 8 จังหวัดอุดรธานี'  + '\n\n' + 'หมายเหตุ :ภายใน 7 วันทำการท่านจะได้รับหนังสือรับรองฯ หากเอกสารครบถ้วนถูกต้อง สามารถติดตามสถานะผ่านลิ้งค์ https://bit.ly/3DKWS2r' +'\n'+ '1. ขอรับหนังสือรับรองฯ ในวันราชการ เวลา 09.00 น. - 15.30 น.' + '\n' + '2. ตรวจสอบข้อมูลในเล่ม ' + '\n' + '3. ถ้าพบข้อมูลไม่ถูกต้องให้แจ้งเจ้าหน้าที่ทันทีเพื่อแก้ไขหรือโทร 099-089-9940' + '\n' +'4. กรุณาลงลายมือชื่อในช่อง whose signature follows หน้าที่ 3 ให้เหมือนกับที่ลงนามในหนังสือเดินทาง' + '\n' + 
'************************************************' +'\n\n' + 'Subject: Response for requesting COVID-19 Vaccine Passport(Auto reply do not response)' + '\n\n' +  'Dear   ' + row[1] + '\n\n' + '    The Office of Disease Prevention and Control 8 Udonthani had received your request for COVID – 19 Vaccine Passport.You can check your status via https://bit.ly/3DKWS2r , any inquiries please contact (+6699-089-9940)'  + '\n\n\n' +    'Best regards,' + '\n\n' + 'The Office of Disease Prevention and Control 8 Udonthani ' + '\n\n' + 'Annotation: COVID-19 Vaccine Passport will be issued within 7 official days , Following your status via https://bit.ly/3DKWS2r' + '\n' + '1. Pick up during official days (09.00 a.m. - 03.30 p.m.' + '\n' + '2. Recheck your COVID-19 Vaccine Passport.' + '\n' + '3.Please notify the officers in case of incorrect information or contact +6699-089-9940.' + '\n' + '4. Sign in the box (whose signature follows page , the same as signature on your passport). ' + '\n' + 
'************************************************' +'\n\n' + 'Objet : Réponse à la demande de passeport sanitaire pour le vaccin COVID-19 (La réponse automatique ne répond pas)' + '\n\n' + ' Chère   ' + row[1] + '\n\n' + '     Le Bureau de prévention et de contrôle des maladies 8 Udonthani a reçu votre formulaire de demande de vaccination COVID-19 : passeport vaccin COVID-19). Vous pouvez vérifier le résultat du processus via https://bit.ly/3DKWS2r De plus, si vous avez des questions, veuillez nous contacter (+6699-089-9940)' + '\n\n\n' + ' Meilleures salutations, ' + '\n\n' + 'Le Bureau de prévention et de contrôle des maladies 8 Udonthani' + '\n\n' +  'Annotation :Pour Le passeport vaccin COVID -19 va publié dans les 7 jours officiels Vous pouvez suivre votre situation via https://bit.ly/3DKWS2r' + '\n' + '1.Récupérer le dans le  jour official(09.00 a.m. - 03.30 p.m.)' + '\n' + '2.Revérifiez votre passeport pour le vaccin COVID-19. '+ '\n' + '3.Veuillez informer les agents si vous trouvez des informations incorrectes ou contactez le +6699-089-9940. '+ '\n' + '4.Signez dans la case, dont la signature suit la page la même que la signature sur votre passeport.    '
;  // ระบุ Message ตรงนี้
    var emailSent = row[8]; // แถวที่ 6+1  คือตำแหน่งของข้อความ EMAIL_SENT ปรากฏ
    if (emailSent != EMAIL_SENT) { // Prevents sending duplicates
      var subject = "ระบบตอบกลับข้อความอัตโนมัติ(Auto reply)(ไทย ,English, French Version) "; // ระบุ Subject ตรงนี้
      MailApp.sendEmail(emailAddress, subject, message);
      sheet.getRange(i+1, 9).setValue(EMAIL_SENT);  // ตัวเลข 7 คือตำแหน่งของข้อความ EMAIL_SENT ปรากฏ
// Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
 
