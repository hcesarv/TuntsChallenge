import { GoogleSpreadsheet, GoogleSpreadsheetRow } from 'google-spreadsheet';

class Aluno {
    public avgGrade: number;

    constructor(
        public matricula: number,
        public nome: string,
        public faltas: number,
        public P1: number,
        public P2: number,
        public P3: number,
        public situacao: String,
        public Naf: number
    ){
        //Calculating the average grade...
         this.avgGrade = Number(((P1 + P2 + P3)/3).toFixed(2));
    }
}
var callStudents;

const doc = new GoogleSpreadsheet('1j0p6AO0gkDe1mL8y_-fiocti60x6oJ1G4urTL_VmzyM');

(async function() {
    await major();
  }());

async function major(){
    await ReadStudents();
    
    callStudents.forEach( aluno => {
       calculateStudentSituation(aluno);
       calculateStudentNaf(aluno); 
    })
    
    saveStudents();

}

function calculateStudentSituation(aluno: Aluno){
   // If students absences are above 25% of the 60 classes, , students situation will be 'Reprovado por falta'
    if (aluno.faltas > 15){
        aluno.situacao = 'Reprovado por Falta';
        return 
    }

    // If Average grade from P1, P2 and P3 are beneath 5, students situation will be 'Reprovado por nota'
    if (aluno.avgGrade < 50){
        aluno.situacao = 'Reprovado por nota';
        return
    }

    // If Average grade from P1, P2 and P3 are >= 5 and <= 7, student situation will be 'Exame final'
    if (aluno.avgGrade >= 50 && aluno.avgGrade < 70){
        aluno.situacao = 'Exame Final'
        return
    }

    // If Average grade from P1, P2 and P3 >= 7, student situation will be 'Aprovado'
    aluno.situacao = 'Aprovado'
        
    
}

function calculateStudentNaf(aluno: Aluno){
    // Naf is the minimum grade students need to achieve on the Final Exam. 
    // Naf's formula: 5 <= (m + naf)/2
    // naf >= 10 - m, so Naf's minimum value is (10 - m)
    // If Students situation is different than 'Final Exam', naf = 0.
        if (aluno.situacao != 'Exame Final'){
        aluno.Naf = 0;
        return
    }

    // Round value to next integer. 
    aluno.Naf = Math.round(100 - aluno.avgGrade);
    
}

async function ReadStudents(){
    await accessSpreadsheet();
    let studentArray = [];
    const sheet = doc.sheetsByIndex[0]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]
    console.log(sheet.title);

    
    let lines = await sheet.getRows();
    let countStudents = lines.length;
    

    await sheet.loadCells('A1:H500'); // loads a range of cells


    for(var i=3; i <=countStudents; i++){
        let matricula = Number(sheet.getCell(i,0).value);
        let nome = sheet.getCell(i,1).value;
        let faltas = Number(sheet.getCell(i,2).value);
        let p1 = Number(sheet.getCell(i,3).value);
        let p2 = Number(sheet.getCell(i,4).value);
        let p3 = Number(sheet.getCell(i,5).value);
        let situacao = sheet.getCell(i,6).value;
        let Naf = Number(sheet.getCell(i,7).value);
        studentArray.push(new Aluno(matricula, nome, faltas, p1, p2, p3,situacao,Naf));
    }

callStudents = studentArray;

}

 async function accessSpreadsheet() {
    
    
    await doc.useServiceAccountAuth({
          client_email: "desafiotunts@silicon-dialect-319119.iam.gserviceaccount.com",
          private_key:"-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCxM+zbg/0m2m6r\n6cVhfM0fXzIVfb8Wgm7DN59YorvHrKr+WjkmH6UMKLbdiiabUPsZ1OWh3KFxE6tz\nxd3wymV3qrg7hpo4JPQv1/TZtFhI3YGUCCoMKdZK0FAO26yb+FEyVNi7gtmsLFy5\nxvPCt5ZPBVZ7NdEi2cDsCLORUNpRTEbkcZL9g9UsVG7F3B3ILNwZfzpFMCf3HUMF\nqDLconTk1S61WL2a0Ao/y4JNwqhApD+RuzuRckeIcMrItM3pGA1SZIbEgFYUrSU6\nyro9nZCh4VDahuINKqf7aqkEkhAECDhXO8d42pM79VpuXc50TVLFXbAmjMDbXfRc\nc1V6xSylAgMBAAECggEAAs/9quWXOit5BTPSelUg66YMT/NXzz4mZDp48M8IZvIS\ndrNXjoeDDESF09H9AeBmU9x2q+Z/Y3bkn5Oc4v+mCPZbMoqqM55BMpSIDNUuxN1L\nJgkpIwSeTkAB0p029J9w086CmsXyuX0CGGq+ygY1la8n1mXIxM0ATF6+UYifJm97\nixJGmWp3cJnd4FeA51cSvO/+NBd1AcFDNxLqbC+hD1+JOgNLaYcOACJ6iMA/5zuH\nJYoscTunN7fQIEVj7F5ws69WOuYDlGwCB0RtI6fRluPV4R8VlvHqrCLlr7GqIogR\nHSPeJlKueEUM/8uXBW/TqfOYraALo13CRlV4jymmcwKBgQDlP+rwrLP+K9GW08lu\nmOl7WAFRshXAM8aaAM8L9nAmoMWyl/n8jpzdA5BR7fsTRjxO2Kvmewrx/A9x4KOc\nseoWu216kKKngBdC7ws7hUYKKKgrMAITHVi236tmlz75K1BRj8fv0F9cizG8uaNN\nZULHb0niU24xU7/MEgmle4CJ0wKBgQDF4UdXNB+TTcww8llCPEWdYR9IqVZE2Dui\n0wTWjw8sgbBBSAZV3hpFblITujKlD+A86289R0Zd8qydWL4Urzwjbws//icnBzZS\nX9FHfP9lpd7WwYoWwQQJyfAdpsm3rNG/rZLyEYgahbrlE3XspSrXr65U+BBDIgn2\nOX5YrPkspwKBgQDMQLm7u7QW1sXDY2kKMBV+vvdV6Zx1hewCFIxktSpRUFc0ezHR\nMuNSC7XPOYDFOIPNIEFwddpXpePA64v6tY3CuTWeyTUSlg6jpUXVvzWbIYRMDlca\n3r/HF9un6UPDTzMdqERUR8xfMOmco6167Kil9mLW0szQCDVPxhZwKWxp3wKBgGX9\ntlqhGAFBoRQ9ZXo5PJxgedJmzXtQhHRpFV5NgEGtWp5bNEC/6ISO1ykp2H6xTx+3\nLa/E8+TqdsPnAJoCtBmDW6YRJjb8hagxkNmq+Kx4sQG54aXWuHEfL27pD6FnJvkH\nyuyP0rnw4aK+xBJEE2/2MgHDlgY0HjRV7+Rey1OTAoGADR+YGnQp0VIKWkR2afGA\n5viPyyIHtYbiGjSjfJbefqH/8KlIgSL7IU4XL5KDdeu9pMv8cb82PuBxhG5fQu/1\nSbZgIbQTXFwOsdMyNp2xDz8Iptv9O3Ksc8EWfiw/22ylLfipk8xrHK+KnY3bDJtK\nfn8jM8h0CXDqk5bCwka5jCQ=\n-----END PRIVATE KEY-----\n",
        });
      
    await doc.loadInfo(); // loads document properties and sheets
        console.log(doc.title); // print Sheet title
    
    
}


async function saveStudents(){
    const sheet = doc.sheetsByIndex[0];

    let lines = await sheet.getRows();
    let countStudents = lines.length;

    await sheet.loadCells('A1:H500'); // loads a range of cells
    

    // writes Situation and Approval Grade on Sheet
    for(var i=3; i <=countStudents; i++){ 
        let studentIndex = i-3;

        let situacao = sheet.getCell(i,6);
        let Naf = sheet.getCell(i,7);

        situacao.value = callStudents[studentIndex].situacao;
        Naf.value = callStudents[studentIndex].Naf;

        await sheet.saveUpdatedCells(); // save all updates made.
        
    }

}
