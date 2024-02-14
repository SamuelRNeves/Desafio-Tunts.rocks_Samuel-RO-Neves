const { google } = require('googleapis');
const fs = require('fs');

async function accessAndManipulateSpreadsheet() {
  try {
    // Load the JSON credentials file
    const credentials = JSON.parse(fs.readFileSync('C:\\Users\\NRDS\\Desktop\\Desafio\\crucial-cycling-414019-7e84624edbaa.json'));

    // Set up authentication
    const auth = new google.auth.GoogleAuth({
      credentials: credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    // Create a client instance of the Google Sheets API
    const sheets = google.sheets({ version: 'v4', auth });

    // Access the spreadsheet and read data
    const readResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: '1PHP7xuSWk2kMiTAuXJUOFjEsX3NlYcauGNKSgzFnC8I',
      range: 'EngenhariaDeSoftware!A4:H27',
    });
    const values = readResponse.data.values;

    // Ensure values exist and have correct structure
    if (!values || !Array.isArray(values) || values.length === 0) {
      throw new Error('No data found in the spreadsheet or incorrect data structure.');
    }

    // Define constants for calculation
    const totalClasses = 60; // Total number of classes
    const maxAllowedAbsences = totalClasses * 0.25; // Maximum allowed absences (25% of total classes)

    // Iterate through the values and calculate the situation for each student
    const updatedValues = values.map(row => {
      const enrollment = row[0];
      const Alunos = row[1];
      const Faltas = parseInt(row[2]);
      const p1 = parseFloat(row[3]);
      const p2 = parseFloat(row[4]);
      const p3 = parseFloat(row[5]);

      const average = (p1 + p2 + p3) / 3;

      let situation;
      let finalGrade = 0; // Default value for final grade

      if (Faltas > maxAllowedAbsences) {
        situation = 'Failed due to Absences';
      } else if (average < 5) {
        situation = 'Failed due to Grade';
      } else if (average < 7) {
        situation = 'Final Exam';
        const naf = Math.max(0, Math.ceil(10 - average * 2)); // Calculate the NAF and round up if necessary
        finalGrade = Math.ceil((average + naf) / 2); // Calculate the final grade and round up if necessary
        if (finalGrade >= 5) {
          situation = 'Approved';
        }
      } else {
        situation = 'Approved';
      }

      return [enrollment, Alunos, Faltas, p1, p2, p3, situation, finalGrade];
    });

    // Write the results back to the spreadsheet
    const updateResponse = await sheets.spreadsheets.values.update({
      spreadsheetId: '1PHP7xuSWk2kMiTAuXJUOFjEsX3NlYcauGNKSgzFnC8I',
      range: 'EngenhariaDeSoftware!A4:H27', // Adjust the range to write final grade and situation
      valueInputOption: 'RAW',
      requestBody: { values: updatedValues },
    });

    console.log('Results written to the spreadsheet successfully.');

  } catch (error) {
    console.error('Error accessing or manipulating the spreadsheet:', error);
  }
}

accessAndManipulateSpreadsheet();
