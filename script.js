// Function to read and parse the Excel file
document.getElementById('upload-excel').addEventListener('change', handleFile, false);

let certificateData = [];

function handleFile(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Convert worksheet to JSON array
        certificateData = XLSX.utils.sheet_to_json(worksheet);
        console.log(certificateData);
        alert('Excel file uploaded successfully.');
    };

    reader.readAsArrayBuffer(file);
}

// Event listener to handle form submission
document.getElementById('certificate-form').addEventListener('submit', function(e) {
    e.preventDefault();
    const certificateNumberInput = document.getElementById('certificate-input').value.trim(); // Get the certificate number input
    const resultDiv = document.getElementById('result');
    const table = document.getElementById('certificate-details');

    // Find the certificate by certificate number
    const certificate = certificateData.find(item => item['Certificate Number'] === certificateNumberInput);

    if (certificate) {
        // Fill table with certificate details
        document.getElementById('result-certificate-number').textContent = certificate['Certificate Number'];
        document.getElementById('result-participant-name').textContent = certificate['Participant Name'];
        document.getElementById('result-category').textContent = certificate['Category'];
        document.getElementById('result-event-name').textContent = certificate['Event Name'];
        document.getElementById('result-date').textContent = certificate['Date'];

        table.style.display = 'table';
        resultDiv.textContent = 'Certificate is valid';
        resultDiv.style.color = 'green'; // Display message in green color
    } else {
        resultDiv.textContent = 'Certificate not found for certificate number: ' + certificateNumberInput; // Updated message
        resultDiv.style.color = 'red'; // Display error message in red color
        table.style.display = 'none'; // Hide the table if not found
    }
});
