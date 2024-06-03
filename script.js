document.addEventListener('DOMContentLoaded', function() {
    const calendar = document.getElementById('calendar');
    const monthYear = document.getElementById('month-year');
    const daysContainer = document.getElementById('days');
    const prevButton = document.getElementById('prev');
    const nextButton = document.getElementById('next');
    const exportButton = document.getElementById('export');

    const transpiredContainer = document.getElementById('transpired-container');
    const ipcrContainer = document.getElementById('ipcr-container');
    const employeeNumberInput = document.getElementById('employee-number');

    let currentDate = new Date();
    let selectedDates = [];

    function updateCalendar() {
        daysContainer.innerHTML = '';
        const year = currentDate.getFullYear();
        const month = currentDate.getMonth();

        const firstDay = new Date(year, month, 1).getDay();
        const lastDate = new Date(year, month + 1, 0).getDate();

        monthYear.textContent = currentDate.toLocaleDateString('default', { month: 'long', year: 'numeric' });

        for (let i = 0; i < firstDay; i++) {
            const emptyDiv = document.createElement('div');
            daysContainer.appendChild(emptyDiv);
        }

        for (let day = 1; day <= lastDate; day++) {
            const dayDiv = document.createElement('div');
            dayDiv.textContent = day;
            dayDiv.addEventListener('click', () => toggleDateSelection(day));
            if (selectedDates.some(date => date.getDate() === day && date.getMonth() === month && date.getFullYear() === year)) {
                dayDiv.classList.add('selected');
            }
            daysContainer.appendChild(dayDiv);
        }
    }

    function toggleDateSelection(day) {
        const date = new Date(currentDate.getFullYear(), currentDate.getMonth(), day);
        const dateIndex = selectedDates.findIndex(selectedDate => 
            selectedDate.getDate() === date.getDate() && 
            selectedDate.getMonth() === date.getMonth() && 
            selectedDate.getFullYear() === date.getFullYear()
        );
        
        if (dateIndex >= 0) {
            // Date is already selected, deselect it
            selectedDates.splice(dateIndex, 1);
        } else {
            // Date is not selected, select it
            selectedDates.push(date);
        }
        
        updateCalendar();
        console.log(`Selected dates: ${selectedDates}`);
    }

    function exportToExcel() {
        const employeeNumber = employeeNumberInput.value;
        const transpiredInputs = document.querySelectorAll('#transpired-container input[type="text"]');
        const ipcrInputs = document.querySelectorAll('#ipcr-container input[type="text"]');

        // Sort selected dates in chronological order
        selectedDates.sort((a, b) => a - b);

        const workbook = XLSX.utils.book_new();
        const worksheet_data = [['Date', 'What has Transpired', 'IPCR Code', 'Employee Number']];
        
        selectedDates.forEach(date => {
            transpiredInputs.forEach((transpiredInput, index) => {
                const transpired = transpiredInput.value;
                const ipcr = ipcrInputs[index].value;
                worksheet_data.push([date.toLocaleDateString(), transpired, ipcr, employeeNumber]);
            });
        });
        
        const worksheet = XLSX.utils.aoa_to_sheet(worksheet_data);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        XLSX.writeFile(workbook, 'Selected_Dates.xlsx');
    }

    function addTranspiredInput() {
        const newInput = document.createElement('div');
        newInput.innerHTML = `
            <label for="transpired">What has Transpired:</label>
            <input type="text">
        `;
        transpiredContainer.appendChild(newInput);
    }

    function addIPCRInput() {
        const newInput = document.createElement('div');
        newInput.innerHTML = `
            <label for="ipcr">IPCR Code:</label>
            <input type="text">
        `;
        ipcrContainer.appendChild(newInput);
    }

    prevButton.addEventListener('click', () => {
        currentDate.setMonth(currentDate.getMonth() - 1);
        updateCalendar();
    });

    nextButton.addEventListener('click', () => {
        currentDate.setMonth(currentDate.getMonth() + 1);
        updateCalendar();
    });

    exportButton.addEventListener('click', exportToExcel);

    document.getElementById('add-transpired').addEventListener('click', addTranspiredInput);
    document.getElementById('add-ipcr').addEventListener('click', addIPCRInput);

    updateCalendar();
});
