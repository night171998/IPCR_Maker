document.addEventListener("DOMContentLoaded", function() {
    const calendar = document.getElementById("calendar");
    const monthYear = document.getElementById("month-year");
    const prev = document.getElementById("prev");
    const next = document.getElementById("next");
    const daysContainer = document.getElementById("days");
    const transpiredContainer = document.getElementById("transpired-container");
    const ipcrContainer = document.getElementById("ipcr-container");
    const exportButton = document.getElementById("export");
    const employeeNumberInput = document.getElementById("employee-number");

    let currentMonth = new Date().getMonth();
    let currentYear = new Date().getFullYear();
    let selectedDates = new Map();

    function generateCalendar(month, year) {
        const firstDay = new Date(year, month).getDay();
        const daysInMonth = 32 - new Date(year, month, 32).getDate();
        
        monthYear.textContent = `${new Date(year, month).toLocaleString('default', { month: 'long' })} ${year}`;
        daysContainer.innerHTML = "";

        for (let i = 0; i < firstDay; i++) {
            const emptyDiv = document.createElement("div");
            daysContainer.appendChild(emptyDiv);
        }

        for (let day = 1; day <= daysInMonth; day++) {
            const dayDiv = document.createElement("div");
            dayDiv.textContent = day;
            dayDiv.classList.add("day");

            const dateKey = `${year}-${month + 1}-${day}`;
            if (selectedDates.has(dateKey)) {
                dayDiv.classList.add("selected");
            }

            dayDiv.addEventListener("click", () => {
                if (selectedDates.has(dateKey)) {
                    selectedDates.delete(dateKey);
                    dayDiv.classList.remove("selected");
                } else {
                    selectedDates.set(dateKey, []);
                    dayDiv.classList.add("selected");
                }
            });

            daysContainer.appendChild(dayDiv);
        }
    }

    function updateCalendar() {
        generateCalendar(currentMonth, currentYear);
    }

    prev.addEventListener("click", () => {
        currentMonth--;
        if (currentMonth < 0) {
            currentMonth = 11;
            currentYear--;
        }
        updateCalendar();
    });

    next.addEventListener("click", () => {
        currentMonth++;
        if (currentMonth > 11) {
            currentMonth = 0;
            currentYear++;
        }
        updateCalendar();
    });

    document.getElementById("add-transpired").addEventListener("click", () => {
        const newTranspired = document.createElement("div");
        newTranspired.innerHTML = `<label for="transpired">What has Transpired:</label><input type="text" class="transpired"><button class="remove-transpired">Remove</button>`;
        transpiredContainer.appendChild(newTranspired);

        newTranspired.querySelector(".remove-transpired").addEventListener("click", () => {
            transpiredContainer.removeChild(newTranspired);
        });
    });

    document.getElementById("add-ipcr").addEventListener("click", () => {
        const newIpcr = document.createElement("div");
        newIpcr.innerHTML = `<label for="ipcr">IPCR Code:</label><input type="text" class="ipcr"><button class="remove-ipcr">Remove</button>`;
        ipcrContainer.appendChild(newIpcr);

        newIpcr.querySelector(".remove-ipcr").addEventListener("click", () => {
            ipcrContainer.removeChild(newIpcr);
        });
    });

    exportButton.addEventListener("click", () => {
        const transpiredElements = transpiredContainer.querySelectorAll(".transpired");
        const ipcrElements = ipcrContainer.querySelectorAll(".ipcr");
        const employeeNumber = employeeNumberInput.value;

        const data = [];
        selectedDates.forEach((value, key) => {
            transpiredElements.forEach(transpired => {
                ipcrElements.forEach(ipcr => {
                    data.push({
                        Date: key,
                        'What has Transpired': transpired.value,
                        'IPCR Code': ipcr.value,
                        'Employee Number': employeeNumber
                    });
                });
            });
        });

        data.sort((a, b) => new Date(a.Date) - new Date(b.Date));

        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

        // Wrap text style
        const wrapTextStyle = { alignment: { wrapText: true } };

        // Apply wrap text style to "What has Transpired" column
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        for (let row = range.s.r; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: 1 }); // Column B is index 1
            if (!worksheet[cellAddress]) worksheet[cellAddress] = {};
            worksheet[cellAddress].s = wrapTextStyle;
        }

        XLSX.writeFile(workbook, "interactive_calendar.xlsx");
    });

    updateCalendar();
});
