document
  .getElementById("studentAllotmentForm")
  .addEventListener("submit", async function (event) {
    event.preventDefault();
    document.documentElement.scrollTop = 0;

    const fileInput = document.getElementById("fileInput");
    const roomCapacitiesInput = document.getElementById("roomCapacitiesInput");
    const roomCapacities = roomCapacitiesInput.value
      .split(",")
      .map((capacity) => parseInt(capacity.trim(), 10));

    const file = fileInput.files[0];

    // Check if a file and valid room capacities are provided
    if (
      !file ||
      roomCapacities.some((capacity) => isNaN(capacity) || capacity <= 0)
    ) {
      alert(
        "Please upload a valid Excel file and enter valid room capacities!"
      );
      return;
    }

    try {
      // Parse the Excel file to get the student IDs
      let studentIds = await parseExcel(file);
      studentIds = sortStudentIds(studentIds);
      const rowCapacity = 3;
      const rooms = allotSeats(studentIds, roomCapacities, rowCapacity);
      displayTableResults(rooms);
      document.getElementById("result").style.display = "block";
    } catch (error) {
      console.error(error);
      alert("Error reading the file. Please try again.");
    }
  });

// Parse the uploaded Excel file
function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const studentIds = XLSX.utils
        .sheet_to_json(worksheet, { header: 1 })
        .flat()
        .filter((id) => typeof id === "string" && id.trim() !== "");
      resolve(studentIds);
    };
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
}
// Sort student IDs
function sortStudentIds(studentIds) {
  return studentIds.sort((a, b) => {
    const [yearA, branchA, numberA] = a
      .split("-")
      .map((part) => (isNaN(part) ? part : parseInt(part)));
    const [yearB, branchB, numberB] = b
      .split("-")
      .map((part) => (isNaN(part) ? part : parseInt(part)));

    if (yearA !== yearB) return yearA - yearB;
    if (branchA !== branchB) return branchA.localeCompare(branchB);
    return numberA - numberB;
  });
}

// Group students by branch
function groupStudentsByBranch(studentIds) {
  const grouped = {};
  studentIds.forEach((studentId) => {
    const branchCode = extractBranchCode(studentId);
    if (!grouped[branchCode]) {
      grouped[branchCode] = [];
    }
    grouped[branchCode].push(studentId);
  });
  return grouped;
}

// Extract branch code
function extractBranchCode(studentId) {
  return studentId.split("-")[1];
}

// Alternate students from different branches
function alternateBranches(groupedStudents) {
  const alternated = [];
  const branchKeys = Object.keys(groupedStudents);
  let index = 0;

  while (branchKeys.some((branch) => groupedStudents[branch].length > 0)) {
    const currentBranch = branchKeys[index % branchKeys.length];
    if (groupedStudents[currentBranch].length > 0) {
      alternated.push(groupedStudents[currentBranch].shift());
    }
    index++;
  }

  return alternated;
}

// Allot seats to rooms based on room capacities and fixed row capacity
function allotSeats(studentIds, roomCapacities, rowCapacity) {
  const groupedStudents = groupStudentsByBranch(studentIds);
  const alternatedStudents = alternateBranches(groupedStudents);

  const rooms = [];
  let currentStudentIndex = 0;

  roomCapacities.forEach((capacity, index) => {
    const roomStudents = alternatedStudents.slice(
      currentStudentIndex,
      currentStudentIndex + capacity
    );
    if (roomStudents.length > 0) {
      const rows = [];
      const totalRows = Math.ceil(roomStudents.length / rowCapacity);

      for (let i = 0; i < totalRows; i++) {
        const row = [];
        for (let j = 0; j < rowCapacity; j++) {
          const studentIndex = i + j * totalRows;
          if (studentIndex < roomStudents.length) {
            row.push(roomStudents[studentIndex]);
          } else {
            row.push("");
          }
        }
        rows.push(row);
      }

      rooms.push({
        name: `Room ${index + 1}`,
        rows: rows,
      });
    }
    currentStudentIndex += capacity;
  });

  return rooms;
}

// Display the seating arrangement in rows of three
function displayTableResults(rooms) {
  const tableResults = document.getElementById("tableResults");
  tableResults.innerHTML = "";

  rooms.forEach((room) => {
    if (room.rows.length > 0) {
      const table = document.createElement("table");
      table.classList.add("roomTable");

      const thead = document.createElement("thead");
      const headerRow = document.createElement("tr");
      const roomHeader = document.createElement("th");
      roomHeader.setAttribute("colspan", 3);
      roomHeader.textContent = `${room.name}`;
      headerRow.appendChild(roomHeader);
      thead.appendChild(headerRow);
      table.appendChild(thead);

      const tbody = document.createElement("tbody");

      // Loop through each row and create table rows
      room.rows.forEach((row) => {
        const studentRow = document.createElement("tr");

        row.forEach((studentId) => {
          const studentCell = document.createElement("td");
          studentCell.textContent = studentId;
          studentRow.appendChild(studentCell);
        });

        // If the row has less than three students, fill the empty cells
        while (studentRow.children.length < 3) {
          const emptyCell = document.createElement("td");
          emptyCell.textContent = "";
          studentRow.appendChild(emptyCell);
        }

        tbody.appendChild(studentRow);
      });

      table.appendChild(tbody);
      tableResults.appendChild(table);
    }
  });

  // Create download buttons
  const excelButton = document.createElement("button");
  excelButton.textContent = "Download Excel";
  excelButton.onclick = () => downloadExcel(rooms);
  tableResults.appendChild(excelButton);
}

// Function to download results as Excel
function downloadExcel(rooms) {
  const workbook = XLSX.utils.book_new();

  rooms.forEach((room) => {
    const sheetData = [];
    // Fill in student IDs
    room.rows.forEach((row) => {
      sheetData.push(row);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(workbook, worksheet, room.name);
  });

  XLSX.writeFile(workbook, "seat_allotment.xlsx");
}
