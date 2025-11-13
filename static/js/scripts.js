// JavaScript code with dark theme adjustments
let pairCount = 1;

document.getElementById("addRow").addEventListener("click", () => {
  const container = document.getElementById("filePairs");
  const div = document.createElement("div");
  div.classList.add("card", "pair-card", "glass-card", "p-3", "mb-3");
  div.innerHTML = `
    <div class="text-end">
        <button type="button" class="btn btn-danger btn-sm remove-btn">
            <i class="fas fa-times"></i>
        </button>
    </div>
    <div class="mb-3">
        <label class="form-label">
            <i class="fas fa-file-upload me-2"></i>Actual File
        </label>
        <input type="file" name="actual_${pairCount}" class="form-control" accept=".xlsx, .xls" required>
    </div>
    <div class="mb-3">
        <label class="form-label">
            <i class="fas fa-file-download me-2"></i>Expected File
        </label>
        <input type="file" name="expected_${pairCount}" class="form-control" accept=".xlsx, .xls" required>
    </div>
    <div id="error-${pairCount}" class="text-center badge bg-danger m-2 p-2" style="display: none"></div>
            `;
  container.appendChild(div);
  pairCount++;
  updateRemoveButtons();
});

function updateRemoveButtons() {
  document.querySelectorAll(".remove-btn").forEach((btn, idx) => {
    btn.disabled = idx === 0;
    btn.onclick = () => !btn.disabled && btn.closest(".pair-card").remove();
  });
}

updateRemoveButtons();

const handleSubmit = async (event) => {
  event.preventDefault();

  const loader = document.getElementById("loader");
  const resultsSection = document.getElementById("results-section");
  const form = document.getElementById("uploadForm");

  loader.style.display = "block";
  resultsSection.style.display = "none";
  resultsSection.innerHTML = "";

  const body = new FormData(form);

  // Validate files
  let hasErrors = false;
  for (let i = 0; i < pairCount; i++) {
    const actualFile = body.get(`actual_${i}`);
    const expectedFile = body.get(`expected_${i}`);
    const errorDiv = document.getElementById(`error-${i}`);

    if (errorDiv) {
      errorDiv.style.display = "none";
    }

    if (
      !actualFile ||
      !expectedFile ||
      actualFile.size === 0 ||
      expectedFile.size === 0
    ) {
      if (errorDiv) {
        errorDiv.innerHTML = "Please select both files";
        errorDiv.style.display = "block";
      }
      hasErrors = true;
    }
  }

  if (hasErrors) {
    loader.style.display = "none";
    return;
  }

  try {
    const response = await fetch("/process", {
      method: "POST",
      body: body,
    });

    const data = await response.json();
    displayResults(data);
  } catch (error) {
    resultsSection.innerHTML = `
      <div class="alert alert-danger" role="alert">
          <i class="fas fa-exclamation-triangle me-2"></i>
          Error: ${error.message}
      </div>
    `;
  } finally {
    loader.style.display = "none";
    resultsSection.style.display = "block";
  }
};

document.getElementById("uploadForm").addEventListener("submit", handleSubmit);

function displayResults(data) {
  const resultsSection = document.getElementById("results-section");

  if (data.error) {
    resultsSection.innerHTML = `
    <div class="alert alert-danger" role="alert">
        <i class="fas fa-exclamation-triangle me-2"></i>
        ${data.error}
    </div>
  `;
    return;
  }

  if (data.length === 0) {
    resultsSection.innerHTML = `
    <div class="alert alert-warning text-center" role="alert">
        <i class="fas fa-info-circle me-2"></i>
        No valid file pairs found for comparison.
    </div>
  `;
    return;
  }

  let resultsHTML = `
    <h3 class="text-primary mb-4">
        <i class="fas fa-chart-bar me-2"></i>Comparison Results
    </h3>
  `;

  data.forEach((pair, pairIndex) => {
    const results = pair.results;

    if (results.error) {
      resultsHTML += `
        <div class="alert alert-warning">
            <i class="fas fa-exclamation-triangle me-2"></i>
            ${results.error}
        </div>
      `;
      return;
    }

    // Overall statistics
    const totalSheets = results.total_sheets || 0;
    let totalColumns = 0;
    let matchingColumns = 0;
    let differentColumns = 0;

    results.sheets?.forEach((sheet) => {
      totalColumns += sheet.total_columns || 0;
      matchingColumns += sheet.matching_columns || 0;
      differentColumns += sheet.different_columns || 0;
    });

    const cardHeader = `
      <div class="card-header bg-primary text-white border-0">
        <div class="d-flex justify-content-between align-items-center">
          <div class="d-flex align-items-between">
            <h5 class="mb-0">
                <i class="fas fa-file-contract me-2"></i>
                ${pair.pair}
            </h5>
            <small class="ms-2 badge bg-info">${results.comparison_time || "Unknown time"}</small>
          </div>
          <div class="d-flex align-items-center">

            <!-- Download Button -->
            <a href="/download/reports/${pair.report_file}" class="btn btn-sm me-2">
                <i class="fa-solid fa-file-pdf me-2"></i>
            </a>
            ${pair.has_pdf ? `
              <a href="/download/reports/${pair.json_report_file}" class="btn btn-sm">
                  <i class="fas fa-code me-2"></i>
              </a>` : ''
            }
          </div>
        </div>
      </div>    
    `;

    const stats = [
      { label: "Total Sheets", value: totalSheets, icon:"fa-layer-group", color:"" },
      { label: "Total Columns", value: totalColumns, icon:"fa-columns", color:"warning" },
      { label: "Matching Columns", value: matchingColumns, icon:"fa-check-circle", color: "success" },
      { label: "Different Columns", value: differentColumns, icon:"fa-times-circle", color: "danger" },
    ]
    const summaryStats = `
      <div class="row mb-4">
        ${stats.map((stat) => `
          <div class="col-md-3">
              <div class="stat-card ${stat.color} text-center">
                  <i class="fas ${stat.icon} fa-2x mb-2"></i>
                  <h4>${stat.value}</h4>
                  <small>${stat.label}</small>
              </div>
          </div>  
        `).join("")}
    </div>`;

    const getSheetSummary = (sheet) => `
      <div class="sheet-summary p-3 border border-secondary rounded" onclick="toggleSheetDetails(${pairIndex}, '${sheet.sheet_name}')">
        <div class="d-flex justify-content-between align-items-center">
          <h6 class="mb-0 text-light">
            <i class="fas fa-table me-2"></i> ${sheet.sheet_name}
            <span class="badge bg-secondary ms-2">${sheet.total_columns} columns</span>
          </h6>
          <div>
              <span class="badge bg-success">${sheet.matching_columns || 0} matching</span>
              <span class="badge bg-danger">${sheet.different_columns || 0} different</span>
              <i class="fas fa-chevron-down ms-2"></i>
          </div>
        </div>
    </div>`;

    const getSheetDetails = (sheet) => `
    <div id="sheet-${pairIndex}-${sheet.sheet_name}" class="sheet-details mt-3" style="display: none;">
      ${sheet.columns?.map((column) => `
        <div class="column-comparison p-3 mb-2 rounded ${
          column.status === "different"? "column-diff": "column-match"}">
          <div class="d-flex justify-content-between align-items-start">
              <div>
                  <strong class="text-light">${column.name}</strong>
                  <span class="badge ${column.type === "numeric" ? "bg-info": "bg-warning"} ms-2">${column.type}</span>
                  <span class="badge ${column.status === "different" ? "bg-danger" : "bg-success"} ms-1">${column.status}</span>
              </div>
              <small class="text-light mt-1">Click to expand</small>
          </div>
          <div class="column-details mt-2" style="display: none;">
              ${renderColumnDetails(
                column
              )}
          </div>
        </div>
    `).join("") ||'<p class="text-light mt-1">No columns to display</p>'
    }
  </div>
`;

    resultsHTML += `
      <div class="card results-card glass-card mb-4">
        ${cardHeader}
        <div class="card-body">
          <!-- Summary Statistics -->
          ${summaryStats}

          <!-- Sheets Comparison -->
          <div class="sheets-comparison">
              ${results.sheets?.map((sheet) => `
                <div class="sheet-section mb-4">
                  ${getSheetSummary(sheet)}
                  ${getSheetDetails(sheet)}
                </div>`).join("") ||'<p class="text-light mt-1">No sheets to display</p>'
              }
          </div>
        </div>
      </div>
  `;
  });

  resultsSection.innerHTML = resultsHTML;

  // Add click handlers for column details
  document.querySelectorAll(".column-comparison").forEach((column) => {
    column.addEventListener("click", function () {
      const details = this.querySelector(".column-details");
      details.style.display =
        details.style.display === "none" ? "block" : "none";
    });
  });
}

function renderColumnDetails(column) {
  if (column.status === "matching") {
    if (column.type === "numeric" && column.statistics) {
      return `
        <div class="row text-light mt-1">
            ${Object.entries(column.statistics)
              .map(
                ([stat, values]) => `
                <div class="col-md-6">
                    <small><strong>${stat.toUpperCase()}:</strong> ${values.file1}</small>
                </div>
            `
              )
              .join("")}
        </div>
      `;
    }
    return '<small class="text-success">✓ All values match perfectly</small>';
  } else {
    if (column.type === "numeric" && column.differences) {
      return `
        <div class="table-responsive">
            <table class="table table-dark table-sm table-bordered">
                <thead>
                    <tr>
                        <th>Statistic</th>
                        <th>File 1 Value</th>
                        <th>File 2 Value</th>
                        <th>Difference</th>
                    </tr>
                </thead>
                <tbody>
                    ${column.differences
                      .map(
                        (diff) => `
                        <tr>
                            <td><strong>${
                              diff.statistic
                            }</strong></td>
                            <td>${diff.file1_value}</td>
                            <td>${diff.file2_value}</td>
                            <td class="text-danger">${
                              diff.difference > 0 ? "+" : ""
                            }${diff.difference}</td>
                        </tr>
                    `
                      )
                      .join("")}
                </tbody>
            </table>
        </div>`;
    } else if (column.differences) {
      return `
        <div class="table-responsive">
            <table class="table table-dark table-sm table-bordered">
                <thead>
                    <tr>
                        <th>Value</th>
                        <th>File 1 Count</th>
                        <th>File 2 Count</th>
                        <th>Difference</th>
                    </tr>
                </thead>
                <tbody>
                    ${column.differences
                      .map(
                        (diff) => `
                        <tr>
                            <td><code>${diff.value}</code></td>
                            <td>${diff.file1_count}</td>
                            <td>${diff.file2_count}</td>
                            <td class="text-danger">${
                              diff.file2_count -
                              diff.file1_count
                            }</td>
                        </tr>
                    `
                      )
                      .join("")}
                </tbody>
            </table>
        </div>`;
    }
  }
  return '<small class="text-light mt-1">No detailed information available</small>';
}

// Utility functions
function toggleSheetDetails(pairIndex, sheetName) {
  const details = document.getElementById(`sheet-${pairIndex}-${sheetName}`);
  details.style.display = details.style.display === "none" ? "block" : "none";
}

function scrollToTop() {
  window.scrollTo({ top: 0, behavior: "smooth" });
}

// Show/hide scroll to top button
window.onscroll = function () {
  const scrollBtn = document.getElementById("scrollToTop");
  if (
    document.body.scrollTop > 100 ||
    document.documentElement.scrollTop > 100
  ) {
    scrollBtn.style.display = "flex";
  } else {
    scrollBtn.style.display = "none";
  }
};

// Initialize tooltips
const tooltipTriggerList = [].slice.call(
  document.querySelectorAll('[data-bs-toggle="tooltip"]')
);
const tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
  return new bootstrap.Tooltip(tooltipTriggerEl);
});
