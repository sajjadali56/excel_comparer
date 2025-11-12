let pairCount = 1;

document.getElementById("addRow").addEventListener("click", () => {
  const container = document.getElementById("filePairs");
  const div = document.createElement("div");
  div.classList.add("card", "pair-card", "shadow-sm", "p-3");
  div.innerHTML = `
  <div class="text-end">
    <button type="button" class="btn btn-danger btn-sm remove-btn pair-card-close">
      <i class="fas fa-close"></i>
    </button>
  </div>
  <div class="mb-3">
    <label class="form-label fw-semibold text-secondary">Actual File:</label>
    <input type="file" name="actual_${pairCount}" class="form-control" accept=".xlsx, .xls" required>
  </div>
  <div class="mb-3">
    <label class="form-label fw-semibold text-secondary">Expected File:</label>
    <input type="file" name="expected_${pairCount}" class="form-control" accept=".xlsx, .xls" required>
  </div>
  <div id="error-${pairCount}" class="text-center badge bg-danger m-3 p-2" style="display: none">
  </div>
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
  console.log("Form submitted");
 
  const loader = document.getElementById("loader");
  const resultsSection = document.getElementById("results-section");
  const form = document.getElementById("uploadForm");
  
  loader.style.display = "block";
  resultsSection.style.display = "none";

  const body = new FormData(form);

  const values = {};

  body.forEach((value, key) => {
    values[key] = value;
  })
  const keys = Object.keys(values);
  console.log("Form data", values, keys);

  for(let i=0; i<keys.length/2; i++){
    const actualFile = values[`actual_${i}`];
    const expectedFile = values[`expected_${i}`];
    
    console.log(actualFile, expectedFile);
    if((!actualFile || !expectedFile) || (actualFile.size === 0 || expectedFile.size === 0)){
      const errorDiv = document.getElementById(`error-${i}`);
      errorDiv.innerHTML = "Please select both files";
      errorDiv.style.display = "block";
      loader.style.display = "none";
      resultsSection.style.display = "block";
      return
    }
  }
  
  const response = await fetch("/process", {
    method: "POST",
    body: body
  });

  const data = await response.json();

  console.log("Response received", data);

  const header = `<h3 class="text-info mb-3"> <i class="fas fa-chart-bar me-2"></i>Comparison Results</h3>`

  if(data.error){
    resultsSection.innerHTML = `<div class="alert alert-danger" role="alert">${data.error}</div>`
    loader.style.display = "none";
    resultsSection.style.display = "block";
    return
  }

  let resultCards = ""
  for (let i = 0; i < data.length; i++) {
    let sheetsHTML = ""
    for (let j = 0; j < data[i].results.length; j++) {
      sheetsHTML += `
        <details class="mb-3">
          <summary>${data[i].results[j].sheet}</summary>
          <pre class="mt-2">${data[i].results[j].message}</pre>
        </details>
      `
    }
    resultCards += `
      <div class="card results-card p-3">
        <div
          class="d-flex justify-content-between align-items-center mb-2 text-info fw-semibold"
        >
          <span><i class="fas fa-file-contract me-2"></i>${data[i].pair}</span>
          <a
            href="{{ url_for('download_file', folder='reports', filename=${data[i].report_file}) }}"
            class="btn btn-sm btn-outline-primary"
          >
            <i class="fas fa-download"></i> Download
          </a>
        </div>
        ${sheetsHTML}
      </div>
    `
  }
  const resultsHTML = `${header}${resultCards}`

  resultsSection.innerHTML = resultsHTML

  loader.style.display = "none";
  resultsSection.style.display = "block";
};

let descriptionVisible = false;
const toggleDescription = () => {
  const desc = document.getElementById("toolDesc");

  const getButton = (icon) => `<button class="btn btn-link text-white-50 p-0 ms-1 align-baseline" onclick="" title="Show more details" style="text-decoration: none;"><i class="fas fa-${icon}"></i></button>`;
  const text = "This tool automatically compares two Excel workbooks — “Actual” and “Expected” — across all common sheets. It checks both text and numeric columns, highlights statistical differences, and generates a clean, downloadable report for each pair.";
  if (!descriptionVisible) {
    desc.innerHTML = text + getButton("chevron-up");
  } else {
    desc.innerHTML = "This tool automatically compares two Excel workbooks" + getButton("chevron-down");
  }

  descriptionVisible = !descriptionVisible;
};

