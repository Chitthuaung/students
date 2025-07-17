document.getElementById("upload").addEventListener("change", function(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const html = XLSX.utils.sheet_to_html(sheet);
    document.getElementById("output").innerHTML = html;
  };

  reader.readAsArrayBuffer(file);
});