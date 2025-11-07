document.getElementById("processBtn").addEventListener("click", async function() {
  const fileInput = document.getElementById("fileInput");
  const file = fileInput.files[0];
  if (!file) {
    alert("Lütfen bir Word dosyası yükleyin!");
    return;
  }

  const margin = parseFloat(document.getElementById("margin").value);
  const font = document.getElementById("font").value;
  const fontSize = parseInt(document.getElementById("fontSize").value);
  const lineSpacing = parseFloat(document.getElementById("lineSpacing").value);

  const reader = new FileReader();
  reader.onload = async function(e) {
    const arrayBuffer = e.target.result;

    const { Document, Packer, Paragraph, TextRun } = docx;

    // Yeni doküman oluştur
    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: margin * 567, // cm → twips
              bottom: margin * 567,
              left: margin * 567,
              right: margin * 567,
            },
          },
        },
        children: [
          new Paragraph({
            spacing: { line: lineSpacing * 240 },
            children: [
              new TextRun({
                text: "Bu biçimlendirme test dokümanıdır.\nYüklediğiniz dosyalar bu formatta düzenlenecektir.",
                font: font,
                size: fontSize * 2,
              }),
            ],
          }),
        ],
      }],
    });

    // Biçimlendirilmiş dosyayı oluştur ve indir
    const blob = await Packer.toBlob(doc);
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "bicimlendirilmis_dosya.docx";
    link.click();
  };

  reader.readAsArrayBuffer(file);
});
