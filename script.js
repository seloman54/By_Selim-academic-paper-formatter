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

  // Word içeriğini HTML'e dönüştür
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer });
  const content = result.value; // HTML formatında içerik

  // Şimdi yeni docx dosyası oluştur
  const { Document, Packer, Paragraph } = docx;

  // HTML içeriğini paragraflara dönüştür
  const paragraphs = content
    .replace(/<[^>]+>/g, "\n") // HTML etiketlerini temizle
    .split("\n")
    .filter(p => p.trim().length > 0)
    .map(p => new Paragraph({
      spacing: { line: lineSpacing * 240 },
      children: [
        new docx.TextRun({
          text: p.trim(),
          font: font,
          size: fontSize * 2,
        })
      ],
    }));

  // Yeni Word dokümanını oluştur
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: margin * 567,
            bottom: margin * 567,
            left: margin * 567,
            right: margin * 567,
          },
        },
      },
      children: paragraphs,
    }],
  });

  // Dosyayı indir
  const blob = await Packer.toBlob(doc);
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "bicimlendirilmis_" + file.name;
  link.click();

  alert("✅ Dosya biçimlendirildi ve indirildi!");
});
