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
  const pageNumbers = document.getElementById("pageNumbers").value;
  const template = document.getElementById("template").value;

  // Dosyayı oku
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer });
  const content = result.value;
  document.getElementById("previewContainer").classList.remove("hidden");
  document.getElementById("preview").innerHTML = content;

  const { Document, Packer, Paragraph, TextRun, Header, Footer } = docx;

  const paragraphs = content
    .replace(/<[^>]+>/g, "\n")
    .split("\n")
    .filter(p => p.trim().length > 0)
    .map(p => {
      let style = {};
      if (p.toLowerCase().includes("kaynakça") || p.toLowerCase().includes("references")) {
        style = { bold: true, underline: {} };
      }
      return new Paragraph({
        spacing: { line: lineSpacing * 240 },
        children: [
          new TextRun({
            text: p.trim(),
            font: font,
            size: fontSize * 2,
            ...style
          })
        ]
      });
    });

  let header, footer;
  if (pageNumbers === "top") {
    header = new Header({
      children: [new Paragraph({ text: "Sayfa: ", alignment: docx.AlignmentType.RIGHT })],
    });
  } else if (pageNumbers === "bottom") {
    footer = new Footer({
      children: [new Paragraph({ text: "Sayfa: ", alignment: docx.AlignmentType.CENTER })],
    });
  }

  const doc = new Document({
    sections: [{
      headers: { default: header },
      footers: { default: footer },
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

  const blob = await Packer.toBlob(doc);
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "bicimlendirilmis_" + file.name;
  link.click();

  alert("✅ Dosya biçimlendirildi ve indirildi!");
});
