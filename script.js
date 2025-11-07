document.getElementById("processBtn").addEventListener("click", async function() {
  const fileInput = document.getElementById("fileInput");
  const file = fileInput.files[0];
  const status = document.getElementById("status");

  if (!file) {
    alert("Lütfen bir Word dosyası yükleyin!");
    return;
  }

  status.innerText = "⏳ Dosya işleniyor, lütfen bekleyin...";

  const margin = parseFloat(document.getElementById("margin").value);
  const font = document.getElementById("font").value;
  const fontSize = parseInt(document.getElementById("fontSize").value);
  const lineSpacing = parseFloat(document.getElementById("lineSpacing").value);
  const pageNumbers = document.getElementById("pageNumbers").value;

  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer });
  const content = result.value;

  document.getElementById("previewContainer").classList.remove("hidden");
  document.getElementById("preview").innerHTML = content;

  const { Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType } = docx;

  const paragraphs = content
    .replace(/<[^>]+>/g, "\n")
    .split("\n")
    .filter(p => p.trim().length > 0)
    .map(p => new Paragraph({
      spacing: { line: lineSpacing * 240 },
      alignment: AlignmentType.JUSTIFIED,
      children: [
        new TextRun({
          text: p.trim(),
          font: font,
          size: fontSize * 2
        })
      ]
    }));

  let header, footer;
  if (pageNumbers === "top") {
    header = new Header({
      children: [new Paragraph({ text: "Sayfa [NUM]", alignment: AlignmentType.RIGHT })],
    });
  } else if (pageNumbers === "bottom") {
    footer = new Footer({
      children: [new Paragraph({ text: "Sayfa [NUM]", alignment: AlignmentType.CENTER })],
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
            right: margin * 567
          }
        }
      },
      children: paragraphs,
    }]
  });

  const blob = await Packer.toBlob(doc);
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "bicimlendirilmis_" + file.name;
  link.click();

  status.innerText = "✅ Dosya biçimlendirildi ve indirildi!";
});
