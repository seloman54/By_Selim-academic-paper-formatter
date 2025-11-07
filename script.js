document.getElementById("processBtn").addEventListener("click", function() {
  const file = document.getElementById("fileInput").files[0];
  if (!file) {
    alert("Lütfen bir Word dosyası yükleyin!");
    return;
  }

  const margin = document.getElementById("margin").value;
  const font = document.getElementById("font").value;
  const fontSize = document.getElementById("fontSize").value;
  const lineSpacing = document.getElementById("lineSpacing").value;

  alert(`Dosya biçimlendirilecek:\n
  Kenar boşluğu: ${margin} cm
  Yazı tipi: ${font}
  Punto: ${fontSize}
  Satır aralığı: ${lineSpacing}
  
  (Bu sadece ön gösterimdir. Biçimlendirme işlemi bir sonraki sürümde yapılacak.)`);
});
