const oandaFileInput = document.getElementById('input-sheet-file');

oandaFileInput.addEventListener('change', function () {
  if (this.files.length > 0) {
    this.form.submit();
  }
});