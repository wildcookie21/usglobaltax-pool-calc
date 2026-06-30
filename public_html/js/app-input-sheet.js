const inputSheetFileInput = document.getElementById('input-sheet-file');

inputSheetFileInput.addEventListener('change', function () {
  if (this.files.length > 0) {
    this.form.submit();
  }
});