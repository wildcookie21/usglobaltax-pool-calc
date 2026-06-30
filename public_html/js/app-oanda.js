const oandaFileInput = document.getElementById('oanda-file');

oandaFileInput.addEventListener('change', function () {
  if (this.files.length > 0) {
    this.form.submit();
  }
});