document.addEventListener('DOMContentLoaded', () => {
  const form = document.querySelector('form[action="cgt-upload.php"]');
  const input = document.getElementById('input-sheet-file');

  document.addEventListener('dragover', e => e.preventDefault());

  document.addEventListener('drop', e => {
    e.preventDefault();

    if (!e.dataTransfer.files.length) return;

    const dt = new DataTransfer();
    dt.items.add(e.dataTransfer.files[0]);

    input.files = dt.files;

    form.submit();
  });
});