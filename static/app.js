(function () {
  const sampleCheckbox = document.querySelector('input[name="use_sample"]');
  const sampleNote = document.querySelector('.sample-note');
  const fileInputs = document.querySelectorAll('input[type="file"]');

  function toggleSampleNote() {
    if (!sampleNote) return;
    sampleNote.style.display = sampleCheckbox?.checked ? 'block' : 'none';
  }

  function updateFileSummary(input) {
    const summary = input.parentElement.querySelector('.file-summary');
    if (!summary) return;
    summary.textContent = input.files?.length
      ? `${input.files[0].name} 업로드됨`
      : '선택된 파일이 없습니다 (샘플 데이터 사용)';
  }

  toggleSampleNote();
  sampleCheckbox?.addEventListener('change', toggleSampleNote);

  fileInputs.forEach((input) => {
    updateFileSummary(input);
    input.addEventListener('change', () => updateFileSummary(input));
  });
})();
