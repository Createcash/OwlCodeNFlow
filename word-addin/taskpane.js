Office.onReady(() => {
  document.getElementById('extractBtn').addEventListener('click', async () => {
    const fileInput = document.getElementById('imageInput');
    if (fileInput.files.length === 0) {
      return;
    }
    const file = fileInput.files[0];
    const base64 = await toBase64(file);
    const text = await runOCR(base64);
    const result = await callLocalLLM(text);
    document.getElementById('output').textContent = result;
  });
});

function toBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function runOCR(base64Image) {
  // OCR using a locally hosted copy of Tesseract.js
  const worker = Tesseract.createWorker();
  await worker.load();
  await worker.loadLanguage('eng');
  await worker.initialize('eng');
  const {
    data: { text }
  } = await worker.recognize(`data:image/png;base64,${base64Image}`);
  await worker.terminate();
  return text;
}

async function callLocalLLM(text) {
  const response = await fetch('http://localhost:4891/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: 'gpt4all',
      messages: [
        { role: 'user', content: text }
      ]
    })
  });
  const data = await response.json();
  return data.choices[0].message.content;
}
