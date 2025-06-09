# Offline OCR Word Add-in

This sample demonstrates how to build a Word add-in that extracts text from images using OCR and sends the recognized text to a local LLM such as GPT4ALL or Ollama. The add-in loads a simple task pane where you can choose an image and then process it.

## Setup

1. Install [Node.js](https://nodejs.org/) and a simple HTTP server such as `http-server`.
2. Download `office.js` from <https://appsforoffice.microsoft.com/lib/1/hosted/office.js> and place it in this folder.
3. Download `tesseract.min.js` from the [Tesseract.js releases](https://github.com/naptha/tesseract.js/) page and place it here as well so it can be loaded offline.
4. Serve the `word-addin` folder locally (for example with `npx http-server`) and update the URLs in `manifest.xml` to match your local server.
5. Start your local LLM server (for example `gpt4all-lora` or `ollama serve`).
6. Sideload the add-in manifest into Word for testing. See [Microsoft documentation](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) for sideloading steps.

## Files

- `manifest.xml` – The Office add-in manifest.
- `taskpane.html` – UI for the task pane.
- `taskpane.js` – Handles OCR using Tesseract.js and sends the extracted text to the local LLM endpoint.
- `tesseract.min.js` – The Tesseract.js library file to load offline.

## Running

1. Ensure your local LLM server is running. `taskpane.js` sends recognized text to `http://localhost:4891/v1/chat/completions`, which should be served by your offline model.
2. Start your web server so that the add-in files are reachable at the URLs listed in the manifest.
3. In Word, sideload `manifest.xml`.
4. Choose an image from the task pane and click **Extract Text**. The recognized text appears in the pane after processing by your local model.

Edit `taskpane.js` if your LLM server uses a different endpoint or model name.
