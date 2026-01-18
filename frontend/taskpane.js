/*
 * taskpane.js
 *
 * Script for the multi‑tenant Word add‑in front‑end. This module reads
 * client identifiers from the query string, displays a chat‑like interface,
 * and communicates with the backend API to generate AI‑assisted legal text.
 */

// Parse query parameters from the current URL (e.g. ?cid=client1&token=secret1).
function getQueryParams() {
  const params = {};
  const query = window.location.search.substring(1);
  const vars = query.split('&');
  vars.forEach((v) => {
    const pair = v.split('=');
    if (pair.length === 2) {
      params[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1]);
    }
  });
  return params;
}

// When Office is ready, wire up UI events and check for client credentials.
Office.onReady(() => {
  const params = getQueryParams();
  const clientId = params.cid;
  const clientSecret = params.token;
  const statusEl = document.getElementById('status');

  if (!clientId || !clientSecret) {
    statusEl.textContent =
      'Erro: parâmetros cid e token ausentes na URL. Configure o manifest com ?cid=<id>&token=<secret>.';
    document.getElementById('sendButton').disabled = true;
    return;
  }

  // Setup event handler for the send button
  document.getElementById('sendButton').onclick = () => {
    runIA(clientId, clientSecret);
  };

  // Setup event handler for config button
  document.getElementById('configButton').onclick = () => {
    configureAssistant(clientId, clientSecret);
  };

  // Setup event handler for Save Key button
  const saveKeyBtn = document.getElementById('saveKeyButton');
  if (saveKeyBtn) {
    saveKeyBtn.onclick = () => {
      saveApiKeyOnly(clientId, clientSecret);
    };
  }

  // Setup event handler for analyze button
  const analyzeBtn = document.getElementById('analyzeButton');
  if (analyzeBtn) {
    analyzeBtn.onclick = () => {
      analyzeDocument(clientId, clientSecret);
    };
  }
});


// Append a message to the chat container. type is 'user' or 'ai'.
function appendMessage(content, type) {
  const chat = document.getElementById('chat');
  const messageDiv = document.createElement('div');
  messageDiv.className = `message ${type}`;
  messageDiv.textContent = content;
  chat.appendChild(messageDiv);
  chat.scrollTop = chat.scrollHeight;
}

// Runs the AI call: reads selected text, sends request, inserts reply.
async function runIA(clientId, clientSecret) {
  const statusEl = document.getElementById('status');
  const promptEl = document.getElementById('prompt');
  const mode = document.getElementById('mode').value;
  const prompt = (promptEl.value || '').trim();
  if (!prompt) {
    statusEl.textContent = 'Informe um comando para a IA.';
    return;
  }
  statusEl.textContent = '';

  // Show user message in chat
  appendMessage(prompt, 'user');
  promptEl.value = '';

  // Read selected text from Word
  let selectedText = '';
  try {
    selectedText = await getSelectedTextFromWord();
  } catch (err) {
    console.error(err);
    statusEl.textContent = 'Falha ao ler seleção do Word.';
    return;
  }

  // Compose request body
  const body = {
    clientId,
    clientSecret,
    prompt,
    selectedText,
    mode,
  };

  try {
    // Determine API base URL. Use relative path if served from same domain;
    // otherwise fall back to localhost:4000 in development.
    let apiBase = '';
    // If the front‑end is served from a different origin than the backend,
    // set API_BASE_URL in an injected script or environment variable.
    if (window.API_BASE_URL) {
      apiBase = window.API_BASE_URL;
    }
    const apiUrl = `${apiBase}/api/ia`;

    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!response.ok) {
      throw new Error(`Erro HTTP ${response.status}`);
    }
    const data = await response.json();
    const aiText = (data.text || '').trim();
    appendMessage(aiText, 'ai');

    // Optionally insert into Word
    if (mode !== 'chat') {
      try {
        await insertTextIntoWord(aiText, mode);
      } catch (err) {
        console.error(err);
        statusEl.textContent = 'Falha ao inserir texto no documento.';
      }
    }
  } catch (err) {
    console.error(err);
    statusEl.textContent = `Erro: ${err.message}`;
  }
}

// Read the currently selected text in the Word document.
async function getSelectedTextFromWord() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load('text');
    await context.sync();
    return selection.text || '';
  });
}

// Insert text into the Word document according to the specified mode.
async function insertTextIntoWord(text, mode) {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    if (mode === 'after') {
      selection.insertText('\n' + text + '\n', Word.InsertLocation.after);
    } else {
      selection.insertText(text, Word.InsertLocation.replace);
    }
    await context.sync();
  });
}

// Calls /api/configure to optimize and save the assistant prompt
async function configureAssistant(clientId, clientSecret) {
  const roleInput = document.getElementById('roleDescription');
  const resultDiv = document.getElementById('configResult');
  const rawPrompt = roleInput.value.trim();

  if (!rawPrompt) {
    resultDiv.textContent = "Por favor, descreva o papel do assistente.";
    return;
  }

  resultDiv.textContent = "Otimizando prompt... aguarde.";

  // Determine API base URL
  let apiBase = '';
  if (window.API_BASE_URL) {
    apiBase = window.API_BASE_URL;
  }
  const apiUrl = `${apiBase}/api/configure`;

  try {
    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ clientId, clientSecret, rawPrompt }),
    });

    if (!response.ok) {
      throw new Error(`Erro HTTP ${response.status}`);
    }

    const data = await response.json();
    if (data.success) {
      resultDiv.innerHTML = "<strong>Sucesso!</strong> Prompt otimizado:<br/><br/>" + data.optimizedPrompt;
      console.log("Optimized Prompt:", data.optimizedPrompt);
    } else {
      resultDiv.textContent = "Erro ao salvar configuração.";
    }
  } catch (err) {
    console.error(err);
    resultDiv.textContent = `Erro: ${err.message}`;
  }
}

// Saves ONLY the API Key
async function saveApiKeyOnly(clientId, clientSecret) {
  const keyInput = document.getElementById('apiKeyInput');
  const resultDiv = document.getElementById('configResult');
  const apiKey = keyInput.value.trim();

  if (!apiKey) {
    resultDiv.textContent = "Por favor, digite uma API Key válida.";
    return;
  }

  resultDiv.textContent = "Salvando API Key...";

  // Determine API base URL
  let apiBase = '';
  if (window.API_BASE_URL) {
    apiBase = window.API_BASE_URL;
  }

  try {
    const response = await fetch(`${apiBase}/api/save-key`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ clientId, clientSecret, apiKey }),
    });

    if (!response.ok) {
      throw new Error(`Erro HTTP ${response.status}`);
    }

    const data = await response.json();
    if (data.success) {
      resultDiv.textContent = "API Key salva com sucesso!";
      keyInput.value = ""; // Clear for security
      keyInput.placeholder = "API Key salva (oculta por segurança)";
    } else {
      resultDiv.textContent = "Erro ao salvar API Key.";
    }
  } catch (err) {
    console.error(err);
    resultDiv.textContent = `Erro ao salvar Key: ${err.message}`;
  }
}

// Full Document Analysis
async function analyzeDocument(clientId, clientSecret) {
  const statusEl = document.getElementById('status');
  // Clear status
  statusEl.textContent = "Lendo documento para análise...";

  // Read full document text
  let documentText = "";
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      documentText = body.text;
    });
  } catch (err) {
    console.error(err);
    statusEl.textContent = "Erro ao ler o documento.";
    return;
  }

  if (!documentText || !documentText.trim()) {
    statusEl.textContent = "O documento está vazio.";
    return;
  }

  statusEl.textContent = "Solicitando análise do especialista...";
  appendMessage("Iniciando análise completa do documento...", "user");

  // Determine API base URL
  let apiBase = '';
  if (window.API_BASE_URL) {
    apiBase = window.API_BASE_URL;
  }
  const apiUrl = `${apiBase}/api/analyze`;

  try {
    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ clientId, clientSecret, documentText }),
    });

    if (!response.ok) {
      throw new Error(`Erro HTTP ${response.status}`);
    }

    const data = await response.json();
    const analysis = data.analysis;

    // Display analysis in chat
    appendMessage(analysis, "ai");
    statusEl.textContent = "Análise concluída.";

  } catch (err) {
    console.error(err);
    statusEl.textContent = `Erro na análise: ${err.message}`;
    appendMessage(`Falha na análise: ${err.message}`, "ai");
  }
}