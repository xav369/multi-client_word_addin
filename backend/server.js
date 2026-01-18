import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import { readFileSync, writeFileSync, existsSync } from "fs";
import path from "path";
import OpenAI from "openai";

// Load environment variables from .env if present
dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

// Serve static files from the frontend directory
app.use(express.static(path.join(process.cwd(), '../frontend')));

// Load the list of clients from clients.json. This simple JSON file
// contains an array of client objects with the properties:
//  - id: unique identifier for the client (string)
//  - secret: pre‑shared secret used to authenticate the client (string)
//  - openaiApiKey: API key to use when calling the OpenAI API on
//    behalf of this client (string)
//
// In a production system you should store this information in a database
// and rotate secrets periodically. For the sake of simplicity, this
// example reads the data at startup. If you modify the file, you must
// restart the server.
const clientsPath = path.join(process.cwd(), "clients.json");
let clients = [];
try {
  const raw = readFileSync(clientsPath, { encoding: "utf-8" });
  clients = JSON.parse(raw);
  if (!Array.isArray(clients)) {
    throw new Error("clients.json must be an array of clients");
  }
} catch (err) {
  console.error(`Could not read clients.json: ${err.message}`);
  clients = [];
}

// Load assistants custom prompts
const assistantsPath = path.join(process.cwd(), "assistants.json");
let assistants = {};
function loadAssistants() {
  try {
    if (existsSync(assistantsPath)) {
      const raw = readFileSync(assistantsPath, { encoding: "utf-8" });
      assistants = JSON.parse(raw);
    }
  } catch (err) {
    console.error(`Could not read assistants.json: ${err.message}`);
    assistants = {};
  }
}
// Initial load
loadAssistants();

function saveAssistant(clientId, optimizedPrompt) {
  // Reload first to ensure we don't overwrite other manual changes if possible (though race condition exists)
  loadAssistants();
  assistants[clientId] = optimizedPrompt;
  writeFileSync(assistantsPath, JSON.stringify(assistants, null, 2));
}

// Helper to find a client by id and secret. Returns the client record
// or undefined if not found or secret mismatch.
function getClientRecord(clientId, clientSecret) {
  if (!clientId || !clientSecret) {
    return undefined;
  }
  return clients.find(
    (c) => c.id === clientId && c.secret && c.secret === clientSecret
  );
}

// POST /api/ia
// Body: { clientId, clientSecret, prompt, selectedText, mode }
//
// This endpoint accepts a prompt and selected text from the Word add‑in,
// authenticates the client using the provided id and secret, and then
// calls the OpenAI API using the per‑client API key. The response is
// returned as plain text with an optional mode field to instruct the
// add‑in how to insert the returned text (e.g. replace or append).
app.post("/api/ia", async (req, res) => {
  try {
    const { clientId, clientSecret, prompt, selectedText, mode } = req.body;
    if (!clientId || !clientSecret) {
      return res.status(400).json({ error: "Missing clientId or clientSecret" });
    }
    const clientRecord = getClientRecord(clientId, clientSecret);
    if (!clientRecord) {
      return res.status(401).json({ error: "Invalid client credentials" });
    }
    if (!prompt) {
      return res.status(400).json({ error: "Missing prompt" });
    }

    // Compose the full prompt that will be sent to the OpenAI API.
    const fullPrompt = `\nVocê é um assistente jurídico brasileiro.\nTarefa: ${prompt}\nTexto de referência (se houver):\n${selectedText || "(nenhum texto selecionado)"
      }\nResponda apenas com o texto final, limpo, sem marcadores, sem asteriscos, sem emojis e pronto para ser utilizado em um documento Word (.docx).\n`.trim();

    // Determine which API key to use. Prefer the client‑specific key
    // from clients.json. If absent, fall back to the global OPENAI_API_KEY
    // from the environment. If neither is provided, respond with an error.
    const apiKey = clientRecord.openaiApiKey || process.env.OPENAI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: "No OpenAI API key configured for this client" });
    }

    // Create a fresh OpenAI client for this request. The OpenAI SDK is
    // lightweight and can be instantiated per request.
    const openai = new OpenAI({ apiKey });

    // Build messages with a system prompt to guide the model's tone and
    // content. Use a conservative temperature for deterministic output.

    // Reload assistants to pick up any manual changes
    loadAssistants();

    // Check if there is a custom prompt for this client
    const customSystemPrompt = assistants[clientId] ||
      "Você é um assistente jurídico brasileiro especializado em Direito Civil, Trabalhista e Previdenciário. Use linguagem técnica, clara e objetiva, conforme prática forense brasileira.";

    const messages = [
      {
        role: "system",
        content: customSystemPrompt,
      },
      { role: "user", content: fullPrompt },
    ];
    const completion = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages,
      temperature: 0.3,
    });
    // Safely extract the assistant's reply.  Use optional chaining
    // on array elements to avoid exceptions when the API response
    // structure changes or is undefined.
    let answer = 'Não foi possível gerar resposta.';
    if (completion && Array.isArray(completion.choices) && completion.choices.length > 0) {
      const choice = completion.choices[0];
      if (choice && choice.message && typeof choice.message.content === 'string') {
        answer = choice.message.content.trim();
      }
    }
    res.json({ text: answer, mode: mode || "replace" });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Erro interno" });
  }
});

// POST /api/configure
// Body: { clientId, clientSecret, rawPrompt }
//
// This endpoint takes a raw role description (e.g. "Civil Lawyer"),
// optimizes it using GPT-4o-mini into a detailed system prompt,
// and saves it for the client.
app.post("/api/configure", async (req, res) => {
  try {
    const { clientId, clientSecret, rawPrompt } = req.body;
    if (!clientId || !clientSecret) {
      return res.status(400).json({ error: "Missing clientId or clientSecret" });
    }
    const clientRecord = getClientRecord(clientId, clientSecret);
    if (!clientRecord) {
      return res.status(401).json({ error: "Invalid client credentials" });
    }
    if (!rawPrompt) {
      return res.status(400).json({ error: "Missing rawPrompt" });
    }

    const apiKey = clientRecord.openaiApiKey || process.env.OPENAI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: "No OpenAI API key configured" });
    }

    const openai = new OpenAI({ apiKey });

    // Optimize the prompt
    const optimizationMessages = [
      {
        role: "system",
        content: "Você é um especialista em Prompt Engineering para LLMs. Sua tarefa é transformar uma descrição curta de um papel profissional em um 'System Prompt' detalhado, robusto e otimizado para gerar documentos jurídicos de alta qualidade. O prompt otimizado deve instruir a IA a agir como esse profissional, usando terminologia correta e estilo formal. Responda APENAS com o prompt otimizado, sem explicações adicionais."
      },
      {
        role: "user",
        content: `Crie um System Prompt otimizado para esta descrição: "${rawPrompt}"`
      }
    ];

    const completion = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: optimizationMessages,
      temperature: 0.7
    });

    let optimizedPrompt = "Você é um assistente jurídico.";
    if (completion && completion.choices && completion.choices.length > 0) {
      const content = completion.choices[0].message.content;
      if (content) optimizedPrompt = content.trim();
    }

    // Save
    saveAssistant(clientId, optimizedPrompt);

    res.json({ success: true, optimizedPrompt });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Internal Error" });
  }
});

const EXPERT_PROMPT = `
Você é uma Autoridade Suprema em Direito Brasileiro e Linguística, atuando como o mais experiente Consultor Jurídico e Revisor de Textos.
Você detém conhecimento enciclopédico de todas as áreas do Direito (Civil, Penal, Trabalhista, Tributário, Constitucional, etc.) e domínio absoluto da Norma Culta da Língua Portuguesa.

Sua missão é realizar uma ANÁLISE FORENSE E LINGUÍSTICA DO DOCUMENTO fornecido.

Diretrizes de Análise:
1.  **Rigidez Jurídica**: Verifique a solidez dos argumentos, a correta aplicação dos institutos jurídicos e a vigência das leis citadas. Aponte fragilidades, riscos de nulidade ou teses ultrapassadas.
2.  **Excelência Linguística**: Identifique erros gramaticais, de sintaxe, pontuação e regência. Avalie a clareza, a coesão e a elegância do texto. O estilo deve ser culto, formal e persuasivo, sem ser pedante.
3.  **Estratégia Processual**: Analise se o texto atinge seu objetivo (convencer o juiz, notificar a parte, etc.) com eficácia.

Formato da Resposta:
Forneça um RELATÓRIO TÉCNICO estruturado:
-   **Resumo da Análise**: Uma visão geral da qualidade do documento.
-   **Pontos Críticos (Jurídico)**: Erros graves ou sugestões de fundamentação.
-   **Correções Linguísticas**: Lista de erros gramaticais ou de estilo encontrados.
-   **Sugestões de Melhoria**: Recomendações para elevar o nível do documento.

Se o documento estiver perfeito, reconheça a excelência. Se estiver ruim, seja implacável mas construtivo.
`;

// POST /api/analyze
// Body: { clientId, clientSecret, documentText }
//
// Performs a full analysis of the document using the Expert Prompt.
app.post("/api/analyze", async (req, res) => {
  try {
    const { clientId, clientSecret, documentText } = req.body;
    if (!clientId || !clientSecret) {
      return res.status(400).json({ error: "Missing clientId or clientSecret" });
    }
    const clientRecord = getClientRecord(clientId, clientSecret);
    if (!clientRecord) {
      return res.status(401).json({ error: "Invalid client credentials" });
    }
    if (!documentText) {
      return res.status(400).json({ error: "Missing documentText" });
    }

    const apiKey = clientRecord.openaiApiKey || process.env.OPENAI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: "No OpenAI API key configured" });
    }

    const openai = new OpenAI({ apiKey });

    const messages = [
      {
        role: "system",
        content: EXPERT_PROMPT
      },
      {
        role: "user",
        content: `Aqui está o texto do documento para análise:\n\n${documentText}`
      }
    ];

    const completion = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages,
      temperature: 0.4 // Lower temperature for more analytical/objective output
    });

    let analysis = "Não foi possível gerar a análise.";
    if (completion && completion.choices && completion.choices.length > 0) {
      const content = completion.choices[0].message.content;
      if (content) analysis = content.trim();
    }

    res.json({ analysis });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Internal Error" });
  }
});

// POST /api/save-key
// Body: { clientId, clientSecret, apiKey }
//
// Updates the client's record with a new OpenAI API key.
app.post("/api/save-key", (req, res) => {
  try {
    const { clientId, clientSecret, apiKey } = req.body;
    if (!clientId || !clientSecret) {
      return res.status(400).json({ error: "Missing clientId or clientSecret" });
    }
    if (!apiKey) {
      return res.status(400).json({ error: "Missing apiKey" });
    }

    const clientRecord = getClientRecord(clientId, clientSecret);
    if (!clientRecord) {
      return res.status(401).json({ error: "Invalid client credentials" });
    }

    // Update the key in memory and on disk
    // Note: In a real production app, use a secure vault.
    const clientsPath = path.join(process.cwd(), "clients.json");
    let allClients = [];
    if (existsSync(clientsPath)) {
      allClients = JSON.parse(readFileSync(clientsPath, "utf-8"));
    }

    const clientIndex = allClients.findIndex(c => c.id === clientId && c.secret === clientSecret);
    if (clientIndex >= 0) {
      allClients[clientIndex].openaiApiKey = apiKey;
      clients = allClients; // Update memory cache
      writeFileSync(clientsPath, JSON.stringify(allClients, null, 2));
      res.json({ success: true, message: "API Key salva com sucesso." });
    } else {
      // Should not happen if getClientRecord passed
      res.status(500).json({ error: "Client found in cache but not in file." });
    }
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Health check endpoint for monitoring
app.get("/health", (req, res) => {
  res.json({ ok: true });
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => {
  console.log(`Servidor IA multi‑cliente rodando em http://localhost:${PORT}`);
});