# Multi‑client Word Add‑in for Legal Assistance

This repository contains a **multi‑tenant** implementation of the legal assistant Word add‑in.  It
allows you to provide the same add‑in to multiple clients, each with its own
OpenAI API key, without exposing those keys in the front‑end.  A single
backend instance serves requests for all clients by authenticating callers via
a simple `clientId`/`clientSecret` mechanism.

The project is split into two parts:

* **`backend/`** – A Node.js API that receives requests from the Word task pane,
  authenticates the client, calls the OpenAI API on behalf of that client,
  and returns the generated text.
* **`frontend/`** – A task pane implemented as an HTML/JavaScript page with a
  chat‑style user interface.  It reads the `clientId` and `clientSecret` from
  the query string, sends them to the backend, and inserts AI responses into
  the document.
* **`manifest-template.xml`** – A base manifest for the Word add‑in.  You
  generate one manifest per client by replacing `TASKPANE_URL` with the
  appropriate URL containing that client's credentials.

## Getting started

### 1. Configure client records

Edit `backend/clients.json` and add an entry for each client.  Each entry
must include:

* `id` – A unique identifier for the client (e.g. `cliente123`).
* `secret` – A pre‑shared secret used to authenticate the client.  Choose
  a random string and share it privately with that client.  **Do not embed
  your OpenAI API key in the manifest.**
* `openaiApiKey` – The OpenAI API key to use when serving that client's
  requests.

Example:

```json
[
  {
    "id": "cliente123",
    "secret": "minha-senha-secreta",
    "openaiApiKey": "sk-live-abcdefg..."
  },
  {
    "id": "cliente456",
    "secret": "outra-senha",
    "openaiApiKey": "sk-live-hijklmn..."
  }
]
```

### 2. Start the backend

Install dependencies and run the server in the `backend` folder:

```bash
cd backend
npm install
cp .env.example .env
# Set any variables in .env as needed (for example PORT or fallback OPENAI_API_KEY)
npm start
```

By default, the backend listens on port `4000`.  In production you should
deploy it behind HTTPS and a reverse proxy.  The `/health` endpoint can be
used for monitoring.

### 3. Serve the frontend

The task pane consists of static files in `frontend/`.  You can serve this
folder with any web server (e.g. Nginx, Apache, Vercel, Azure Static Web
Apps).  Make sure the server uses HTTPS in production.  The backend API
may be hosted on the same domain or on a separate subdomain.

For local development you can run a simple static server on port `3000`:

```bash
cd frontend
npx serve -l 3000 .
```

### 4. Generate client‑specific manifests

Copy `manifest-template.xml` for each client and replace the following
placeholders:

* `REPLACE-WITH-UNIQUE-GUID` – Generate a new GUID for each manifest (for
  example using `npx uuid` or an online GUID generator).  Each add‑in needs
  a unique ID.
* `TASKPANE_URL` – Replace this with the full HTTPS URL to your
  `taskpane.html` **including** the client's `cid` and `token` as query
  parameters.  For example:

  ```xml
  <SourceLocation DefaultValue="https://ia.seu-dominio.com/taskpane.html?cid=cliente123&amp;token=minha-senha-secreta" />
  ```

Assign each manifest to the corresponding client.  You can distribute
manifests to clients in a variety of ways:

* **Centralized deployment** – If your clients belong to the same Microsoft 365
  tenant, you can upload the manifest to the admin center and assign it to
  specific users or groups.  See Microsoft's documentation on
  centralised deployment【189162168263982†L50-L69】.
* **AppSource publication** – For a public SaaS offering you can submit the
  add‑in to AppSource.  Customers will install it from the store and you can
  direct them to a configuration page to generate their credentials.
* **Manual sideload** – During development and testing, clients can sideload
  the manifest locally via the “My Add‑ins” menu.

### 5. Client workflow

1. The user opens Word and loads the add‑in.  The task pane is loaded
   from your `taskpane.html` and receives the `cid` and `token` from the
   URL.
2. When the user enters a prompt and clicks **Enviar**, the front‑end
   sends a request to the backend with the client credentials.
3. The backend validates the credentials, uses the client‑specific
   OpenAI API key to generate a response, and returns the text.
4. The front‑end displays the AI response in the chat and inserts it
   into the Word document according to the selected mode (`replace` or
   `after`).

### Security considerations

* **Protect secrets** – The `clientSecret` acts as a password for API
  requests.  Do not share it publicly.  The manifest includes this
  secret in the URL, so control access to the manifest and your front‑end.
  For greater security, consider implementing an authentication
  flow (e.g. OAuth or JWT) instead of passing secrets in the URL.
* **HTTPS** – Always serve both the front‑end and backend over HTTPS in
  production.  Office will refuse to load non‑HTTPS add‑ins by default.
* **Rate limiting and logging** – To prevent abuse, implement per‑client
  rate limiting and log usage.  This example does not include these
  features.

## Next steps

This repository is intended as a starting point.  For a full SaaS
implementation you might want to add:

* A web portal for clients to sign up, view usage and rotate their
  secrets and API keys.
* A database to store clients instead of a JSON file.
* Billing and subscription management.
* Per‑client model configuration (model, temperature, etc.).
* Improved UI/UX with additional prompt templates and command buttons.

Pull requests are welcome!