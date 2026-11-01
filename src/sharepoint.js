// Upload do relatório para o SharePoint via Microsoft Graph (client credentials).

const fs = require("fs");
const {
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET,
  SHAREPOINT_SITE_URL, SHAREPOINT_FOLDER, SHAREPOINT_FILENAME,
} = require("./config");

async function getGraphToken() {
  const url = `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type:    "client_credentials",
    client_id:     AZURE_CLIENT_ID,
    client_secret: AZURE_CLIENT_SECRET,
    scope:         "https://graph.microsoft.com/.default",
  });
  const res = await fetch(url, { method: "POST", body });
  if (!res.ok) throw new Error(`Erro ao obter token Azure: ${await res.text()}`);
  return (await res.json()).access_token;
}

async function uploadToSharePoint(localFilePath, opts = {}) {
  const siteUrlStr = opts.siteUrl  || SHAREPOINT_SITE_URL;
  const folder     = (opts.folder  ?? SHAREPOINT_FOLDER) || "";
  const filename   = opts.filename || SHAREPOINT_FILENAME;
  const label      = opts.label    || "SharePoint";

  const configured = AZURE_TENANT_ID && AZURE_CLIENT_ID && AZURE_CLIENT_SECRET && siteUrlStr;
  if (!configured) {
    console.log(`  [${label}] Variáveis não configuradas — upload ignorado.`);
    return;
  }

  console.log(`Enviando para ${label}...`);

  const token   = await getGraphToken();
  const headers = { Authorization: `Bearer ${token}` };

  const siteUrl = new URL(siteUrlStr);
  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteUrl.hostname}:${siteUrl.pathname}`,
    { headers }
  );
  if (!siteRes.ok) throw new Error(`Site ${label} não encontrado: ${await siteRes.text()}`);
  const { id: siteId } = await siteRes.json();

  const remotePath  = folder ? `${folder}/${filename}` : filename;
  const encodedPath = remotePath.split("/").map(encodeURIComponent).join("/");

  const fileContent = fs.readFileSync(localFilePath);
  const uploadRes   = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodedPath}:/content`,
    {
      method:  "PUT",
      headers: { ...headers, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
      body:    fileContent,
    }
  );
  if (!uploadRes.ok) {
    const errText = await uploadRes.text();
    if (uploadRes.status === 423 || errText.includes("resourceLocked")) {
      throw new Error(`Arquivo bloqueado no ${label} — feche-o no Excel ou SharePoint Online e tente novamente.`);
    }
    throw new Error(`Erro no upload (${label}): ${errText}`);
  }

  const { webUrl } = await uploadRes.json();
  console.log(`  ✓ Atualizado no ${label}: ${webUrl}`);
}

module.exports = { uploadToSharePoint };
