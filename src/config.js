// Configuração lida do ambiente. O .env é carregado aqui, antes de qualquer leitura.
require("dotenv").config();

const TOPDESK_URL   = (process.env.TOPDESK_URL || "").replace(/\/$/, "");
const TOPDESK_EMAIL = process.env.TOPDESK_EMAIL || "";
const TOPDESK_TOKEN = process.env.TOPDESK_TOKEN || "";

// SharePoint / Azure AD
const AZURE_TENANT_ID     = process.env.AZURE_TENANT_ID || "";
const AZURE_CLIENT_ID     = process.env.AZURE_CLIENT_ID || "";
const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET || "";
const SHAREPOINT_SITE_URL = (process.env.SHAREPOINT_SITE_URL || "").replace(/\/$/, "");
const SHAREPOINT_FOLDER   = (process.env.SHAREPOINT_FOLDER || "").replace(/^\/|\/$/g, "");
const SHAREPOINT_FILENAME = process.env.SHAREPOINT_FILENAME || "Nao Devolvidos.xlsx";

// Relatório separado dos ativos sem CC (não recuperado nem no asset, nem no histórico).
// Vai para um site/pasta próprio para que a equipe possa preencher manualmente.
// Sem SHAREPOINT_EXAURIDOS_SITE_URL o arquivo é gerado só localmente.
const SHAREPOINT_EXAURIDOS_SITE_URL = (process.env.SHAREPOINT_EXAURIDOS_SITE_URL || "").replace(/\/$/, "");
const SHAREPOINT_EXAURIDOS_FOLDER   = (process.env.SHAREPOINT_EXAURIDOS_FOLDER || "").replace(/^\/|\/$/g, "");
const SHAREPOINT_EXAURIDOS_FILENAME = process.env.SHAREPOINT_EXAURIDOS_FILENAME || "Nao_Devolvidos_Sem_CC.xlsx";

// IDs da instância TOPdesk (templates de ativo, status e pessoa usada pela automação
// que abre os chamados). Variam por instância, então vêm do ambiente.
const TEMPLATE_NOTEBOOK    = process.env.TOPDESK_TEMPLATE_NOTEBOOK || "";
const TEMPLATE_CELULAR     = process.env.TOPDESK_TEMPLATE_CELULAR || "";
const STATUS_NAO_DEV       = process.env.TOPDESK_STATUS_NAO_DEVOLVIDO || "";
const AUTOMATION_CALLER_ID = process.env.TOPDESK_AUTOMATION_CALLER_ID || "";

// Sem estes valores as consultas saem com filtro vazio. O script para antes de
// chamar a API, em vez de varrer ativos e histórico com parâmetros errados.
const REQUIRED_ENV = {
  TOPDESK_URL, TOPDESK_EMAIL, TOPDESK_TOKEN,
  TOPDESK_TEMPLATE_NOTEBOOK: TEMPLATE_NOTEBOOK,
  TOPDESK_TEMPLATE_CELULAR: TEMPLATE_CELULAR,
  TOPDESK_STATUS_NAO_DEVOLVIDO: STATUS_NAO_DEV,
  TOPDESK_AUTOMATION_CALLER_ID: AUTOMATION_CALLER_ID,
};

function missingEnv(required = REQUIRED_ENV) {
  return Object.entries(required).filter(([, v]) => !v).map(([k]) => k);
}

module.exports = {
  TOPDESK_URL, TOPDESK_EMAIL, TOPDESK_TOKEN,
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET,
  SHAREPOINT_SITE_URL, SHAREPOINT_FOLDER, SHAREPOINT_FILENAME,
  SHAREPOINT_EXAURIDOS_SITE_URL, SHAREPOINT_EXAURIDOS_FOLDER, SHAREPOINT_EXAURIDOS_FILENAME,
  TEMPLATE_NOTEBOOK, TEMPLATE_CELULAR, STATUS_NAO_DEV, AUTOMATION_CALLER_ID,
  missingEnv,
};
