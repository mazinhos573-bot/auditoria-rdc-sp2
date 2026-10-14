
const XLSX = require("xlsx");
const crypto = require("node:crypto");

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const ARQUIVO = process.env.ARQUIVO_EXCEL || "baixado_hoje.xlsx";
const TABELA = "recebimentos_imile";

function normalizarCabecalho(valor) {
  return String(valor ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function valor(linha, nome) {
  const alvo = normalizarCabecalho(nome);
  const chave = Object.keys(linha).find(
    (k) => normalizarCabecalho(k) === alvo
  );
  return chave ? linha[chave] : null;
}

function texto(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  return s === "" ? null : s;
}

function dataScan(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return v.toISOString().slice(0, 19);
  }

  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) {
      return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")} ${String(d.H).padStart(2, "0")}:${String(d.M).padStart(2, "0")}:${String(Math.floor(d.S)).padStart(2, "0")}`;
    }
  }

  const s = texto(v);
  if (!s) return null;

  const m = s.match(
    /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/
  );

  return m
    ? `${m[1]}-${m[2]}-${m[3]} ${m[4]}:${m[5]}:${m[6]}`
    : s;
}

function chaveUnica(dados) {
  const base = [
    dados.scan_time,
    dados.waybill_number,
    dados.bag_number,
    dados.uuid,
    dados.scan_type,
    dados.scan_station_code,
    dados.operator
  ].map((v) => v ?? "").join("|");

  return crypto.createHash("sha256").update(base).digest("hex");
}

async function main() {
  if (!SUPABASE_URL || !SUPABASE_KEY) {
    throw new Error(
      "Configure SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY nas variáveis de ambiente."
    );
  }

  const workbook = XLSX.readFile(ARQUIVO, { cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const linhas = XLSX.utils.sheet_to_json(sheet, { defval: null });

  if (!linhas.length) {
    throw new Error("O Excel não contém linhas para importar.");
  }

  const primeira = linhas[0];
  const obrigatorias = ["Waybill Number", "Scan Time"];
  for (const coluna of obrigatorias) {
    if (!Object.keys(primeira).some(
      (k) => normalizarCabecalho(k) === normalizarCabecalho(coluna)
    )) {
      throw new Error(`Coluna obrigatória ausente: ${coluna}`);
    }
  }

  const registros = [];
  let ignorados = 0;

  for (const linha of linhas) {
    const waybill = texto(valor(linha, "Waybill Number"));
    if (!waybill) {
      ignorados++;
      continue;
    }

    const dados = {
      scan_time: dataScan(valor(linha, "Scan Time")),
      waybill_number: waybill,
      bag_number: texto(valor(linha, "Bag Number")),
      uuid: texto(valor(linha, "UUID")),
      scan_type: texto(valor(linha, "Scan Type")),
      scan_station: texto(valor(linha, "Scan Station")),
      scan_station_code: texto(valor(linha, "Scan Station Code")),
      operator: texto(valor(linha, "Operator")),
      client_name: texto(valor(linha, "Client Name")),
      route_code: texto(valor(linha, "Route Code")),
      arrival_type: texto(valor(linha, "Arrival Type")),
      raw_data: linha
    };

    registros.push({
      ...dados,
      dedupe_key: chaveUnica(dados)
    });
  }

  const endpoint =
    `${SUPABASE_URL.replace(/\/$/, "")}/rest/v1/${TABELA}`;

  let enviados = 0;

  for (let i = 0; i < registros.length; i += 500) {
    const lote = registros.slice(i, i + 500);

    const resposta = await fetch(
      `${endpoint}?on_conflict=dedupe_key`,
      {
        method: "POST",
        headers: {
          apikey: SUPABASE_KEY,
          Authorization: `Bearer ${SUPABASE_KEY}`,
          "Content-Type": "application/json",
          Prefer: "resolution=ignore-duplicates,return=minimal"
        },
        body: JSON.stringify(lote)
      }
    );

    if (!resposta.ok) {
      const erro = await resposta.text();
      throw new Error(
        `Falha no lote ${Math.floor(i / 500) + 1}: ${resposta.status} ${erro}`
      );
    }

    enviados += lote.length;
    console.log(`Lote processado: ${enviados}/${registros.length}`);
  }

  console.log("Importação concluída.");
  console.log(`Linhas válidas enviadas: ${enviados}`);
  console.log(`Linhas sem waybill ignoradas: ${ignorados}`);
  console.log("Registros já existentes são ignorados pela chave única.");
}

main().catch((erro) => {
  console.error("Erro na importação:", erro.message);
  process.exitCode = 1;
});
