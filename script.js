/* ====== PDF.js worker ====== */
pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

/* ====== helpers UI ====== */
const $ = (s) => document.querySelector(s);
const setStatus = (m) => { const el = $("#status"); if (el) el.textContent = m; };

/* ====== Utils (1:1 com Python) ====== */
function ajustar_valor_BR(valor) {
  if (valor == null || valor === "") return "0";
  let v = String(valor).trim();
  let neg = "";
  if (v.startsWith("-")) { neg = "-"; v = v.slice(1); }
  v = v.replace(/,/g, "TEMP").replace(/\./g, ",").replace(/TEMP/g, ".");
  return neg + v;
}


function us_to_float(s) { return parseFloat(String(s).replace(/,/g, "")); }
function br2num(s) { return parseFloat(String(s).replace(/\./g,"").replace(",", ".")); }
function num2br(n) { return Number(n).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2}); }

/* ====== Regex do Python ====== */
const RX_PAD_INICIO  = /^(\d{1,3})\s+(TKTT|\+?EMD|EMDA|EMDS|RTDN|RFND|ADM[A-Z]*|ACM[A-Z]*|CANX|SPCR|SPDR|ADNT)\s+(\d{10})/;
const RX_TOTALIZADOR = /\b(ISSUES\s+TOTAL|DEBIT\s+MEMOS\s+TOTAL|CREDIT\s+MEMOS\s+TOTAL|AGENT\s+TOTALS?|GRAND\s+TOTAL|BSP\s+TOTALS?)\b/i;
const RX_NUM_US_G    = /-?\d{1,3}(?:,\d{3})*\.\d{2}/g;
const RX_NUM_US      = /-?\d{1,3}(?:,\d{3})*\.\d{2}/;

/* ====== MAPA CIA -> NOME_CIA ====== */
const MAP_CIA = {
  '001': 'AA','005':'CO','006':'DL','014':'AC','016':'UA',
  '037':'US','045':'LA','047':'TP','055':'AZ','057':'AF',
  '064':'OK','074':'KL','075':'IB','077':'MS','080':'LO',
  '081':'QF','083':'SA','086':'NZ','105':'AY','117':'SK',
  '125':'BA','127':'G3','131':'JL','139':'AM','142':'KF',
  '160':'CX','165':'JP','180':'KE','182':'MA','205':'NH',
  '217':'TG','220':'LH','230':'CM','235':'TK','257':'OS',
  '353':'NU','462':'XL','469':'4M','512':'RJ','544':'LP',
  '555':'SU','577':'AD','618':'SQ','680':'JK','688':'EG',
  '706':'KQ','708':'JO','724':'LX','774':'FM','831':'OU',
  '957':'JJ','988':'OZ','996':'UX','999':'CA','428':'JC',
  '134':'AV','118':'DT','044':'AR','930':'OB','176':'EK',
  '169':'HR','157':'QR','071':'ET','147':'AT','605':'H2',
  '973':'JA','275':'GP','799':'BNS','JC':'JC','952':'SPCR'
};
function nomeCIA(cia) {
  if (cia == null) return "";
  const raw = String(cia).trim();
  const k3 = raw.padStart(3, "0");
  // tenta como 3 dígitos, depois como chave literal (ex.: 'JC')
  return MAP_CIA[k3] ?? MAP_CIA[raw] ?? raw;
}

/* ====== Leitura do PDF: sem “consertos” adicionais ====== */
async function lerPDF(file) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

  window.__PAGES__ = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();

    // itens em ordem de leitura
    const items = tc.items.map((it) => {
      const [, , , , x, y] = it.transform;
      return { x, y, w: it.width || 0, str: it.str };
    }).sort((a,b)=> b.y - a.y || a.x - b.x);

    // agrupar por linha — tolerância pequena para não fundir linhas
    const epsY = 1.6;
    const linhas = [];
    for (const it of items) {
      let g = linhas.find(L => Math.abs(L.y - it.y) <= epsY);
      if (!g) { g = { y: it.y, itens: [] }; linhas.push(g); }
      g.itens.push(it);
    }
    linhas.forEach(L => L.itens.sort((a,b)=> a.x - b.x));

    // juntar itens com espaço simples (sem colar números/pontos/vírgulas)
    const linhasTexto = linhas.map(L =>
      L.itens.map(t => t.str).join(" ")
        .replace(/\u00A0/g," ")
        .replace(/\s+/g," ")
        .trim()
    ).filter(Boolean);

    window.__PAGES__.push(linhasTexto);
  }
}

/* ====== extrair IATA por página ====== */
function extrair_iata(linhas_pagina) {
  const texto = (linhas_pagina||[]).join("\n");
  let m = texto.match(/\b(\d{2})\s*-\s*(\d)\s+(\d{4})\s+(\d)\b/);
  if (m) return `${m[1]}-${m[2]} ${m[3]} ${m[4]}`;
  m = texto.match(/\b(\d{2})\s+(\d)\s+(\d{4})\s+(\d)\b/);
  if (m) return `${m[1]}-${m[2]} ${m[3]} ${m[4]}`;
  m = texto.match(/\b(\d{8})\b/);
  if (m) { const s = m[1]; return `${s.slice(0,2)}-${s.slice(2,3)} ${s.slice(3,7)} ${s.slice(7)}`; }
  return "";
}

/* ====== construir blocos (idêntico) ====== */
function construir_blocos() {
  const blocos = [];
  let bloco_atual = [];
  let linha_asterisco_buffer = null;
  let CURRENT_IATA = "";

  
  for (const linhas_pagina of window.__PAGES__ || []) {
    linha_asterisco_buffer = null; // zera por página
    const iata_pagina = extrair_iata(linhas_pagina);
    if (iata_pagina) CURRENT_IATA = iata_pagina;

    for (const linha of linhas_pagina) {
      if (RX_TOTALIZADOR.test(linha)) {
        if (bloco_atual.length) { bloco_atual.__IATA = CURRENT_IATA; blocos.push(bloco_atual); }
        bloco_atual = [];
        linha_asterisco_buffer = null;
        continue;
      }
      const mIni = linha.match(RX_PAD_INICIO);
      if (mIni) {
        if (bloco_atual.length) { bloco_atual.__IATA = CURRENT_IATA; blocos.push(bloco_atual); }
        bloco_atual = [linha];
        if (linha_asterisco_buffer && !RX_TOTALIZADOR.test(linha_asterisco_buffer)) {
          bloco_atual.push(linha_asterisco_buffer);
        }
        linha_asterisco_buffer = null;
        continue;
      }
      if (linha.includes("**")) {
        if (/ISSUES/i.test(linha) || RX_TOTALIZADOR.test(linha)) {
          linha_asterisco_buffer = null; continue;
        }
        if (bloco_atual.length) bloco_atual.push(linha);
        else linha_asterisco_buffer = linha; // pendura para o próximo cabeçalho
        continue;
      }
      if (bloco_atual.length) bloco_atual.append?.(linha) ?? bloco_atual.push(linha);
    }
  }
  if (bloco_atual.length) { bloco_atual.__IATA = (bloco_atual.__IATA || ""); blocos.push(bloco_atual); }
  return blocos;
}

/* ====== +RTDN (1:1) ====== */
function extrair_rtdn(bloco) {
  const nums=[];
  for (let i=0;i<bloco.length;i++){
    const s = bloco[i];
    if (s.includes("+RTDN")) {
      const janela = bloco.slice(i,i+6).join(" ");
      for (const m of janela.matchAll(/\b(\d{10,13})\b/g)) {
        if (!nums.includes(m[1])) nums.push(m[1]);
      }
    }
  }
  return nums.join(" | ");
}

/* ====== TAXAS (1:1 com Python) ====== */
function extrair_taxas(bloco) {
  const CAB_RE = /^(?:\d{3}\s+)?(?:\+?TKTT|RTDN|RFND|\+?EMD[A-Z]*|ADM[A-Z]*|ACM[A-Z]*|CANX|SPCR|SPDR|ADNT)\s+\d{10}\b/;

  let inicio = null;
  for (let i=0;i<bloco.length;i++){
    if (/^\d{3}\s+(TKTT|EMD[A-Z]*|ADM[A-Z]*|ACM[A-Z]*|RTDN|CANX|RFND|SPCR|SPDR|ADNT)\s+\d{10}\b/.test(bloco[i])) { inicio = i; break; }
  }
  if (inicio === null) return "";

  let fim = bloco.length;
  for (let i=inicio+1;i<bloco.length;i++){
    const l = bloco[i];
    if (CAB_RE.test(l)) { fim = i; break; }
    if (/(ISSUES TOTAL|DEBIT MEMOS TOTAL|CREDIT MEMOS TOTAL|AGENT TOTAL|GRAND TOTAL|TOUR:|VAT:)/i.test(l)) { fim = i; break; }
    if (/^Page\s*:/i.test(l)) continue;
  }

  const linhas_bloco = bloco.slice(inicio, fim).map(s=>s.trim());

  let linhas_taxas = [];
  let coletando = false;
  for (const s of linhas_bloco) {
    if (/\*\*/.test(s)) {
      coletando = true;
      const pos = s.split("**", 2)[1]?.trim() ?? "";
      linhas_taxas.push(pos || s.replace("**", "").trim());
      continue;
    }
    if (coletando) {
      if (/(COMISS|INCENTIV|IMP\s|SALDO A PAGAR|COBL|VAT|TOUR:|GRAND TOTAL|AGENT TOTAL|ISSUES TOTAL)/i.test(s)) break;
      if (CAB_RE.test(s)) break;
      if (/(FCGRBILLDET|AGENT GROUP BILLING DETAILS|CIA\s+TRNC|VALOR\s+TRANSACAO|TARIFA|TAXAS, IMP)/i.test(s)) continue;
      if (s) linhas_taxas.push(s);
    }
  }
  if (!linhas_taxas.length) linhas_taxas = linhas_bloco;

  const num = "(\\d{1,3}(?:,\\d{3})*\\.\\d{2})";
  const padrao_sigla = "([A-Z]{1}\\d{2}|[A-Z]{2}\\d{1}|[A-Z]{1}\\d{1}|[A-Z]{3}|[A-Z]{2}|\\d{1,2}[A-Z]|[A-Z]{4}|[A-Z]{5}|[A-Z]{4}\\d{1}|[A-Z]{5}\\d{1})";
  const inval = new Set(["CA","CC","EX"]);

  const out = [];
  for (const s of linhas_taxas) {
    const vistos = new Set(); // por linha

    for (const m of s.matchAll(new RegExp(`${num}\\s+${padrao_sigla}\\b`, "g"))) {
      const valor = m[1], sig = m[2];
      if (!inval.has(sig) && /[A-Z]/.test(sig)) {
        const v = valor.replace(/,/g,"").replace(/\./g,",");
        const chave = `${v} ${sig}`;
        if (!vistos.has(chave)) { vistos.add(chave); out.push(chave); }
      }
    }
    for (const m of s.matchAll(new RegExp(`${num}${padrao_sigla}\\b`, "g"))) {
      const valor = m[1], sig = m[2];
      if (!inval.has(sig) && /[A-Z]/.test(sig)) {
        const v = valor.replace(/,/g,"").replace(/\./g,",");
        const chave = `${v} ${sig}`;
        if (!vistos.has(chave)) { vistos.add(chave); out.push(chave); }
      }
    }
    for (const sig of ["OBFCA","OBFCA0","BR4","OD","OBT02","OAAG","CA","CC","EX"]) {
      const rx = (sig==="CA"||sig==="CC"||sig==="EX")
        ? new RegExp(`${num}\\s*${sig}(?!\\d)\\b`, "g")
        : new RegExp(`${num}\\s*${sig}\\b`, "g");
      for (const m of s.matchAll(rx)) {
        const v = m[1].replace(/,/g,"").replace(/\./g,",");
        const chave = `${v} ${sig}`;
        if (!vistos.has(chave)) { vistos.add(chave); out.push(chave); }
      }
    }
    for (const m of s.matchAll(new RegExp(`${num}\\s*DU\\b`, "g"))) {
      const v = m[1].replace(/,/g,"").replace(/\./g,",");
      const chave = `${v} DU`;
      if (!vistos.has(chave)) { vistos.add(chave); out.push(chave); }
    }
  }
  return out.join(", ");
}

/* ====== pós-processo TAXAS/DU/YQ (iguais ao Python) ====== */
function somenteValoresTaxas(t) {
  if (!t) return [];
  const rx = /\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2}/g;
  return t.match(rx) || [];
}
function somaBR(lista) {
  let tot=0;
  for (const v of lista) tot += br2num(v);
  return Number(tot.toFixed(2));
}
function Extrair_DU_deTXT(t) {
  if (!t) return "";
  const rx = /(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})\s*DU\b/g;
  let tot=0, m; while((m=rx.exec(t))!==null) tot += br2num(m[1]);
  return tot ? num2br(tot) : "";
}
function extrair_YQ_deTXT(t) {
  if (!t) return "";
  const rx = /(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})\s*YQ\b/g;
  let tot=0, m; while((m=rx.exec(t))!==null) tot += br2num(m[1]);
  return tot ? num2br(tot) : "";
}

/* ====== processar_bloco (mantendo LIQUIDO intocado) ====== */
function processar_bloco(bloco, IATA_GLOBAL) {
  const linha_cabeca = bloco[0] || "";
  const texto = bloco.join(" ");

  // CIA: se não estiver no início, faz fallback no texto
  let cia = "";
  const cia_m = linha_cabeca.match(/^(\d{1,3})/);
  if (cia_m) cia = cia_m[1];
  else {
    const c2 = texto.match(/\b(\d{1,3})\s+(TKTT|\+?EMD|EMDA|EMDS|RTDN|RFND|ADM[A-Z]*|ACM[A-Z]*|CANX|SPCR|SPDR|ADNT)\s+\d{10}\b/);
    if (c2) cia = c2[1];
  }

  const tipo_m = linha_cabeca.match(/\b(TKTT|\+?EMD|EMDA|EMDS|RTDN|RFND|ADM[A-Z]*|ACM[A-Z]*|CANX|SPCR|SPDR|ADNT)\b/);
  const no_trnc = (tipo_m ? tipo_m[1] : "").replace(/^\+/, "").toUpperCase();

  const doc_m = (linha_cabeca.match(/\b(\d{10})\b/) || texto.match(/\b(\d{10})\b/));
  const documento = doc_m ? doc_m[1] : "";

  const dm = texto.match(/\b\d{2}\/[A-Za-z]{3}\/\d{2}\b/);
  const data_emissao = dm ? dm[0] : "";

  const cpn_m = texto.match(/\b[A-Z]{4}(?:\s+NR[:=]\s*[A-Z0-9]+)?\b/);
  const cpn = cpn_m ? cpn_m[0] : "";

  const stat = /\bI\b/.test(linha_cabeca) ? "I" : (/\bD\b/.test(linha_cabeca) ? "D" : "");
  const fop_m = texto.match(/\b(CC|CA|EX)\b/);
  const fop = fop_m ? fop_m[0] : "";

  const pad_num = /-?\d{1,3}(?:,\d{3})*\.\d{2}/;
  let valor_transacao = "0", tarifa = "0", imp_comiss = "0", liquido = "0";

  const _eh_totalizador = (l) => RX_TOTALIZADOR.test(l);

  let idx_ast = null;
  for (let i=0;i<bloco.length;i++){
    const l = bloco[i];
    if (l.includes("**") && !/ISSUES/i.test(l) && !_eh_totalizador(l) && pad_num.test(l)) { idx_ast = i; break; }
  }

  if (idx_ast !== null && (no_trnc === "TKTT" || no_trnc.startsWith("EMD"))) {
    const apos = bloco[idx_ast].split("**", 2)[1]?.trim() ?? "";
    const nums = [...apos.matchAll(RX_NUM_US_G)].map(m=>m[0]);
    if (nums.length >= 1) valor_transacao = ajustar_valor_BR(nums[0]);
    if (nums.length >= 2) tarifa          = ajustar_valor_BR(nums[1]);
    if (nums.length >= 4) {
      imp_comiss = ajustar_valor_BR(nums[nums.length-2]);
      liquido    = ajustar_valor_BR(nums[nums.length-1]);
    } else { imp_comiss = "0"; liquido = "0"; }

  } else if (no_trnc === "TKTT") {
    const m_ft = linha_cabeca.match(new RegExp(`\\b(CC|CA|EX)\\b\\s+(${RX_NUM_US.source})\\s+(${RX_NUM_US.source})`));
    if (m_ft) {
      valor_transacao = ajustar_valor_BR(m_ft[2]);
      tarifa          = ajustar_valor_BR(m_ft[3]);
    }
    const header_nums = [...(linha_cabeca.matchAll(RX_NUM_US_G))].map(m=>m[0]);
    if (header_nums.length >= 2) {
      imp_comiss = ajustar_valor_BR(header_nums[header_nums.length-2]);
      liquido    = ajustar_valor_BR(header_nums[header_nums.length-1]);
    }

  } else if (idx_ast !== null && (no_trnc==="RFND" || no_trnc.startsWith("ADM") || no_trnc.startsWith("ACM"))) {
    const apos = bloco[idx_ast].split("**", 2)[1]?.trim() ?? "";
    const nums = [...apos.matchAll(RX_NUM_US_G)].map(m=>m[0]);
    if (nums.length >= 1) valor_transacao = ajustar_valor_BR(nums[0]);
    if (nums.length >= 2) tarifa          = ajustar_valor_BR(nums[1]);
    if (nums.length >= 2) liquido         = ajustar_valor_BR(nums[nums.length-1]);
    if (nums.length >= 3) imp_comiss      = ajustar_valor_BR(nums[nums.length-2]);

  } else if (no_trnc==="RFND" || no_trnc==="ADMA" || no_trnc.startsWith("ACM")) {
    const m = linha_cabeca.match(new RegExp(`\\b(CC|CA|EX)\\b\\s+(${RX_NUM_US.source})\\s+(${RX_NUM_US.source})`));
    if (m) {
      valor_transacao = ajustar_valor_BR(m[2]);
      tarifa          = ajustar_valor_BR(m[3]);
    } else {
      const hdr_nums = [...(linha_cabeca.matchAll(RX_NUM_US_G))].map(m=>m[0]);
      if (hdr_nums.length >= 2) {
        valor_transacao = ajustar_valor_BR(hdr_nums[0]);
        tarifa          = ajustar_valor_BR(hdr_nums[1]);
      }
    }
    const hdr = [...(linha_cabeca.matchAll(RX_NUM_US_G))].map(m=>m[0]);
    if (hdr.length >= 1) liquido    = ajustar_valor_BR(hdr[hdr.length-1]);
    if (hdr.length >= 2) imp_comiss = ajustar_valor_BR(hdr[hdr.length-2]);

  } else {
    const hdr = [...(linha_cabeca.matchAll(RX_NUM_US_G))].map(m=>m[0]);
    if (hdr.length >= 1) liquido    = ajustar_valor_BR(hdr[hdr.length-1]);
    if (hdr.length >= 2) imp_comiss = ajustar_valor_BR(hdr[hdr.length-2]);

    const todos = [...(texto.matchAll(RX_NUM_US_G))].map(m=>m[0]);
    for (let i=0;i<todos.length-1;i++){
      const v1 = us_to_float(todos[i]);
      const v2 = us_to_float(todos[i+1]);
      if (v1 >= v2) { valor_transacao = ajustar_valor_BR(todos[i]); tarifa = ajustar_valor_BR(todos[i+1]); break; }
    }
  }

  const taxas = extrair_taxas(bloco);
  const rt = extrair_rtdn(bloco);

  return {
    "CIA": cia,
    "NO TRNC": no_trnc,
    "DOCUMENT": documento || "",    // sem apóstrofo aqui
    "DT DE EMISSAO": data_emissao,
    "CPN": cpn,
    "NR": "",
    "CODIGO": "",
    "STAT": stat,
    "FOP": fop,
    "VALOR TRANSACAO": valor_transacao,
    "TARIFA": tarifa,
    "IMP COMISS": imp_comiss,
    "LIQUIDO": liquido,             // não tocar (está correto)
    "TAXAS": taxas,
    "IATA": IATA_GLOBAL || "",
    "+RT": rt
  };
}

/* ====== extrair_dados_completos + formatação final ====== */
function extrair_dados_completos() {
  const blocos = construir_blocos();
  const dados = [];

  for (const bloco of blocos) {
    const iata = bloco.__IATA || "";
    dados.push(processar_bloco(bloco, iata));
  }

  // Pós-processo (iguais ao Python) + NOME_CIA
  const out = dados.map(row => {
    const CIA = String(row["CIA"]||"");
    const DOC = String(row["DOCUMENT"]||"");
    const BILHETE = "'" + (CIA.padStart(3,"0") + DOC);
    const DU = Extrair_DU_deTXT(row["TAXAS"]);
    const YQ = extrair_YQ_deTXT(row["TAXAS"]);
    const Total_taxa = num2br(somaBR(somenteValoresTaxas(row["TAXAS"])));
    const NOME_CIA = nomeCIA(CIA);

    return {
      "IATA": row["IATA"],
      "CIA": CIA,
      "NOME_CIA": NOME_CIA,
      "NO TRNC": row["NO TRNC"],
      "DOCUMENT": "'" + DOC,     // apóstrofo apenas aqui (como no Python)
      "BILHETE": BILHETE,
      "VALOR TRANSACAO": row["VALOR TRANSACAO"],
      "TARIFA": row["TARIFA"],
      "IMP COMISS": row["IMP COMISS"],
      "LIQUIDO": row["LIQUIDO"],
      "Total taxa": Total_taxa,
      "DU": DU,
      "YQ": YQ,
      "TAXAS": row["TAXAS"],
      "+RT": row["+RT"]
    };
  });

  return out;
}

/* ====== Export ====== */
function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),0);
}
function exportarExcel(rows, nomeBase="fatura_bsp_extraida") {
  const header = ["IATA","CIA","NOME_CIA","NO TRNC","DOCUMENT","BILHETE","VALOR TRANSACAO","TARIFA","IMP COMISS","Total taxa","LIQUIDO","DU","YQ","TAXAS","RT"];
  const sane = (rows?.length ? rows : [Object.fromEntries(header.map(k=>[k,""]))])
    .map(r => Object.fromEntries(header.map(k => [k, r[k] ?? ""])));
  try {
    const ws = XLSX.utils.json_to_sheet(sane, { header });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Resultados");
    const arr = XLSX.write(wb, { bookType:"xlsx", type:"array", compression:true });
    triggerDownload(new Blob([arr], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
                    `${nomeBase}.xlsx`);
    setStatus(`Arquivo gerado: ${nomeBase}.xlsx`);
  } catch {
    const linhas = [header.join(";")];
    for (const r of sane) linhas.push(header.map(k => String(r[k]).replaceAll(";"," ,")).join(";"));
    const csv = "\ufeff" + linhas.join("\r\n");
    triggerDownload(new Blob([csv], {type:"text/csv;charset=utf-8;"}), `${nomeBase}.csv`);
    setStatus(`Arquivo gerado: ${nomeBase}.csv`);
  }
}

/* ====== UI wiring ====== */
let ARQ = null, RESULTADO = [];
function initUI(){
  const input = $("#fileInput");
  const drop  = $("#dropZone");
  const btnLer = $("#btnLer");
  const btnExport = $("#btnExportar");

  input?.addEventListener("change", ()=> {
    ARQ = input.files?.[0] ?? null;
    setStatus(ARQ ? `Arquivo: ${ARQ.name}` : "Nenhum arquivo selecionado.");
    if (btnLer) btnLer.disabled = !ARQ;
    if (btnExport) btnExport.disabled = true;
    RESULTADO = [];
  });

  ["dragenter","dragover"].forEach(ev =>
    drop?.addEventListener(ev, e=>{e.preventDefault();e.stopPropagation();drop.classList.add("dragover");})
  );
  ["dragleave","drop"].forEach(ev =>
    drop?.addEventListener(ev, e=>{e.preventDefault();e.stopPropagation();drop.classList.remove("dragover");})
  );
  drop?.addEventListener("drop", e=>{
    const f = e.dataTransfer.files?.[0];
    if (f && f.type==="application/pdf") {
      ARQ = f; if (input) input.files = e.dataTransfer.files;
      setStatus(`Arquivo: ${f.name}`);
      if (btnLer) btnLer.disabled = false; if (btnExport) btnExport.disabled = true;
      RESULTADO = [];
    } else { alert("Selecione um PDF."); }
  });
  drop?.addEventListener("click", ()=> input?.click());

  btnLer?.addEventListener("click", async ()=>{
    if (!ARQ) return;
    btnLer.disabled = true; if (btnExport) btnExport.disabled = true;
    setStatus("Lendo e processando o PDF…");
    try {
      await lerPDF(ARQ);
      RESULTADO = extrair_dados_completos();

      // conferência rápida no console (opcional)
      const somaCol = (k)=> RESULTADO.reduce((a,r)=> a + br2num(String(r[k]||"0")), 0);
      console.log("TOTAL VALOR TRANSACAO:", num2br(somaCol("VALOR TRANSACAO")));
      console.log("TOTAL TARIFA        :", num2br(somaCol("TARIFA")));
      console.log("TOTAL TAXAS         :", num2br(somaCol("Total taxa")));
      console.log("TOTAL LIQUIDO       :", num2br(somaCol("LIQUIDO")));

      setStatus(`Registros: ${RESULTADO.length}. Pronto para exportar.`);
      if (btnExport) btnExport.disabled = false;
    } catch (e) {
      console.error(e);
      setStatus("Erro ao processar o PDF.");
      alert("Falha ao ler o PDF.");
    } finally {
      btnLer.disabled = !ARQ;
    }
  });

  btnExport?.addEventListener("click", ()=> exportarExcel(RESULTADO));
}
document.addEventListener("DOMContentLoaded", initUI);
