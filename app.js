/* =========================================================
   Extrator BSP (TXT) – Navegador
   Porta a lógica do Python para JavaScript (regex e etapas).
   ========================================================= */

const $ = (sel) => document.querySelector(sel);
const fileInput = $("#fileInput");
const btnProcess = $("#btnProcessar");
const btnCsv = $("#btnCsv");
const statusBox = $("#status");
const thead = $("#thead");
const tbody = $("#tbody");

let loadedFiles = [];     // { name, text }
let resultRows = [];     // array de objetos (resultado final)

fileInput.addEventListener("change", async (e) => {
    loadedFiles = [];
    const files = Array.from(e.target.files || []);
    if (!files.length) {
        btnProcess.disabled = true;
        btnCsv.disabled = true;
        setStatus("Nenhum arquivo carregado.");
        clearTable();
        return;
    }
    setStatus(`Lendo ${files.length} arquivo(s)...`);
    for (const f of files) {
        const text = await readAsText(f, "ISO-8859-1"); // latin1
        loadedFiles.push({ name: f.name, text });
    }
    setStatus(`Arquivos prontos: ${files.map(f => f.name).join(", ")}`);
    btnProcess.disabled = false;
    btnCsv.disabled = true;
    clearTable();
});

btnProcess.addEventListener("click", async () => {
    if (!loadedFiles.length) return;
    setStatus("Processando… (replicando a lógica do Python)");
    try {
        resultRows = await processarArquivos(loadedFiles);
        renderTable(resultRows);
        setStatus(`Processado! Registros: ${resultRows.length}`);
        btnCsv.disabled = resultRows.length === 0;
    } catch (err) {
        console.error(err);
        setStatus("Erro ao processar: " + (err?.message || err));
    }
});

btnCsv.addEventListener("click", () => {
    if (!resultRows.length) return;
    const csv = toCSV(resultRows, COL_ORDER);
    downloadBlob(csv, "resultado_bsp.csv", "text/csv;charset=utf-8;");
});

/* ----------------------- Utils UI ----------------------- */
function setStatus(msg) { statusBox.textContent = msg; }
function clearTable() { thead.innerHTML = ""; tbody.innerHTML = ""; }

function renderTable(rows) {
    clearTable();
    if (!rows.length) return;
    // Cabeçalho fixo com a mesma ordem de colunas do Python
    const header = COL_ORDER;
    thead.innerHTML = `<tr>${header.map(h => `<th>${h}</th>`).join("")}</tr>`;
    const frag = document.createDocumentFragment();
    for (const r of rows) {
        const tr = document.createElement("tr");
        for (const k of header) {
            const v = r[k] ?? "";
            tr.innerHTML += `<td>${formatCell(v)}</td>`;
        }
        frag.appendChild(tr);
    }
    tbody.appendChild(frag);
}

function formatCell(v) {
    if (v instanceof Date) return toISODate(v);
    if (typeof v === "number") {
        // exibe com 2 casas (pt-BR)
        return v.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return (v ?? "").toString();
}

function toCSV(rows, columns) {
    const esc = (s) => {
        const str = s == null ? "" : s instanceof Date ? toISODate(s) : s.toString();
        // aspas e separador
        if (/[",\n;]/.test(str)) return `"${str.replace(/"/g, '""')}"`;
        return str;
    };
    const header = columns.join(",");
    const body = rows.map(r => columns.map(c => esc(r[c])).join(",")).join("\n");
    // BOM para Excel abrir em UTF-8
    return "\uFEFF" + header + "\n" + body;
}

function downloadBlob(content, filename, mime) {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 0);
}

function readAsText(file, encoding = "UTF-8") {
    return new Promise((resolve, reject) => {
        const fr = new FileReader();
        fr.onload = () => resolve(fr.result);
        fr.onerror = reject;
        fr.readAsText(file, encoding);
    });
}

/* --------------------- Conversões datas --------------------- */
const MONTHS = { JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5, JUL: 6, AUG: 7, SEP: 8, OCT: 9, NOV: 10, DEC: 11 };

function parseDDMMMYY(s) {
    // "10JAN25" -> Date(2025,0,10)
    if (!/^\d{2}[A-Z]{3}\d{2}$/.test(s)) return null;
    const dd = Number(s.slice(0, 2));
    const m3 = s.slice(2, 5).toUpperCase();
    const yy = Number(s.slice(5, 7));
    const mm = MONTHS[m3]; if (mm == null) return null;
    const yyyy = yy >= 70 ? 1900 + yy : 2000 + yy;
    const d = new Date(yyyy, mm, dd);
    return Number.isNaN(d.getTime()) ? null : d;
}

function toISODate(d) {
    const pad = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/* --------------------- Núcleo (porta do Python) --------------------- */

// Ordem de colunas (igual ao DataFrame do Python)
const COL_ORDER = [
    "Prefixo", "Doc", "Bilhete", "LOC", "Iata", "Class_reserva", "Base_tarif", "Sistema",
    "data_emi", "data_emb", "CIA", "Status", "Tipo_Emd", "Origem", "Destino",
    "Tarifa_BRL", "Taxa_YQ", "Tarifa_RFN", "liquido_venda", "Tipo_Pax", "Reemissao", "nome_arquivo"
];

/** Lê os arquivos, agrupa por blocos e extrai os campos (espelhando seu Python). */
async function processarArquivos(files) {
    const registros = [];

    for (const f of files) {
        for (const f of files) {
            const linhas = f.text.split(/\r?\n/);
            let bloco = novoBloco();
            for (const linhaRaw of linhas) {
                const linha = (linhaRaw ?? "").toString();
                const t = linha.trimStart();                  // 👈 normaliza início
                const prefixo = t.slice(0, 3);                // 👈 lê o prefixo após trimStart

                if (prefixo === "BKT") {
                    if (bloco.BKT.length) {
                        registros.push({ bloco: cloneBloco(bloco), nome_arquivo: f.name });
                        bloco = novoBloco();
                    }
                    bloco.BKT.push(t);                          // 👈 guarda versão “t”
                } else if (prefixo in bloco) {
                    bloco[prefixo].push(t);                     // 👈 guarda “t”
                }
            }
            if (bloco.BKT.length) registros.push({ bloco: cloneBloco(bloco), nome_arquivo: f.name });
        }

    }

    const dados = [];
    for (const { bloco: r, nome_arquivo } of registros) {
        const bilhete = extrair_bilhete(r.BKS);
        const doc = extrai_doc(r.BKS);
        const loc = extrair_loc(r.BKS);
        const iata = extrai_iata(r.BKS);
        const classe = extrai_classe(r.BKI);
        const baseT = extrair_base_tarifaria(r.BKI);
        const data_emb = extrair_data_emb(r.BKI);
        const data_emi = extrair_data_emi(r.BKS);
        const sistema = extrair_sistema(r.BKT);
        const status = extrair_status(r.BKS);
        const [origem, destino] = extrair_origem_destino(r.BKF, [...r.BKI, ...r.BKS, ...r.BAR, ...r.BKT]);
        const cia = extrair_cia(r.BKS, r.BKI);
        const prefixo = bilhete ? bilhete.slice(0, 3) : null;
        const tipo_pax = extrair_tipo_pax(r);
        const tarifa_rfn = extrair_tarifa_rfn(r.BKS);
        const reemissao = extrair_reemissao(r.BKP);
        const tipoEmd = extrair_tipo_emd(r.BMD);
        const liquido = extrair_liquido_venda(r.BKP);

        let tarifa_brl = 0.0;
        if (status === "RFN") {
            tarifa_brl = tarifa_rfn || 0.0;
        } else {
            const res = extrair_tarifas(r.BAR, [...r.BKS, ...r.BKP]);
            tarifa_brl = (res && res[1] != null) ? res[1] : 0.0;
        }

        const taxa_yq = extrair_valor_yq_bk(r.BKS.join(""));

        dados.push({
            Prefixo: prefixo ?? "",
            Doc: doc ?? "",
            Bilhete: bilhete ?? "",
            LOC: loc ?? "",
            Iata: iata ?? "",
            Class_reserva: classe ?? "",
            Base_tarif: baseT ?? "",
            Sistema: sistema ?? "",
            data_emi: data_emi || "",
            data_emb: data_emb || "",
            CIA: cia ?? "",
            Status: status ?? "",
            Tipo_Emd: tipoEmd ?? "",
            Origem: origem ?? "",
            Destino: destino ?? "",
            Tarifa_BRL: fix2(tarifa_brl),
            Taxa_YQ: fix2(taxa_yq || 0),
            Tarifa_RFN: fix2(tarifa_rfn || 0),
            liquido_venda: fix2(liquido || 0),
            Tipo_Pax: tipo_pax ?? "",
            Reemissao: !!reemissao,
            nome_arquivo
        });
    }

    return dados;
}

function novoBloco() {
    return { BKT: [], BKS: [], BKI: [], BAR: [], BKF: [], BKP: [], BMD: [] };
}
function cloneBloco(b) { return { BKT: [...b.BKT], BKS: [...b.BKS], BKI: [...b.BKI], BAR: [...b.BAR], BKF: [...b.BKF], BKP: [...b.BKP], BMD: [...b.BMD] }; }
function fix2(n) { return Number.isFinite(n) ? Number(n.toFixed(2)) : 0; }

/* ---------------------- Portas das funções Python ---------------------- */

// extrair_bilhete
function extrair_bilhete(linhas_bks) {
    for (const linha of linhas_bks) {
        const m = linha.trim().match(/(\d{13})(\s|$)/);
        if (m) return m[1];
    }
    return null;
}

// extrair_loc
function extrair_loc(linhas_bks) {
    for (let linha of linhas_bks) {
        const up = linha.toUpperCase();
        let marcador = null;
        if (up.includes("TKTT")) marcador = "TKTT";
        else if (up.includes("EMDS")) marcador = "EMDS";
        else if (up.includes("EMDA")) marcador = "EMDA";
        if (!marcador) continue;

        // ✅ em JS, split(x, 1) NÃO devolve o lado direito; pegue o que vem depois do marcador:
        const partes = linha.split(marcador);
        if (partes.length < 2) continue;
        let parte = partes[1].trim();     // tudo após TKTT/EMDS/EMDA
        const original = parte;

        // dupla IATA colada (ex: MIAGRU ABC123) → remove 6 letras no início
        const mDuplo = parte.match(/^([A-Z]{6})\s+(.*)$/);
        if (mDuplo) parte = mDuplo[2].trim();

        // se vier LOC com /XYZ → pega o bloco antes da barra
        const mBarra = parte.match(/([A-Z0-9]{4,8})\/[A-Z0-9]+/);
        if (mBarra) return mBarra[1];

        // 5–8 letras seguidas logo após (LOC puro)
        const m5 = parte.match(/\b([A-Z0-9]{5,8})\b/);
        if (m5 && !/^\d+$/.test(m5[1])) return m5[1];

        // fallback: primeira palavra alfanumérica 4–8
        const mGen = original.match(/\b([A-Z0-9]{4,8})\b/);
        if (mGen && !/^\d+$/.test(mGen[1])) return mGen[1];
    }
    return null;
}


// extrair_status
function extrair_status(linhas_bks) {
    for (const l of linhas_bks) {
        if (l.includes("TKTT")) return "TKT";
        else if (l.includes("EMDS")) return "EMD";
        else if (l.includes("EMDA")) return "EMD";
        else if (l.includes("RFND")) return "RFN";
        else if (l.includes("SPDR")) return "SPDR";
        else if (l.includes("ADNT")) return "ADNT";
        else if (l.includes("SPCR")) return "SPCR";
        else if (l.includes("SPCR")) return "SPCR";
        else if (l.includes("CANX")) return "CAN";
        else if (l.includes("ADMA")) return "ADM";
        else if (l.includes("ACMA")) return "ACM";
    }
    return null;
}

// extrair_valor_yq_bk
function extrair_valor_yq_bk(linha_bk) {
    const linha = (linha_bk || "").replace(/ /g, "");
    if (!linha.includes("{YQ")) return 0;
    const idx = linha.indexOf("{YQ");
    const substr = linha.slice(idx + 3);

    let total = 0;
    const blocos = [...substr.matchAll(/(\d+)([A-Za-z{]+)/g)]; // (digits)(letters or {)
    let somando = false;

    for (const m of blocos) {
        const digits = m[1];
        const letters = m[2];

        if (!somando) {
            if (letters.endsWith("YQ") || letters.endsWith("YR")) {
                somando = true;
                const val = intDiv10(digits) + letraPos(letters[0]) / 100;
                total += val;
                if (letters.endsWith("YR")) break;
            } else {
                continue;
            }
        } else {
            if (letters.endsWith("YQ")) {
                const val = intDiv10(digits) + letraPos(letters[0]) / 100;
                total += val;
            } else if (letters.endsWith("YR")) {
                const val = intDiv10(digits) + letraPos(letters[0]) / 100;
                total += val;
                break;
            } else if (letters.includes("{")) {
                // '0000000685{BR' → 68.5
                total += intDiv10(digits);
                break;
            } else {
                break;
            }
        }
    }
    return fix2(total);
}
function intDiv10(s) { return Number(parseInt(s, 10)) / 10; }
function letraPos(ch) {
    const c = (ch || "").toUpperCase().charCodeAt(0);
    if (c >= 65 && c <= 90) return c - 65 + 1;
    return 0;
}

// extrair_tarifas  -> retorna [tarifa_usd, tarifa_brl]
function extrair_tarifas(linhas_bar, linhas_bks) {
    let tarifa_usd = null, tarifa_brl = null;

    // 1) Regra padrão de BAR
    for (const lb of linhas_bar) {
        const bar_clean = lb.replace(/ /g, "").toUpperCase();
        if (bar_clean.includes("ADC")) return [0.0, 0.0];

        const mUSD = lb.toUpperCase().match(/USD\s*([0-9]{1,6}\.\d{2})|USD([0-9]{1,6}\.\d{2})/);
        const mBRL = lb.toUpperCase().match(/BRL\s*([0-9]{1,6}\.\d{2})|BRL([0-9]{1,6}\.\d{2})/);
        if (mUSD) tarifa_usd = parseFloat(mUSD[1] || mUSD[2]);
        if (mBRL && tarifa_brl == null) tarifa_brl = parseFloat(mBRL[1] || mBRL[2]);
    }

    function parse_bks_para_tarifa(linhas) {
        for (const linha of linhas) {
            const ln = (linha || "").trim();
            if (ln.startsWith("BKS") && ln.length >= 38) {
                const restante = ln.slice(38).replace(/ /g, "");
                if (!restante) continue;
                const restante_valores = restante.slice(1); // ignora primeiro dígito
                const corte = restante_valores.indexOf("{");
                if (corte !== -1) {
                    const trecho = restante_valores.slice(0, corte);
                    if (/^0+$/.test(trecho)) return 0.0;

                    const m = trecho.match(/(\d{6})([A-Z])/);
                    if (m) {
                        const dig = parseInt(m[1], 10);
                        const letra = m[2];
                        const inteiro = dig / 10;
                        const decimal = letraPos(letra) / 100;
                        return fix2(inteiro + decimal);
                    } else {
                        const n = parseInt(trecho, 10);
                        if (Number.isFinite(n)) return fix2(n / 10);
                        return 0.0;
                    }
                }
            }
        }
        return null;
    }

    // 2) BAR contém "INV" → usa parse de BKS
    if (linhas_bar.some(l => l.toUpperCase().includes("INV"))) {
        const v = parse_bks_para_tarifa(linhas_bks);
        if (v != null) tarifa_brl = v;
    }

    // 3) BAR contém "IT"
    if (linhas_bar.some(l => l.toUpperCase().includes("IT"))) {
        const v = parse_bks_para_tarifa(linhas_bks);
        if (v != null) tarifa_brl = v;
    }

    // 4) BAR contém "BT"
    if (linhas_bar.some(l => l.toUpperCase().includes("BT"))) {
        const v = parse_bks_para_tarifa(linhas_bks);
        if (v != null) tarifa_brl = v;
    }

    // 3️⃣ BAR numérico puro sem moeda
    for (const lb of linhas_bar) {
        const bar_clean = lb.toUpperCase().replace(/ /g, "");
        if (!["INV", "IT", "BT", "ADC", "USD", "BRL"].some(t => bar_clean.includes(t))) {
            if (/^BAR\d+\d*[0-9\/ ]+$/.test(lb)) {
                const v = parse_bks_para_tarifa(linhas_bks);
                if (v != null) { tarifa_brl = v; break; }
            }
        }
    }

    // 4) "EX" no BKP
    if (linhas_bks.some(l => l.startsWith("BKP") && l.toUpperCase().includes("EX"))) {
        const v = parse_bks_para_tarifa(linhas_bks);
        if (v != null) tarifa_brl = v;
    }

    return [tarifa_usd, tarifa_brl];
}

// extrair_tarifa_rfn
function extrair_tarifa_rfn(linhas_bks) {
    for (let i = 0; i < linhas_bks.length; i++) {
        if (linhas_bks[i].includes("RFND")) {
            for (let j = i + 1; j < linhas_bks.length; j++) {
                const l = linhas_bks[j].replace(/ /g, "");
                if (/0{11}\{/.test(l)) return 0.0;
                const m = l.match(/(\d+)([A-Z])0{5,}\{/);
                if (m) {
                    const raw = m[1];
                    const letra = m[2];
                    const keep = raw.length > 6 ? raw.slice(-6) : raw;
                    const inteiro = parseInt(keep, 10) / 10;
                    const decimal = (letra.charCodeAt(0) - 65) / 100;
                    return fix2(inteiro + decimal);
                }
            }
            break;
        }
    }
    return 0.0;
}

// extrair_liquido_venda (usa última BKP)
function extrair_liquido_venda(linhas_bkp) {
    const bkp = (linhas_bkp || []).filter(s => typeof s === "string" && s.startsWith("BKP"));
    if (!bkp.length) return 0.0;
    const last = bkp[bkp.length - 1];

    const matches = [...last.matchAll(/(\d{7,})([A-Z]|\{|\})(?![A-Z])/g)];
    if (!matches.length) return 0.0;
    const m = matches[matches.length - 1];
    const digitos = m[1], term = m[2];

    const ult7 = parseInt(digitos.slice(-7), 10);
    if (!ult7) return 0.0;

    const inteiro = ult7 / 10;
    let decimal = 0;
    if (term !== "{" && term !== "}") {
        // exceção P=7 (=> 0,07)
        const mapaEx = { "P": 7 };
        const pos = mapaEx[term.toUpperCase()] ?? (term.toUpperCase().charCodeAt(0) - 65 + 1);
        decimal = pos / 100;
    }
    return fix2(inteiro + decimal);
}

// extrair_origem_destino
function extrair_origem_destino(linhas_bkf, linhas_bks_e_bar) {
    // --- util para origem/destino ---
    const INVALIDAS = new Set(["BRL", "USD", "NUC", "END", "ROE", "EUR", "XT", "IT", "LA", "QR"]);
    function descartar_invalidos(sigla) {
        if (!sigla) return null;
        return INVALIDAS.has(sigla) ? null : sigla;
    }

    const descartar = (s) => (s && !INVALIDAS.has(s)) ? s : null;

    // 1) Tratamento ADC prioritário
    for (const l of linhas_bks_e_bar) {
        if (l.startsWith("BKS") && l.includes("ADC") && l.includes("TKTT")) {
            const m = l.match(/TKTT\s+([A-Z]{3})([A-Z]{3})/);
            if (m) return [descartar(m[1]), descartar(m[2])];
        }
    }

    // 2) Fallback sem BKF
    // --- 2) Fallback sem BKF ---
    if (!linhas_bkf || !linhas_bkf.length) {
        // 2.1) padrão tradicional BKI (com índice):
        for (const linha of linhas_bks_e_bar) {
            if (!linha.startsWith("BKI")) continue;
            let m = linha.match(/BKI\d+\s+\d{2}([A-Z]{3})\s+([A-Z]{3})/);
            if (m) {
                const o = descartar_invalidos(m[1]);
                const d = descartar_invalidos(m[2]);
                if (o && d) return [o, d];
            }
        }
        // 2.2) par de IATA seguidos em qualquer lugar da BKI (ex.: " CWB  AEP ")
        for (const linha of linhas_bks_e_bar) {
            if (!linha.startsWith("BKI")) continue;
            const pares = [...linha.matchAll(/\b([A-Z]{3})\s+([A-Z]{3})\b/g)];
            for (const p of pares) {
                const o = descartar_invalidos(p[1]);
                const d = descartar_invalidos(p[2]);
                if (o && d && o !== d) return [o, d];
            }
        }
        return [null, null];
    }


    const texto_bkf = (linhas_bkf || []).join(" ");
    let origem = null, destino = null;

    // Origem (BKF inicial)
    {
        const m = (linhas_bkf[0] || "").match(/BKF\d+\s+\d{2}([A-Z]{3})/);
        if (m) origem = descartar(m[1]);
    }

    // I-XXX
    {
        const m = texto_bkf.match(/I-([A-Z]{3})/);
        if (m) origem = descartar(m[1]);
    }

    // S-XXX simples
    if (texto_bkf.includes("S-")) {
        const m = texto_bkf.match(/S-([A-Z]{3})/);
        if (m) {
            origem = descartar(m[1]);
            const m2 = texto_bkf.match(/([A-Z]{3})\s+Q\d+\.\d{2}M\d+\.\d{2}/);
            if (m2) {
                const cand = descartar(m2[1]);
                if (cand) return [origem, cand];
            }
        }
    }

    // Prefixos com data
    for (const rx of [/S-[0-9]{2}[A-Z]{3}[0-9]{2}([A-Z]{3})/,
        /I-[0-9]{2}[A-Z]{3}[0-9]{2}([A-Z]{3})/]) {
        const m = texto_bkf.match(rx);
        if (m) origem = descartar(m[1]);
    }

    // Qualquer DDMMMYYXXX
    {
        const m = texto_bkf.match(/[0-9]{2}[A-Z]{3}[0-9]{2}([A-Z]{3})/);
        if (m) origem = descartar(m[1]);
    }

    // M/IT lookback em S-charges
    {
        const m_it = texto_bkf.match(/([A-Z]{3})\s+M\/IT/);
        if (m_it) {
            const charges = [...texto_bkf.matchAll(/([A-Z]{3})(?=\s+S\d+\.\d{2})/g)];
            for (let i = charges.length - 1; i >= 0; i--) {
                if (charges[i].index < m_it.index) {
                    const cand = descartar(charges[i][1]);
                    if (cand && cand !== origem) return [origem, cand];
                }
            }
            const cand = descartar(m_it[1]);
            if (cand && cand !== origem) return [origem, cand];
        }
    }

    // padrões de destino
    const patterns = [
        [/([A-Z]{3})([A-Z]{3})(?=\d+\.\d{2})/, 2],
        [/\b([A-Z]{3})\s+\d+\.\d{2}/, 1],
        [/(?<=\s)([A-Z]{3})(?=\d+)/, 1],
        [/([A-Z]{3})\s+\d+\.\d{2}\s*\/-[A-Z0-9]+/, 1],
        [/X\/[A-Z]{3}.*?([A-Z]{3})\s+Q\d+\.\d{2}/, 1],
        [/([A-Z]{3})\s+M\/IT/, 1],
        [/S\d+\.\d{2}\s+[A-Z]{2}\s+([A-Z]{3})/, 1],
        [/([A-Z]{3})\s?[A-Z]?\d{2,4}\.\d{2}/, 1],
    ];
    for (const [rx, grp] of patterns) {
        const m = texto_bkf.match(rx);
        if (m) {
            const cand = descartar(m[grp]);
            if (cand && cand !== origem) return [origem, cand];
        }
    }

    // Sem "X/" → último aeroporto válido (≠ origem), priorizando grudado à tarifa
    if (origem && !texto_bkf.includes("X/")) {
        const cods = [...texto_bkf.matchAll(/\b([A-Z]{3})\b/g)].map(x => x[1]);
        const validos = cods.filter(c => !INVALIDAS.has(c) && c !== origem);
        if (validos.length) {
            const grudados = [...texto_bkf.matchAll(/([A-Z]{3})(?=\d+\.\d{2})/g)].map(x => x[1]);
            for (const g of grudados) if (validos.includes(g)) return [origem, g];
            let last = validos[validos.length - 1];
            if (last === "NUC" && validos.length > 1) last = validos[validos.length - 2];
            return [origem, last];
        }
    }

    // Fallback único após a origem
    if (origem) {
        const pos = texto_bkf.split(origem, 1)[1] || "";
        const poss = [...pos.matchAll(/([A-Z]{3})(?=\d+\.\d{2})/g)].map(x => x[1]).filter(c => !INVALIDAS.has(c));
        if (poss.length === 1) return [origem, poss[0]];
    }

    return [origem, null];
}

// extrair_cia
function extrair_cia(linhas_bks, linhas_bki) {
    const prefixo = extrai_pref(linhas_bks);
    const mapa = {
        '001': 'AA', '005': 'CO', '006': 'DL', '014': 'AC', '016': 'UA', '037': 'US', '045': 'LA', '047': 'TP', '055': 'AZ',
        '057': 'AF', '064': 'OK', '074': 'KL', '075': 'IB', '077': 'MS', '080': 'LO', '081': 'QF', '083': 'SA', '086': 'NZ',
        '105': 'AY', '117': 'SK', '125': 'BA', '127': 'G3', '131': 'JL', '139': 'AM', '142': 'KF', '160': 'CX', '165': 'JP',
        '180': 'KE', '182': 'MA', '205': 'NH', '217': 'TG', '220': 'LH', '230': 'CM', '235': 'TK', '257': 'OS', '353': 'NU',
        '462': 'XL', '469': '4M', '512': 'RJ', '544': 'LP', '555': 'SU', '577': 'AD', '618': 'SQ', '680': 'JK', '688': 'EG',
        '706': 'KQ', '708': 'JO', '724': 'LX', '774': 'FM', '831': 'OU', '957': 'JJ', '988': 'OZ', '996': 'UX', '999': 'CA',
        '428': 'JC', '134': 'AV', '118': 'DT', '044': 'AR', '930': 'OB', '176': 'EK', '169': 'HR', '157': 'QR', '071': 'ET',
        '147': 'AT', '605': 'H2', '973': 'JA', '275': 'GP', '799': 'BNS', 'JC': 'JC', '952': 'SPCR'
    };
    if (prefixo && mapa[prefixo]) return mapa[prefixo];

    // fallback pela cia mais frequente nas BKI
    const cias = [];
    for (const l of linhas_bki) {
        const m = l.match(/\s([A-Z0-9]{2})\s+\d{1,4}\s/);
        if (m) cias.push(m[1]);
    }
    if (cias.length) {
        const freq = {};
        for (const c of cias) freq[c] = (freq[c] || 0) + 1;
        return Object.entries(freq).sort((a, b) => b[1] - a[1])[0][0];
    }
    return null;
}

// extrair_tipo_pax
function extrair_tipo_pax(bloco) {
    const tipos = new Set(['ADT', 'CHD', 'INF', 'CNN']);
    const todas = [...bloco.BKT, ...bloco.BKS, ...bloco.BKI, ...bloco.BAR, ...bloco.BKF, ...bloco.BKP];
    for (const l of todas) {
        for (const t of tipos) {
            if (l.includes(t)) return t;
        }
    }
    return '';
}

// extrai_pref (3 primeiros dígitos do bilhete)
function extrai_pref(linhas_bks) {
    for (const linha of linhas_bks) {
        const m = linha.trim().match(/(\d{13})(?:\s|$)/);
        if (m) return m[1].slice(0, 3);
    }
    return null;
}

// extrai_doc (10+ dígitos após prefixo do bilhete)
function extrai_doc(linhas_bks) {
    for (const linha of linhas_bks) {
        const m = linha.trim().match(/(\d{13})(?:\s|$)/);
        if (m) return m[1].slice(3);
    }
    return null;
}

// extrai_classe
function extrai_classe(linhas_bki) {
    // 1) padrão clássico: "KL 1756 J"
    for (const linha of linhas_bki) {
        const m = linha.match(/\b([A-Z0-9]{2})\s+(\d{1,4})\b(?:.{0,40}?)\b([A-Z])\b/);
        if (m) return m[3];
    }
    // 2) classe colada ao voo: "KL1756 J"
    for (const linha of linhas_bki) {
        const m = linha.match(/\b([A-Z0-9]{2})(\d{1,4})\b(?:.{0,40}?)\b([A-Z])\b/);
        if (m) return m[3];
    }
    // 3) proximidade de "OK" (muitos layouts colocam a classe logo antes do OK)
    for (const linha of linhas_bki) {
        const m = linha.match(/\b([A-Z])\b(?=.*\bOK\b)/);
        if (m) return m[1];
    }
    return null;
}


// extrair_base_tarifaria (versão expandida)
function extrair_base_tarifaria(linhas_bki) {
    for (const linha of linhas_bki) {
        const clean = linha.replace(/\s+/g, " ");
        const mVoo = clean.match(/\b([A-Z0-9]{2})\s+(\d{1,4})\s/);
        if (!mVoo) continue;

        const pos = (mVoo.index ?? clean.indexOf(mVoo[0])) + mVoo[0].length;
        const resto = clean.slice(pos);

        let m;
        if ((m = resto.match(/OK\d*PC([A-Z0-9]+)/))) return m[1];
        if ((m = resto.match(/PC\s*([A-Z0-9]+)/))) return m[1];
        if ((m = resto.match(/OKXX\s+([A-Z0-9/]+)/))) return m[1].split("/")[0];

        if ((m = resto.match(/OK([A-Z0-9]+)/))) {
            const trecho = m[1];
            const m2 = trecho.match(/(\d+)K([A-Z0-9]+)/);
            return m2 ? m2[2] : trecho;
        }
        if ((m = resto.match(/OK\s+([A-Z0-9]+)/))) return m[1];

        if ((m = resto.match(/NSNI([A-Z0-9/]+)/))) {
            let t = m[1];
            if (t.startsWith("NSNIL")) return t.slice(5);
            if (t.startsWith("L")) return t.slice(1);
            return t;
        }
        if ((m = resto.match(/NS(\d+)K([A-Z0-9]+)/))) return m[2];
        if ((m = resto.match(/NSXX\s+([A-Z0-9/]+)/))) return m[1].split("/")[0];

        if ((m = resto.match(/PCC\s+([A-Z0-9]+)/))) return m[1];
        if ((m = resto.match(/OKP\s+([A-Z0-9]+)/))) return m[1];
        if ((m = resto.match(/OBF([A-Z0-9/]+)/))) return m[1].split("/")[0];
        if ((m = resto.match(/TAR([A-Z0-9]+)/))) return m[1];
        if ((m = resto.match(/FBC:\s*([A-Z0-9]+)/))) return m[1];
        if ((m = resto.match(/B\/F:\s*([A-Z0-9]+)/))) return m[1];
        if ((m = resto.match(/BT\s+([A-Z0-9]+)/))) return m[1];
        if ((m = resto.match(/BS\/([A-Z0-9]+)/))) return m[1];
        if ((m = resto.match(/TST:\s*\d+\s+([A-Z0-9]+)/))) return m[1];

        // 🔚 Fallback: pega um “bloco-base” comum (ex.: MLEKM40E, YBR123, etc)
        if ((m = resto.match(/\b([A-Z][A-Z0-9]{2,12})\b/))) return m[1];
    }
    return null;
}




// extrair_data_emb (DDMMMYY -> Date)
function extrair_data_emb(linhas_bki) {
    for (const l of linhas_bki) {
        const m = l.match(/(\d{2}[A-Z]{3}\d{2})/);
        if (m) {
            const d = parseDDMMMYY(m[1]);
            if (d) return d;
        }
    }
    return null;
}

// extrair_data_emi (regras com zeros e fallback)
function extrair_data_emi(linhas_bks) {
    function tentarExtrair(linha_limpa, zeros_consecutivos) {
        const rx = new RegExp(`(\\d+?)(0{${zeros_consecutivos},})`);
        const m = linha_limpa.match(rx);
        if (!m) return null;
        let seq = m[1];
        let ult6 = seq.slice(-6);
        if (ult6.length === 5) ult6 += "0";
        if (ult6.length < 6) return null;
        const [ano, mes, dia] = [ult6.slice(0, 2), ult6.slice(2, 4), ult6.slice(4, 6)];
        const ddmmyy = `${dia}${mes}${ano}`;
        return parseDDMMYY(ddmmyy);
    }
    function tentarExtrairBlocoIntermediario(linha_limpa) {
        const m = linha_limpa.match(/(0{2,})(\d+?)(0{2,})/);
        if (m && m[1].length === m[3].length) {
            let seq = m[2]; let ult6 = seq.slice(-6);
            if (ult6.length === 5) ult6 += "0";
            if (ult6.length < 6) return null;
            const [ano, mes, dia] = [ult6.slice(0, 2), ult6.slice(2, 4), ult6.slice(4, 6)];
            return parseDDMMYY(`${dia}${mes}${ano}`);
        }
        return null;
    }
    function tentarExtrairParada00(linha_limpa) {
        for (let i = 1; i < linha_limpa.length; i++) {
            if (linha_limpa[i - 1] === '0' && linha_limpa[i] === '0') {
                if (i >= 2 && linha_limpa[i - 2] === '0') continue;
                const seq = linha_limpa.slice(0, i - 1);
                let ult6 = seq.slice(-6);
                if (ult6.length === 5) ult6 += "0";
                if (ult6.length < 6) return null;
                const [ano, mes, dia] = [ult6.slice(0, 2), ult6.slice(2, 4), ult6.slice(4, 6)];
                const d = parseDDMMYY(`${dia}${mes}${ano}`);
                if (d) return d;
            }
        }
        return null;
    }
    function tentarExtrairFallback8Digitos(linha_limpa) {
        for (let i = 0; i <= linha_limpa.length - 8; i++) {
            const bloco = linha_limpa.slice(i, i + 8);
            if (/^\d{8}$/.test(bloco)) {
                let ult6 = bloco.slice(-6);
                if (ult6.length === 5) ult6 += "0";
                if (ult6.length < 6) continue;
                const [ano, mes, dia] = [ult6.slice(0, 2), ult6.slice(2, 4), ult6.slice(4, 6)];
                const d = parseDDMMYY(`${dia}${mes}${ano}`);
                if (d) return d;
            }
        }
        return null;
    }
    function parseDDMMYY(s) {
        if (!/^\d{6}$/.test(s)) return null;
        const dd = Number(s.slice(0, 2)), mm = Number(s.slice(2, 4)) - 1, yy = Number(s.slice(4, 6));
        const yyyy = yy >= 70 ? 1900 + yy : 2000 + yy;
        const d = new Date(yyyy, mm, dd);
        return Number.isNaN(d.getTime()) ? null : d;
    }

    for (const linha of linhas_bks) {
        const l = linha.trim().replace(/^BKS0+/, "");
        let d = tentarExtrair(l, 2) || tentarExtrairBlocoIntermediario(l) || tentarExtrairParada00(l) ||
            tentarExtrair(l, 3) || tentarExtrairFallback8Digitos(l);
        if (d) return d;
    }
    return null;
}

// extrair_sistema
function extrair_sistema(linhas_bkt) {
    const mapa = {
        'EDIS': 'NDC',
        'WEBL': 'ELATAM',
        'AGTD': 'AMADEUS',
        'FLGX': 'NDC (FLGX)',
        'SABR': 'SABER',
        'GDSL': 'GALILEO',
        'EARS': 'BSP LINK',
        'SYST': 'sistema BspLink'
    };
    for (const l of linhas_bkt) {
        for (const k of Object.keys(mapa)) if (l.includes(k)) return mapa[k];
    }
    return null;
}

// extrair_reemissao
function extrair_reemissao(linhas_bkp) {
    for (const l of linhas_bkp) if (l.startsWith("BKP") && l.toUpperCase().includes("EX")) return true;
    return false;
}

// extrair_tipo_emd
function extrair_tipo_emd(linhas_bmd) {
    for (const l of linhas_bmd) {
        if (l.includes("SA") || l.includes("SEAT") || l.includes("ASIENTOS")) return "ASSENTO";
        if (l.includes("BG") || l.includes("BAG") || l.includes("23K")) return "BAGAGEM";
        if (l.includes("CHANGE")) return "CHANGE FEE";
        if (l.includes("PENALTY")) return "PENALTY FEE";
        if (l.includes("PT")) return "PETC";
        if (l.includes("UN")) return "MENOR DESACOMPANHADO";
        if (l.includes("99")) return "RESIDUAL VALUE";
    }
    return null;
}

// extrai_iata
function extrai_iata(linhas_bks) {
    for (const linha of linhas_bks) {
        const partes = linha.split(/\s+/);
        for (const p of partes) {
            if (/^\d{8}$/.test(p)) return p;
            let m = p.match(/^(\d{8})([A-Z])$/);
            if (m) return m[1];
            m = p.match(/^(\d{8})[A-Z0-9]{2,}/);
            if (m) return m[1];
        }
    }
    return null;
}

// (opcional, não usado na montagem do DF final aqui, mas fiel ao Python)
function extrair_valor_trecho(linhas_bkf, destino, cia_principal, linhas_bar) {
    for (const l of linhas_bar) if (l.toUpperCase().includes("ADC")) return 0.0;
    const texto = (linhas_bkf || []).join(" ").toUpperCase().replace(/  +/g, " ");
    if (!destino || !texto) return 0.0;

    const tokens = texto.split(/\s+/);
    let total = 0;
    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        if (token === destino) {
            for (let j = i; j >= 0; j--) {
                const tk = tokens[j];
                if (/^[A-Z]{3}$/.test(tk) && tk !== destino) break;
                const ciaToken = tokens[j + 1] || "";
                const isCiaPrincipal = ciaToken === cia_principal;
                const m = tk.match(/(Q|M)?(\d{1,5}\.\d{2})/);
                if (m) {
                    const valor = parseFloat(m[2]);
                    if (j + 1 >= tokens.length || isCiaPrincipal) total += valor;
                }
            }
            break;
        }
    }
    return fix2(total);
}
