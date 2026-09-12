'use strict';

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');
const { pathToFileURL } = require('url');

const ROOT = __dirname;
const config = JSON.parse(fs.readFileSync(path.join(ROOT, 'config.json'), 'utf8'));
const MODE = (process.argv[2] || 'full').toLowerCase();
const LOG_DIR = path.join(ROOT, 'logs');
const AUTOMATION_LOCAL_DIR = path.join(process.env.LOCALAPPDATA || os.tmpdir(), 'Portal504_Automacao');
const AUTH_STATE = path.join(AUTOMATION_LOCAL_DIR, 'athena_storage_tpo.json');
fs.mkdirSync(LOG_DIR, { recursive: true });
fs.mkdirSync(AUTOMATION_LOCAL_DIR, { recursive: true });

const TIMEOUT = Number(config.tempoMaximoPassoMs || 30000);
const LOGIN_TIMEOUT = Number(config.tempoMaximoLoginMs || 180000);
const downloadsDir = path.join(os.homedir(), 'Downloads');

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function nowStamp() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}
function log(msg) { console.log(`[TPO] ${msg}`); }
function step(n, total, msg) { console.log(`\n[${n}/${total}] ${msg}`); }
function currentMonthLabel() {
  return ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'][new Date().getMonth()];
}
function ymd(d) {
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}`;
}
function dmy(d) {
  const p = n => String(n).padStart(2, '0');
  return `${p(d.getDate())}/${p(d.getMonth()+1)}/${d.getFullYear()}`;
}
function newestExcelInDownloads(afterMs, onlyTpoReport=false) {
  const exts = new Set(['.xlsx','.xls']);
  const rows = fs.readdirSync(downloadsDir)
    .map(name => ({ name, full: path.join(downloadsDir,name) }))
    .filter(x => exts.has(path.extname(x.name).toLowerCase()))
    .filter(x => !onlyTpoReport || /^relatorio-generico_/i.test(x.name))
    .map(x => ({ ...x, st: fs.statSync(x.full) }))
    .filter(x => x.st.mtimeMs >= afterMs - 5000)
    .sort((a,b) => b.st.mtimeMs - a.st.mtimeMs);
  return rows[0]?.full || null;
}
function expectedCurrentRange() {
  const today = new Date();
  const firstOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  // REGRA TPO: sempre trazer também o último dia do mês anterior.
  // Isso captura lotes que abriram no fim do mês anterior e fecharam/coletaram no dia 1º
  // do mês atual (ex.: lote aberto em 31/08 e amostras coletadas em 01/09).
  const overlapStart = new Date(firstOfMonth);
  overlapStart.setDate(overlapStart.getDate() - 1);
  return { first: overlapStart, firstOfMonth, today };
}
function validateAthenaFilenameRange(file, range) {
  const name = path.basename(file);
  const m = name.match(/^relatorio-generico_(\d{2})_(\d{2})_(\d{4})_(\d{2})_(\d{2})_(\d{4})(?: \(\d+\))?\.xlsx?$/i);
  if (!m) {
    log(`AVISO: não consegui validar o período pelo nome do arquivo: ${name}`);
    return;
  }
  const from = `${m[1]}/${m[2]}/${m[3]}`;
  const to = `${m[4]}/${m[5]}/${m[6]}`;
  const expFrom = dmy(range.first);
  const expTo = dmy(range.today);
  if (from !== expFrom || to !== expTo) {
    throw new Error(`O Athena baixou um período diferente do solicitado. Esperado ${expFrom} a ${expTo}, mas o arquivo indica ${from} a ${to}. O Portal TPO NÃO será alterado.`);
  }
  log(`Período confirmado pelo arquivo baixado: ${from} a ${to}.`);
}
async function saveDebug(page, label) {
  if (!page) return;
  const file = path.join(LOG_DIR, `${nowStamp()}_${label.replace(/[^a-z0-9_-]+/gi,'_')}.png`);
  try {
    await page.screenshot({ path: file, fullPage: true });
    log(`Imagem de diagnóstico salva em: ${file}`);
  } catch (_) {}
}
async function waitVisible(locator, timeout=TIMEOUT) {
  await locator.waitFor({ state: 'visible', timeout });
  return locator;
}
async function clickText(page, text, opts={}) {
  const loc = page.getByText(text, { exact: !!opts.exact }).filter({ visible: true }).first();
  await waitVisible(loc, opts.timeout || TIMEOUT);
  await loc.click();
}
async function clickButtonByText(page, text, timeout=TIMEOUT) {
  const re = typeof text === 'string' ? new RegExp(text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i') : text;
  const btn = page.getByRole('button', { name: re }).first();
  await waitVisible(btn, timeout);
  await btn.click();
}



function portal504HashCredential(value) {
  let h1 = 0xdeadbeef ^ value.length;
  let h2 = 0x41c6ce57 ^ value.length;
  for (let i = 0, ch; i < value.length; i++) {
    ch = value.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}

async function autorizarPortal504(page) {
  // Usa o mesmo token da autenticação local do Portal 504, sem gravar a senha em texto puro.
  // window.name sobrevive à navegação Athena -> file:// e é reconhecido pelos guards dos portais.
  const usuario = 'acesso';
  const senhaHash = '1j78k6ardjv';
  const token = portal504HashCredential(`portal504|${usuario}|${senhaHash}|autorizado`);
  const sessao = { token, expiraEm: Date.now() + 12 * 60 * 60 * 1000 };
  await page.evaluate((s) => {
    window.name = `PORTAL504_AUTH:${encodeURIComponent(JSON.stringify(s))}`;
    try { localStorage.setItem('portal504_auth_v1', JSON.stringify(s)); } catch (_) {}
  }, sessao).catch(()=>{});
}

async function tratarConfirmacaoEntradaAthena(page, esperaMs=4500) {
  const fim = Date.now() + esperaMs;
  const modalRe = /^(Entrar|Entrar no mês|Entrar no mes|Confirmar|Continuar|Acessar|Prosseguir|Sim|OK)$/i;
  const globalRe = /^(Entrar|Entrar no mês|Entrar no mes|Confirmar|Continuar|Acessar|Prosseguir)$/i;
  let total = 0;
  let ultimaConfirmacao = 0;

  while (Date.now() < fim) {
    let clicou = false;

    // Procura primeiro dentro de modal/dialog. Aqui aceitamos também Sim/OK,
    // porque são respostas seguras apenas quando o botão está dentro do modal.
    const modais = page.locator('[role="dialog"]:visible, .ant-modal:visible, .modal:visible, mat-dialog-container:visible');
    const n = await modais.count().catch(()=>0);
    for (let i=0;i<n;i++) {
      const botao = modais.nth(i).getByRole('button',{name:modalRe}).first();
      if (await botao.count() && await botao.isVisible().catch(()=>false)) {
        const txt = ((await botao.innerText().catch(()=>'')) || 'Entrar').trim();
        log(`Athena pediu confirmação de entrada. Clicando em "${txt}"...`);
        await botao.click({timeout:8000});
        total++;
        ultimaConfirmacao = Date.now();
        clicou = true;
        await sleep(850);
        break;
      }
    }

    if (clicou) continue;

    // Fallback global fica restrito aos nomes conhecidos para não clicar em OK/Sim fora de um modal.
    const global = page.getByRole('button',{name:globalRe}).filter({visible:true}).first();
    if (await global.count() && await global.isVisible().catch(()=>false)) {
      const txt=(await global.innerText().catch(()=>'' )).trim();
      log(`Confirmação condicional do Athena detectada: "${txt || 'Entrar'}".`);
      await global.click({timeout:8000});
      total++;
      ultimaConfirmacao = Date.now();
      await sleep(850);
      continue;
    }

    // Depois de clicar em uma confirmação, espera uma janela curta sem novos diálogos.
    // Isso cobre o caso real do Athena pedir duas confirmações seguidas.
    if (total > 0 && Date.now() - ultimaConfirmacao >= 1200) break;
    await sleep(250);
  }

  if (total > 1) log(`Athena exigiu ${total} confirmações sequenciais; todas foram tratadas.`);
  return total;
}

function isTargetCrashError(err) {
  const msg = String(err && err.message ? err.message : err || '');
  return /Target crashed|Target page, context or browser has been closed|browser has been closed|page crashed|crash/i.test(msg);
}

async function launchAutomationBrowser() {
  // V3: NÃO usa launchPersistentContext nem o Chrome corporativo.
  // Usa o Chromium isolado do Playwright, portanto o Portal de Indicadores
  // continua aberto no Chrome normal sem conflito de sessão/perfil.
  log('Abrindo navegador isolado do Playwright (sem perfil persistente do Chrome).');
  let browser;
  try {
    browser = await chromium.launch({
      headless: false,
      args: ['--start-maximized', '--disable-gpu', '--no-first-run', '--no-default-browser-check']
    });
  } catch (e) {
    const msg = String(e && e.message ? e.message : e || '');
    if (/Executable doesn't exist|browserType\.launch: Executable/i.test(msg)) {
      throw new Error('O Chromium do Playwright não está instalado neste VD. Rode uma vez: npx playwright install chromium');
    }
    throw e;
  }
  const contextOpts = { acceptDownloads: true, viewport: null, permissions: ['geolocation'] };
  if (fs.existsSync(AUTH_STATE)) {
    contextOpts.storageState = AUTH_STATE;
    log(`Sessão Athena carregada de: ${AUTH_STATE}`);
  } else {
    log('Primeiro uso deste navegador isolado: o Athena poderá pedir login/MFA uma vez.');
  }
  const context = await browser.newContext(contextOpts);
  try {
    await context.grantPermissions(['geolocation'], { origin: new URL(config.athenaUrl).origin });
    log('Permissão de localização do Athena liberada automaticamente.');
  } catch (e) {
    log(`Aviso: não foi possível pré-liberar localização: ${e.message}`);
  }
  const pg = await context.newPage();
  pg.setDefaultTimeout(TIMEOUT);

  // Também cobre confirmações nativas do navegador usadas pelo Athena.
  pg.on('dialog', async dialog => {
    try {
      const msg = dialog.message() || '';
      if (/athena/i.test(pg.url()) && /entrar|confirm|continuar|m[eê]s/i.test(msg)) {
        log(`Confirmação nativa do Athena: ${msg}`);
        await dialog.accept();
      }
    } catch (_) {}
  });

  return { browser, context, page: pg };
}

async function saveAthenaAuthState(page) {
  try {
    await page.context().storageState({ path: AUTH_STATE });
    log('Sessão Athena salva para as próximas execuções.');
  } catch (_) {}
}

async function openAthenaWithOneRecovery(browser, page) {
  try {
    await page.goto(config.athenaUrl, { waitUntil: 'domcontentloaded', timeout: LOGIN_TIMEOUT });
    await ensureAthenaLogin(page);
    await tratarConfirmacaoEntradaAthena(page, 1500).catch(()=>false);
    await saveAthenaAuthState(page);
    return { browser, page };
  } catch (err) {
    if (!isTargetCrashError(err)) throw err;
    log('O Chrome controlado caiu ao abrir o Athena. Reiniciando automaticamente uma vez...');
    try { await browser.close(); } catch (_) {}
    await sleep(3000);
    const relaunched = await launchAutomationBrowser();
    await relaunched.page.goto(config.athenaUrl, { waitUntil: 'domcontentloaded', timeout: LOGIN_TIMEOUT });
    await ensureAthenaLogin(relaunched.page);
    await tratarConfirmacaoEntradaAthena(relaunched.page, 1500).catch(()=>false);
    await saveAthenaAuthState(relaunched.page);
    log('Navegador recuperado; continuando a automação.');
    return relaunched;
  }
}

async function ensureAthenaLogin(page) {
  // O Athena pode abrir diretamente em Relatórios quando a sessão corporativa já existe.
  if (await page.getByText('Relatórios', { exact: true }).count()) return;

  const sso = page.locator('#btn-sso').first();
  if (await sso.count()) {
    log('Botão SSO encontrado. Clicando em "Acesse Aqui"...');
    await sso.click();
  } else {
    const byText = page.getByRole('button', { name: /Acesse Aqui/i }).first();
    if (await byText.count()) await byText.click();
  }

  log('Aguardando autenticação corporativa. Se aparecer MFA/login, conclua manualmente nesta janela.');
  const deadline = Date.now() + LOGIN_TIMEOUT;
  while (Date.now() < deadline) {
    if (await page.getByText('Relatórios', { exact: true }).count()) return;
    if (/relatorios/i.test(page.url())) return;
    await sleep(1000);
  }
  throw new Error('Tempo excedido aguardando o login do Athena.');
}

async function telaRelatoriosPronta(page) {
  const busca = page.locator('input[placeholder*="Buscar" i]:visible').first();
  if (await busca.count() && await busca.isVisible().catch(()=>false)) return true;
  const cards = page.locator('chb-relatorio-favorito-title:visible');
  if (await cards.count().catch(()=>0)) return true;
  return false;
}

async function aguardarTelaRelatorios(page, timeout=TIMEOUT) {
  const fim = Date.now() + timeout;
  while (Date.now() < fim) {
    await tratarConfirmacaoEntradaAthena(page, 700).catch(()=>0);
    if (await telaRelatoriosPronta(page)) return true;
    await sleep(300);
  }
  return false;
}

async function openReports(page) {
  let ultimoErro = null;

  for (let tentativa=1; tentativa<=3; tentativa++) {
    try {
      await tratarConfirmacaoEntradaAthena(page, 1200).catch(()=>0);
      if (await aguardarTelaRelatorios(page, 1800)) {
        await saveAthenaAuthState(page);
        return;
      }

      const menu = page.getByText('Relatórios', { exact: true }).filter({visible:true}).first();
      if (await menu.count() && await menu.isVisible().catch(()=>false)) {
        log(`Abrindo Relatórios (tentativa ${tentativa}/3)...`);
        await menu.click({timeout:TIMEOUT});
      } else {
        log(`Menu Relatórios ainda não está disponível; recarregando a entrada do Athena (tentativa ${tentativa}/3)...`);
        await page.goto(config.athenaUrl, { waitUntil: 'domcontentloaded', timeout: LOGIN_TIMEOUT });
        await ensureAthenaLogin(page);
        const menu2 = page.getByText('Relatórios', { exact: true }).filter({visible:true}).first();
        if (await menu2.count() && await menu2.isVisible().catch(()=>false)) await menu2.click({timeout:TIMEOUT});
      }

      await tratarConfirmacaoEntradaAthena(page, 5500).catch(()=>0);
      if (await aguardarTelaRelatorios(page, Math.min(TIMEOUT, 15000))) {
        log('Tela de Relatórios carregada e pronta para busca.');
        await saveAthenaAuthState(page);
        return;
      }
      throw new Error('A URL abriu, mas a lista/busca de Relatórios não ficou pronta.');
    } catch (e) {
      ultimoErro = e;
      log(`Falha ao preparar Relatórios na tentativa ${tentativa}/3: ${e.message}`);
      await saveDebug(page, `relatorios_tentativa_${tentativa}`);
      if (tentativa < 3) {
        await page.goto(config.athenaUrl, { waitUntil: 'domcontentloaded', timeout: LOGIN_TIMEOUT }).catch(()=>{});
        await ensureAthenaLogin(page).catch(()=>{});
        await tratarConfirmacaoEntradaAthena(page, 3500).catch(()=>0);
      }
    }
  }

  throw new Error(`Não consegui deixar a tela de Relatórios pronta após 3 tentativas. ${ultimoErro?.message || ''}`.trim());
}

async function findAndOpenReport(page) {
  const name = config.relatorioNome;
  const code = name.split(' ')[0];

  if (!(await aguardarTelaRelatorios(page, TIMEOUT))) {
    throw new Error('A tela de Relatórios não está pronta para iniciar a busca.');
  }

  // Usa a busca apenas para reduzir a lista de relatórios.
  const search = page.locator('input[placeholder*="Buscar" i]').first();
  if (await search.count()) {
    await search.fill(code);
    await search.press('Enter').catch(()=>{});
    await page.locator('chb-relatorio-favorito-title').filter({hasText:code}).first().waitFor({state:'visible',timeout:TIMEOUT}).catch(()=>{});
  }

  // IMPORTANTE: no Athena a abertura do relatório é feita clicando NO TÍTULO
  // do relatório. Os ícones da direita são outras ações (editar/excluir/lupa)
  // e podem abrir a tela de Clonar. Por isso esta rotina NÃO clica em nenhum
  // botão de ação da linha.
  let clicked = false;

  // Seletor observado no DevTools do Athena:
  // <chb-relatorio-favorito-title> ... <div class="mt-1 ml-2">TÍTULO...</div>
  const component = page.locator('chb-relatorio-favorito-title')
    .filter({ hasText: code })
    .filter({ hasText: /TPO COMPLETO/i })
    .first();

  if (await component.count()) {
    const titleArea = component.locator('div.mt-1.ml-2').first();
    if (await titleArea.count()) {
      await titleArea.waitFor({ state: 'visible', timeout: TIMEOUT });
      log('Clicando diretamente no título do relatório (não na lupa).');
      await titleArea.click();
      clicked = true;
    }
  }

  // Fallback: clicar diretamente no texto do relatório. Continua sem tocar
  // nos ícones de ação da direita.
  if (!clicked) {
    let txt = page.getByText(name, { exact: false }).first();
    if (!(await txt.count())) {
      txt = page.getByText(/96192.*TPO COMPLETO RGB.*LN/i).first();
    }
    await txt.waitFor({ state: 'visible', timeout: TIMEOUT });
    log('Clicando diretamente no texto do relatório.');
    await txt.click();
    clicked = true;
  }

  if (!clicked) throw new Error('Não consegui clicar no título do relatório 96192.');

  // Em algumas entradas o Athena mostra uma confirmação intermediária antes de abrir o mês/relatório.
  // Se aparecer, confirma; se não aparecer, segue normalmente sem gerar timeout.
  await tratarConfirmacaoEntradaAthena(page, 8000).catch(()=>0);

  // A abertura pode carregar os dados por alguns segundos. Enquanto espera,
  // continua drenando qualquer confirmação tardia que apareça.
  const exportBtn = page.getByRole('button', { name: /^Exportar$/i }).last();
  const fimExport = Date.now() + TIMEOUT;
  while (Date.now() < fimExport) {
    if (await exportBtn.count() && await exportBtn.isVisible().catch(()=>false)) break;
    await tratarConfirmacaoEntradaAthena(page, 650).catch(()=>0);
    await sleep(250);
  }

  if (!(await exportBtn.count()) || !(await exportBtn.isVisible().catch(()=>false))) {
    // Diagnóstico adicional: se aparecer Clonar, sabemos que um seletor mudou.
    const cloneBtn = page.getByRole('button', { name: /^Clonar$/i }).first();
    if (await cloneBtn.count() && await cloneBtn.isVisible().catch(()=>false)) {
      throw new Error('O Athena abriu Clonar mesmo após clicar no título. O seletor do título mudou; envie uma foto do elemento destacado no DevTools.');
    }
    throw new Error('Cliquei no título do relatório, mas a janela de Exportar não abriu.');
  }

  log('Janela correta de exportação aberta.');
  await saveAthenaAuthState(page);
}

async function abrirRelatorioComRecuperacao(page) {
  let ultimoErro = null;
  for (let tentativa=1; tentativa<=3; tentativa++) {
    try {
      await openReports(page);
      await findAndOpenReport(page);
      return;
    } catch (e) {
      ultimoErro = e;
      log(`Falha ao abrir o relatório na tentativa ${tentativa}/3: ${e.message}`);
      await saveDebug(page, `abrir_relatorio_tentativa_${tentativa}`);
      if (tentativa < 3) {
        log('Reiniciando a etapa Athena -> Relatórios -> busca, sem continuar do ponto quebrado...');
        await page.goto(config.athenaUrl, { waitUntil: 'domcontentloaded', timeout: LOGIN_TIMEOUT }).catch(()=>{});
        await ensureAthenaLogin(page).catch(()=>{});
        await tratarConfirmacaoEntradaAthena(page, 4500).catch(()=>0);
      }
    }
  }
  throw new Error(`Não consegui abrir ${config.relatorioNome} após 3 tentativas completas. ${ultimoErro?.message || ''}`.trim());
}

async function setNativeDate(input, value) {
  // Playwright fill dispara os eventos que o Angular usa no ngModel.
  // A V3 alterava só o valor DOM e o Athena mantinha a data anterior no modelo.
  await input.waitFor({ state: 'visible', timeout: TIMEOUT });
  await input.click();
  await input.fill(value);
  await input.dispatchEvent('input').catch(()=>{});
  await input.dispatchEvent('change').catch(()=>{});
  await input.press('Tab').catch(()=>{});
  await sleep(350);
  const actual = await input.inputValue();
  if (actual !== value) {
    throw new Error(`Falha ao preencher data no Athena. Esperado ${value}, campo ficou ${actual || '(vazio)'}.`);
  }
}

async function selectPackaging(page) {
  const value = config.areaFuncional;

  // Primeiro tenta localizar o input dentro do bloco que contém "Área Funcional".
  let areaInput = page.locator('input:not([type="date"]):visible').first();
  const areaLabel = page.getByText(/Área Funcional/i).first();
  if (await areaLabel.count()) {
    const candidate = areaLabel.locator('xpath=ancestor::*[self::div or self::label][1]//input[not(@type="date")]').first();
    if (await candidate.count()) areaInput = candidate;
  }

  await areaInput.click();
  await areaInput.fill(value).catch(()=>{});
  await sleep(600);

  // Opção do autocomplete.
  const option = page.getByText(new RegExp(`^${value}$`, 'i')).last();
  if (await option.count()) {
    await option.click();
  } else {
    await areaInput.press('ArrowDown').catch(()=>{});
    await areaInput.press('Enter').catch(()=>{});
  }
  await sleep(500);
}

async function configureReport(page) {
  await tratarConfirmacaoEntradaAthena(page, 1200).catch(()=>false);
  const range = expectedCurrentRange();
  const today = range.today;
  const first = range.first;

  const dataInicio = page.getByText('Data Início', { exact: false }).first();
  if (await dataInicio.count()) await dataInicio.click().catch(()=>{});

  const dateInputs = page.locator('input[type="date"]:visible');
  const count = await dateInputs.count();
  if (count < 2) throw new Error(`Não encontrei os dois campos de data no relatório (encontrei ${count}).`);
  await setNativeDate(dateInputs.nth(0), ymd(first));
  await setNativeDate(dateInputs.nth(1), ymd(today));

  const deFinal = await dateInputs.nth(0).inputValue();
  const ateFinal = await dateInputs.nth(1).inputValue();
  if (deFinal !== ymd(first) || ateFinal !== ymd(today)) {
    throw new Error(`O Athena não manteve o período. De=${deFinal}, Até=${ateFinal}.`);
  }
  log(`Período confirmado na tela: ${dmy(first)} até ${dmy(today)}.`);

  await selectPackaging(page);
  log('Área Funcional configurada como Packaging. Linha NÃO será preenchida no TPO.');
  return range;
}

async function exportAthena(page, expectedRange) {
  const start = Date.now();
  const exportBtn = page.getByRole('button', { name: /^Exportar$/i }).last();
  await exportBtn.waitFor({ state: 'visible', timeout: TIMEOUT });

  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 120000 }),
    exportBtn.click()
  ]);

  const suggested = download.suggestedFilename() || `TPO_${ymd(new Date())}.xlsx`;
  let filename = suggested;
  if (!/\.xlsx?$/i.test(filename)) filename += '.xlsx';
  const target = path.join(downloadsDir, filename);
  try { if (fs.existsSync(target)) fs.rmSync(target, { force: true }); } catch (_) {}
  await download.saveAs(target);
  await download.path().catch(()=>null);

  if (!fs.existsSync(target) || fs.statSync(target).size < 1000) {
    const fallback = newestExcelInDownloads(start);
    if (fallback) { validateAthenaFilenameRange(fallback, expectedRange); return fallback; }
    throw new Error('O Athena iniciou a exportação, mas o Excel não apareceu corretamente em Downloads.');
  }
  validateAthenaFilenameRange(target, expectedRange);
  return target;
}

function copyExcelToText(excelPath) {
  const temp = path.join(os.tmpdir(), `tpo_clipboard_${Date.now()}.txt`);
  const ps = path.join(ROOT, 'copiar_excel.ps1');
  execFileSync('powershell.exe', [
    '-NoProfile','-STA','-ExecutionPolicy','Bypass','-File', ps,
    '-Arquivo', excelPath,
    '-Saida', temp
  ], { stdio: 'inherit', windowsHide: false });

  if (!fs.existsSync(temp)) throw new Error('O PowerShell não gerou o texto copiado do Excel.');
  const text = fs.readFileSync(temp, 'utf8').replace(/^\uFEFF/, '');
  try { fs.rmSync(temp, { force: true }); } catch (_) {}
  if (!text.trim()) throw new Error('O Excel foi copiado, mas o conteúdo ficou vazio.');
  return text;
}

async function openTpo(page) {
  const portalPath = config.portalTpoPath;
  if (!fs.existsSync(portalPath)) throw new Error(`Index do Portal TPO não encontrado: ${portalPath}`);

  await autorizarPortal504(page);
  await page.goto(pathToFileURL(portalPath).href, { waitUntil: 'domcontentloaded', timeout: LOGIN_TIMEOUT });
  await sleep(1200);

  // O guard do portal pode redirecionar para o login central.
  const deadline = Date.now() + LOGIN_TIMEOUT;
  while (Date.now() < deadline) {
    if (await page.locator('button[data-page="update"]').count()) return;
    if (/portal indicadores/i.test((await page.title().catch(()=>''))) || /retorno=/i.test(page.url())) {
      log('Portal TPO redirecionou para o login central; reaplicando autorização automática...'); await autorizarPortal504(page); await page.goto(pathToFileURL(portalPath).href,{waitUntil:'domcontentloaded',timeout:LOGIN_TIMEOUT}).catch(()=>{});
    }
    await sleep(1000);
  }
  throw new Error('Não consegui abrir o Portal TPO após aguardar o login central.');
}

async function updateTpo(page, pastedText) {
  await page.locator('button[data-page="update"]').click();
  await page.locator('#update').waitFor({ state: 'visible', timeout: TIMEOUT });

  const month = currentMonthLabel();
  const monthSelect = page.locator('#importMonth');
  await monthSelect.selectOption({ label: month }).catch(async()=>{
    await monthSelect.selectOption(month);
  });
  log(`Mês de destino selecionado inicialmente: ${month}.`);

  await page.locator('#pasteBox').evaluate((el, value) => {
    el.value = value;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }, pastedText);

  await page.locator('#analyzePaste').click();
  await sleep(1400);
  let msg = (await page.locator('#importMsg').innerText().catch(()=>'')) || '';
  if (/erro|inválid|não reconhe|antes\./i.test(msg)) {
    throw new Error(`O Portal TPO não aceitou o relatório ao analisar. Mensagem: ${msg}`);
  }

  const inferredMonth = await monthSelect.inputValue();
  const preview = (await page.locator('#importPreview').innerText().catch(()=>'')) || '';
  log(`Mês predominante identificado pelo relatório: ${inferredMonth || '(não identificado)'}.`);

  // A partir da V6 o relatório do TPO começa no ÚLTIMO DIA DO MÊS ANTERIOR.
  // Portanto o mês predominante inferido pelo portal pode, principalmente no dia 1º,
  // ser o mês anterior. Isso é esperado e NÃO significa que devemos atualizar o mês anterior.
  // Reforçamos explicitamente o mês atual como destino. O próprio Portal TPO filtra por
  // Data coleta e ignora registros de outros meses ao aplicar a atualização.
  if (inferredMonth !== month) {
    log(`Sobreposição entre meses detectada: o portal inferiu ${inferredMonth || 'desconhecido'}, mas o destino será mantido em ${month}.`);
    if (preview) log(`Prévia do relatório: ${preview.replace(/\s+/g,' ').trim().slice(0,500)}`);
  }

  await monthSelect.selectOption(month);
  if ((await monthSelect.inputValue()) !== month) {
    throw new Error(`Não consegui manter ${month} como mês de destino no Portal TPO.`);
  }

  // O botão "Atualizar mês" abre um confirm() nativo do navegador.
  // IMPORTANTE: o diálogo precisa ser ACEITO dentro do próprio evento 'dialog'.
  // Se esperarmos o click() terminar para só depois aceitar, o clique fica bloqueado
  // pelo confirm() e o Playwright estoura timeout (foi o que aconteceu na V4).
  const monthRe = new RegExp(`\\b${month}\\b`, 'i');
  let dialogMsg = '';
  let dialogAccepted = false;

  const dialogHandled = new Promise((resolve) => {
    page.once('dialog', async (dialog) => {
      dialogMsg = dialog.message();
      log(`Confirmação do Portal TPO: ${dialogMsg}`);
      try {
        if (monthRe.test(dialogMsg)) {
          await dialog.accept();
          dialogAccepted = true;
        } else {
          await dialog.dismiss();
        }
      } finally {
        resolve();
      }
    });
  });

  await page.locator('#replaceMonth').click({ timeout: TIMEOUT });
  await Promise.race([
    dialogHandled,
    sleep(TIMEOUT).then(() => { throw new Error('O botão Atualizar mês foi clicado, mas a confirmação não apareceu.'); })
  ]);

  if (!dialogAccepted) {
    throw new Error(`BLOQUEIO DE SEGURANÇA: a confirmação tentou atualizar outro mês. Esperado ${month}; mensagem recebida: ${dialogMsg || '(sem mensagem)'}`);
  }

  const deadline = Date.now() + TIMEOUT;
  while (Date.now() < deadline) {
    msg = (await page.locator('#importMsg').innerText().catch(()=>'')) || '';
    if (/atualizado com segurança|foi atualizado/i.test(msg)) break;
    if (/erro|inválid|falha/i.test(msg)) throw new Error(`Falha ao atualizar mês no Portal TPO: ${msg}`);
    await sleep(500);
  }
  if (!/atualizado com segurança|foi atualizado/i.test(msg)) {
    throw new Error(`O Portal TPO não confirmou a atualização do mês. Mensagem atual: ${msg}`);
  }
  if (!new RegExp(`\\b${month}\\b`, 'i').test(msg)) {
    throw new Error(`BLOQUEIO DE SEGURANÇA: o portal confirmou atualização, porém a mensagem não cita ${month}: ${msg}`);
  }
  log(msg.replace(/\s+/g,' ').trim());
}

function validateIndex(file) {
  if (!fs.existsSync(file)) throw new Error(`Novo index não encontrado: ${file}`);
  const st = fs.statSync(file);
  if (st.size < 100000) throw new Error(`Novo index.html parece incompleto (${st.size} bytes).`);
  const head = fs.readFileSync(file, 'utf8');
  if (!/PORTAL TPO|Portal TPO/i.test(head) || !/DATA_START|let DATA=/i.test(head)) {
    throw new Error('O arquivo gerado não passou na validação básica do Portal TPO.');
  }
}

function backupProductionIndex() {
  const dest = config.portalTpoPath;
  if (!fs.existsSync(dest)) return null;
  const backupDir = path.join(path.dirname(dest), 'Backups_Automacao_TPO');
  fs.mkdirSync(backupDir, { recursive: true });
  const backup = path.join(backupDir, `index_${nowStamp()}.html`);
  fs.copyFileSync(dest, backup);
  validateIndex(backup);
  log(`Backup do index anterior preservado em: ${backup}`);
  return backup;
}

function replaceProductionIndex(newIndex) {
  validateIndex(newIndex);
  const dest = config.portalTpoPath;
  const dir = path.dirname(dest);
  const tempDest = path.join(dir, 'index.__novo__.html');
  const oldTemp = path.join(dir, 'index.__anterior__.html');

  const permanentBackup = backupProductionIndex();

  try { fs.rmSync(tempDest, { force: true }); } catch (_) {}
  try { fs.rmSync(oldTemp, { force: true }); } catch (_) {}

  fs.copyFileSync(newIndex, tempDest);
  validateIndex(tempDest);

  if (fs.existsSync(dest)) fs.renameSync(dest, oldTemp);
  try {
    fs.renameSync(tempDest, dest);
    validateIndex(dest);
    if (fs.existsSync(oldTemp)) fs.rmSync(oldTemp, { force: true });
    fs.rmSync(newIndex, { force: true });
    if (permanentBackup) log('Substituição concluída; o backup permanente foi mantido.');
  } catch (err) {
    if (!fs.existsSync(dest) && fs.existsSync(oldTemp)) fs.renameSync(oldTemp, dest);
    throw err;
  }
}

async function downloadGeneratedIndex(page, targetPath) {
  try { fs.rmSync(targetPath, { force: true }); } catch (_) {}
  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 120000 }),
    page.locator('#generateIndex').click()
  ]);
  await download.saveAs(targetPath);
  validateIndex(targetPath);
  return targetPath;
}

async function generateTestIndex(page) {
  const target = path.join(downloadsDir, `index_TPO_TESTE_${nowStamp()}.html`);
  await downloadGeneratedIndex(page, target);
  log(`Index de TESTE gerado em Downloads: ${target}`);
  log('O index.html da pasta do Portal TPO NÃO foi substituído neste teste.');
  return target;
}

async function generateAndReplaceIndex(page) {
  const downloadPath = path.join(downloadsDir, 'index.html');
  await downloadGeneratedIndex(page, downloadPath);
  replaceProductionIndex(downloadPath);
}

async function main() {
  const total = MODE === 'athena' ? 6 : MODE === 'portal' ? 8 : 12;
  let browser = null;
  let page = null;
  let downloadedExcel = null;
  try {
    step(1, total, 'Abrindo Chrome controlado pela automação...');
    ({ browser, page } = await launchAutomationBrowser());

    if (MODE !== 'portal') {
      step(2, total, 'Abrindo Athena...');
      ({ browser, page } = await openAthenaWithOneRecovery(browser, page));

      step(3, total, 'Abrindo Relatórios e aguardando a lista ficar pronta...');
      await openReports(page);

      step(4, total, `Localizando ${config.relatorioNome} com recuperação automática...`);
      await abrirRelatorioComRecuperacao(page);

      step(5, total, 'Configurando período do mês atual e Área Funcional = Packaging...');
      const expectedRange = await configureReport(page);

      step(6, total, 'Exportando Excel do Athena para Downloads...');
      downloadedExcel = await exportAthena(page, expectedRange);
      log(`Excel baixado: ${downloadedExcel}`);

      if (MODE === 'athena') {
        log('TESTE ATHENA concluído. O relatório foi baixado e nenhuma alteração foi feita no Portal TPO.');
        return;
      }
    } else {
      // Modo de teste do portal: usa o Excel mais recente de Downloads.
      downloadedExcel = newestExcelInDownloads(0, true);
      if (!downloadedExcel) throw new Error('Não encontrei nenhum relatorio-generico_*.xlsx em Downloads para testar o Portal TPO.');
      validateAthenaFilenameRange(downloadedExcel, expectedCurrentRange());
      step(2, total, `Usando o relatório TPO mais recente de Downloads: ${path.basename(downloadedExcel)}`);
    }

    const baseStep = MODE === 'portal' ? 3 : 7;
    step(baseStep, total, 'Abrindo Excel, selecionando todos os dados e copiando...');
    const pastedText = copyExcelToText(downloadedExcel);
    log(`Dados copiados do Excel: ${pastedText.length.toLocaleString('pt-BR')} caracteres.`);

    step(baseStep + 1, total, 'Abrindo o index.html atual do Portal TPO...');
    await openTpo(page);

    step(baseStep + 2, total, `Atualizar Base -> ${currentMonthLabel()} -> colar -> Analisar relatório...`);
    await updateTpo(page, pastedText);

    if (MODE === 'portal') {
      step(baseStep + 3, total, 'Gerando index.html de TESTE em Downloads (sem substituir produção)...');
      const testIndex = await generateTestIndex(page);

      step(baseStep + 4, total, 'Validando o index.html de TESTE...');
      validateIndex(testIndex);

      step(baseStep + 5, total, 'TESTE DO PORTAL CONCLUÍDO COM SUCESSO.');
      log('Nenhum arquivo da pasta de produção foi substituído neste teste.');
    } else {
      step(baseStep + 3, total, 'Gerando o novo index.html, criando backup e substituindo produção...');
      await generateAndReplaceIndex(page);

      step(baseStep + 4, total, 'Validando o index.html substituído na pasta do Portal TPO...');
      validateIndex(config.portalTpoPath);

      step(baseStep + 5, total, 'CONCLUÍDO COM SUCESSO.');
      log(`Portal atualizado: ${config.portalTpoPath}`);
    }
  } catch (err) {
    console.error(`\nERRO: ${err.message}`);
    await saveDebug(page, 'erro');
    console.error('\nA automação bloqueia mês divergente e, no modo completo, cria backup permanente antes de substituir o index.html.');
    process.exitCode = 1;
    if (page) {
      log('A janela ficará aberta por 60 segundos para você fotografar a tela/DevTools se precisar.');
      await sleep(60000).catch(()=>{});
    }
  } finally {
    if (page && !config.manterBrowserAbertoAoFinal) {
      try { await page.context().close(); } catch (_) {}
    }
    if (browser && !config.manterBrowserAbertoAoFinal) {
      try { await browser.close(); } catch (_) {}
    }
  }
}

main();
