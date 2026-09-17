# -*- coding: utf-8 -*-
"""
Automação - Consulta de Propostas Canceladas/Integradas (WebFI / Esteira)
============================================================================

Interface corporativa em formato "wizard".
A lógica de automação permanece intacta; apenas a interface foi refinada.

Versão 2.2 (resiliente + restart automático):
- Mantém toda a interface melhorada da versão 2.0.
- Tratamento resiliente por proposta (não derruba tudo se uma falhar).
- Salvamento incremental do arquivo de saída durante a execução.
- Reinício automático do Chrome/Selenium se a sessão morrer no meio da execução.
- Mantida a velocidade: sem esperas extras no fluxo principal.

Observação: o reinício automático tenta reaproveitar a mesma sessão de login,
refazendo navegação para a tela de consulta quando necessário.
"""

import re
import os
import json
import time
import threading
import traceback
import unicodedata
import concurrent.futures
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
    WebDriverException,
    SessionNotCreatedException,
    JavascriptException,
)

from openpyxl import Workbook, load_workbook

# --------------------------------------------------------------------------- #
# CONFIGURAÇÕES GERAIS
# --------------------------------------------------------------------------- #

LOGIN_URL = "URL DO SITE"
DEFAULT_TIMEOUT = 7
MENU_TIMEOUT = 6
JS_TIMEOUT = 5
JS_POLL = 0.12
WAIT_POLL_FREQUENCY = 0.1
CHROME_STARTUP_TIMEOUT = 30

MOTIVO_REGEX = re.compile(r"Motivo\(s\)\s*do\s*Cancelamento:\s*(.+)", re.IGNORECASE)
QUOTED_REGEX = re.compile(r"'([^']*)'")
USERNAME_REGEX = re.compile(r"^[A-ZÀ-Ú0-9]+(?:\.[A-ZÀ-Ú0-9]+)+$", re.IGNORECASE)

SCRIPT_DIR = Path(__file__).resolve().parent
LOG_FILE = SCRIPT_DIR / "esteira_automation.log"
CONFIG_FILE = SCRIPT_DIR / "esteira_config.json"
DIAG_DIR = SCRIPT_DIR / "diagnosticos"

# Número máximo de tentativas de reinício completo do navegador
MAX_BROWSER_RESTARTS = 3

os.environ.setdefault("SE_AVOID_STATS", "true")


def log_para_arquivo(msg: str):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except Exception:
        pass


def _normalizar(texto: str) -> str:
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    return sem_acento.strip().upper()


def _normalizar_coluna(texto: str) -> str:
    n = _normalizar(texto)
    return re.sub(r"[^A-Z0-9]", "", n)


def _parece_usuario(texto: str) -> bool:
    t = (texto or "").strip()
    if not t or ":" in t or "/" in t:
        return False
    if not USERNAME_REGEX.match(t):
        return False
    if t.replace(".", "").isdigit():
        return False
    return True


def _limpar_texto(texto: str) -> str:
    if texto is None:
        return ""
    t = str(texto).replace("\xa0", " ").strip()
    t = t.strip("'").strip('"').strip("[](),")
    t = re.sub(r"\s+", " ", t)
    return t.strip()


# --------------------------------------------------------------------------- #
# SCRIPTS JAVASCRIPT (INALTERADOS)
# --------------------------------------------------------------------------- #

JS_WALK_HELPER = """
function _fiWalkDocs(doc, visited) {
    if (!doc || visited.has(doc)) return [];
    visited.add(doc);
    var docs = [doc];
    var frames = [];
    try { frames = Array.prototype.slice.call(doc.querySelectorAll('iframe, frame')); } catch(e) { frames = []; }
    for (var i = 0; i < frames.length; i++) {
        var subdoc = null;
        try { subdoc = frames[i].contentDocument || (frames[i].contentWindow && frames[i].contentWindow.document); } catch(e) { subdoc = null; }
        if (subdoc) { docs = docs.concat(_fiWalkDocs(subdoc, visited)); }
    }
    return docs;
}
"""

JS_LER_OBSERVACOES = JS_WALK_HELPER + """
var docs = _fiWalkDocs(document, new Set());
for (var i = 0; i < docs.length; i++) {
    var d = docs[i];
    var el = null;
    try {
        el = d.querySelector("#ctl00_cph_UcObs_UcObs_txtObs_CAMPO") ||
             d.querySelector("textarea[name*='txtObs']") ||
             d.querySelector("textarea[id*='txtObs']") ||
             d.querySelector("textarea.FITxtArea") ||
             d.querySelector("textarea[readonly]");
    } catch(e) { el = null; }
    if (el) {
        var v = el.value;
        if (v === undefined || v === null || v === '') { v = el.textContent || ''; }
        return v;
    }
}
return null;
"""

JS_CLICAR_POR_ID_OU_TEXTO = JS_WALK_HELPER + """
var idExact = arguments[0];
var textAllRequired = arguments[1] || [];
var docs = _fiWalkDocs(document, new Set());
for (var i = 0; i < docs.length; i++) {
    var d = docs[i];
    var links = [];
    try { links = Array.prototype.slice.call(d.querySelectorAll('a')); } catch(e) { links = []; }
    for (var j = 0; j < links.length; j++) {
        var a = links[j];
        var idMatch = idExact && a.id === idExact;
        var txt = (a.innerText || a.textContent || '').trim().toUpperCase();
        var txtMatch = false;
        if (textAllRequired.length > 0) {
            txtMatch = true;
            for (var k = 0; k < textAllRequired.length; k++) {
                if (txt.indexOf(textAllRequired[k].toUpperCase()) === -1) { txtMatch = false; break; }
            }
        }
        if (idMatch || txtMatch) {
            try { a.scrollIntoView({block:'center'}); } catch(e) {}
            try { a.click(); } catch(e) { continue; }
            return true;
        }
    }
}
return false;
"""

JS_EXTRAIR_TABELA_ATIVIDADES = JS_WALK_HELPER + """
function _fiNormColuna(s) {
    s = (s || '').toString();
    try { s = s.normalize('NFKD').replace(/[\\u0300-\\u036f]/g, ''); } catch(e) {}
    return s.toUpperCase().replace(/[^A-Z0-9]/g, '');
}
var CABECALHOS_ALVO = ['DESCATV', 'USUARIOINICIAL', 'USUARIOFINAL', 'ATIVIDADE'];
var docs = _fiWalkDocs(document, new Set());
var melhor = null, melhorScore = -1;
for (var i = 0; i < docs.length; i++) {
    var d = docs[i];
    var tables = [];
    try { tables = Array.prototype.slice.call(d.querySelectorAll('table')); } catch(e) { tables = []; }
    for (var k = 0; k < tables.length; k++) {
        var t = tables[k];
        var linhasCab = Array.prototype.slice.call(t.querySelectorAll('tr')).slice(0, 3);
        var headerScore = 0;
        for (var r = 0; r < linhasCab.length; r++) {
            var celulas = Array.prototype.slice.call(linhasCab[r].querySelectorAll('th,td'));
            for (var c = 0; c < celulas.length; c++) {
                var norm = _fiNormColuna(celulas[c].innerText || celulas[c].textContent || '');
                if (CABECALHOS_ALVO.indexOf(norm) !== -1) { headerScore += 100; }
            }
        }
        var txt = (t.innerText || t.textContent || '').toUpperCase();
        var textScore = (txt.split('CANCELADA').length - 1) * 10 + (txt.split('APROVADA').length - 1);
        var score = headerScore + textScore;
        if (headerScore === 0 && txt.indexOf('CANCELADA') === -1 && txt.indexOf('APROVADA') === -1) continue;
        if (score > melhorScore) { melhorScore = score; melhor = t; }
    }
}
if (!melhor) return null;
var rows = Array.prototype.slice.call(melhor.querySelectorAll('tr')).map(function(tr) {
    return Array.prototype.slice.call(tr.querySelectorAll('th,td')).map(function(cel) {
        return (cel.innerText || cel.textContent || '').trim();
    });
});
return rows;
"""

JS_FECHAR_MODAL = JS_WALK_HELPER + """
var docs = _fiWalkDocs(document, new Set());
for (var i = 0; i < docs.length; i++) {
    var d = docs[i];
    var candidatos = [];
    try { candidatos = Array.prototype.slice.call(d.querySelectorAll('#btnFechar_txt, a')); } catch(e) { candidatos = []; }
    for (var j = 0; j < candidatos.length; j++) {
        var el = candidatos[j];
        var txt = (el.innerText || el.textContent || '').trim().toUpperCase();
        var rects = el.getClientRects ? el.getClientRects() : [];
        var visivel = !!(el.offsetWidth || el.offsetHeight || rects.length);
        if ((el.id === 'btnFechar_txt' || txt === 'FECHAR') && visivel) {
            try { el.click(); } catch(e) { continue; }
            return true;
        }
    }
}
return false;
"""

JS_ACHAR_HREF_APROVACAO_CONSULTA = JS_WALK_HELPER + """
var docs = _fiWalkDocs(document, new Set());
for (var i = 0; i < docs.length; i++) {
    var d = docs[i];
    var links = [];
    try { links = Array.prototype.slice.call(d.querySelectorAll('a')); } catch(e) { links = []; }
    for (var j = 0; j < links.length; j++) {
        var a = links[j];
        var idExato = a.id === 'WFP2010_PWCNPROPCI';
        var hrefAttr = a.getAttribute('href') || '';
        var hrefExato = hrefAttr.indexOf('UI.AprovacaoConsultaCanInt.aspx') !== -1;
        var txt = (a.innerText || a.textContent || '').trim().toUpperCase();
        var txtExato = txt.indexOf('CANCELADAS') !== -1 && txt.indexOf('INTEGRADAS') !== -1;
        if (idExato || hrefExato || txtExato) {
            if (hrefAttr.toLowerCase().indexOf('javascript') === 0) { return ""; }
            return a.href || hrefAttr;
        }
    }
}
return null;
"""


# --------------------------------------------------------------------------- #
# CONFIG PERSISTENTE / UTILITÁRIOS (INALTERADOS)
# --------------------------------------------------------------------------- #

def carregar_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def salvar_config(dados: dict):
    try:
        atual = carregar_config()
        atual.update(dados)
        CONFIG_FILE.write_text(json.dumps(atual, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _caminho_e_de_rede(caminho: str) -> bool:
    if not caminho:
        return False
    p = str(caminho)
    if p.startswith("\\\\"):
        return True
    try:
        import ctypes

        DRIVE_REMOTE = 4
        drive = os.path.splitdrive(p)[0]
        if drive:
            tipo = ctypes.windll.kernel32.GetDriveTypeW(drive + "\\")
            return tipo == DRIVE_REMOTE
    except Exception:
        pass
    return False


def _pastas_de_busca():
    home = Path.home()
    pastas = [
        SCRIPT_DIR,
        SCRIPT_DIR.parent,
        Path("C:/Automacao"),
        Path("C:/chrome-automacao"),
        home / "Desktop",
        home / "Downloads",
        home / "Documents",
        Path("C:/Program Files/Google/Chrome/Application"),
        Path("C:/Program Files (x86)/Google/Chrome/Application"),
        home / "AppData/Local/Google/Chrome/Application",
    ]
    return [p for p in pastas if p.exists()]


def buscar_arquivo_automatico(nome_arquivo: str, profundidade_max: int = 3):
    nome_arquivo = nome_arquivo.lower()
    for base in _pastas_de_busca():
        if _caminho_e_de_rede(str(base)):
            continue
        try:
            base_depth = len(base.parts)
            for root, dirs, files in os.walk(base):
                depth = len(Path(root).parts) - base_depth
                if depth >= profundidade_max:
                    dirs[:] = []
                for f in files:
                    if f.lower() == nome_arquivo:
                        return str(Path(root) / f)
        except (PermissionError, OSError):
            continue
    return None


def auto_detectar_caminhos(chrome_salvo: str, driver_salvo: str):
    chrome_path = chrome_salvo if (chrome_salvo and Path(chrome_salvo).exists()) else None
    driver_path = driver_salvo if (driver_salvo and Path(driver_salvo).exists()) else None
    if not chrome_path:
        chrome_path = buscar_arquivo_automatico("chrome.exe")
    if not driver_path:
        driver_path = buscar_arquivo_automatico("chromedriver.exe")
    return chrome_path or "", driver_path or ""


class ConfiguracaoInvalidaError(Exception):
    pass


def validar_caminhos_chrome(chromedriver_path: str, chrome_binary_path: str):
    def nome(p):
        return Path(p).name.lower() if p else ""

    avisos_rede = []

    if chromedriver_path:
        p = Path(chromedriver_path)
        if p.is_dir():
            raise ConfiguracaoInvalidaError(
                f"'{chromedriver_path}' é uma PASTA, selecione o arquivo chromedriver.exe."
            )
        if not p.exists():
            raise ConfiguracaoInvalidaError(
                f"O arquivo do chromedriver não existe: {chromedriver_path}"
            )
        if nome(chromedriver_path) == "chrome.exe":
            raise ConfiguracaoInvalidaError(
                "Você colocou o 'chrome.exe' no campo do 'chromedriver.exe' - são arquivos diferentes!"
            )
        if _caminho_e_de_rede(chromedriver_path):
            avisos_rede.append(chromedriver_path)

    if chrome_binary_path:
        p = Path(chrome_binary_path)
        if p.is_dir():
            raise ConfiguracaoInvalidaError(
                f"'{chrome_binary_path}' é uma PASTA, selecione o arquivo chrome.exe."
            )
        if not p.exists():
            raise ConfiguracaoInvalidaError(
                f"O arquivo do chrome.exe não existe: {chrome_binary_path}"
            )
        if nome(chrome_binary_path).startswith("chromedriver"):
            raise ConfiguracaoInvalidaError(
                "Você colocou o 'chromedriver.exe' no campo do 'chrome.exe' - são arquivos diferentes!"
            )
        if _caminho_e_de_rede(chrome_binary_path):
            avisos_rede.append(chrome_binary_path)

    if avisos_rede:
        raise ConfiguracaoInvalidaError(
            "Os seguintes caminhos estão em um DRIVE DE REDE/MAPEADO:\n"
            + "\n".join(f" - {c}" for c in avisos_rede)
            + "\n\nCopie 'chrome-win64' e 'chromedriver-win64' para um disco LOCAL (ex: C:\\Automacao\\)."
        )


# --------------------------------------------------------------------------- #
# CAMADA DE AUTOMAÇÃO (resiliente + restart)
# --------------------------------------------------------------------------- #


class EsteiraBot:
    def __init__(
        self,
        usuario: str,
        senha: str,
        log_callback=None,
        chromedriver_path: str = "",
        chrome_binary_path: str = "",
    ):
        self.usuario = usuario
        self.senha = senha
        self.log_callback = log_callback or (lambda msg: None)
        self.driver = None
        self.wait = None
        self.chromedriver_path = chromedriver_path.strip() or None
        self.chrome_binary_path = chrome_binary_path.strip() or None
        self.restarts_feitos = 0

    # ------------------------ LOG / HELPERS ------------------------ #

    def log(self, msg: str):
        print(msg)
        log_para_arquivo(msg)
        self.log_callback(msg)

    def salvar_diagnostico(self, contexto: str):
        try:
            DIAG_DIR.mkdir(exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            base = DIAG_DIR / f"{ts}_{contexto}"
            try:
                self.driver.switch_to.default_content()
            except Exception:
                pass
            try:
                self.driver.save_screenshot(str(base.with_suffix(".png")))
            except Exception:
                pass
            try:
                base.with_suffix(".html").write_text(
                    self.driver.page_source, encoding="utf-8"
                )
            except Exception:
                pass
            self.log(f"Diagnóstico salvo em: {base}.png / .html")
        except Exception as e:
            self.log(f"Não foi possível salvar diagnóstico: {e}")

    def _exec_js(self, script: str, args: tuple = ()):
        try:
            self.driver.switch_to.default_content()
        except Exception:
            pass
        try:
            return self.driver.execute_script(script, *args)
        except (JavascriptException, WebDriverException):
            return None

    def _exec_js_retry(
        self,
        script: str,
        args: tuple = (),
        timeout: float = JS_TIMEOUT,
        poll: float = JS_POLL,
        sucesso_check=None,
    ):
        if sucesso_check is None:
            sucesso_check = lambda r: r is not None and r is not False

        fim = time.time() + timeout
        resultado = None
        while True:
            resultado = self._exec_js(script, args)
            if sucesso_check(resultado):
                return resultado
            if time.time() >= fim:
                return resultado
            time.sleep(poll)

    # ------------------------ BROWSER LIFECYCLE ------------------------ #

    def _montar_options(self):
        options = Options()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.page_load_strategy = "eager"
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)

        if self.chrome_binary_path:
            self.log(f"Usando Chrome em: {self.chrome_binary_path}")
            options.binary_location = self.chrome_binary_path
        else:
            self.log("Nenhum chrome.exe informado - detecção automática do Selenium.")
        return options

    def _criar_driver_bloqueante(self, options):
        service = None
        if self.chromedriver_path:
            self.log(f"Usando chromedriver manual em: {self.chromedriver_path}")
            service = Service(executable_path=self.chromedriver_path)
        else:
            self.log("Nenhum chromedriver manual informado - detecção automática.")
        if service is not None:
            return webdriver.Chrome(service=service, options=options)
        return webdriver.Chrome(options=options)

    def start_browser(self):
        self.log("Validando caminhos informados...")
        try:
            validar_caminhos_chrome(
                self.chromedriver_path or "", self.chrome_binary_path or ""
            )
        except ConfiguracaoInvalidaError as e:
            self.log(f"CONFIGURAÇÃO INVÁLIDA: {e}")
            raise RuntimeError(str(e)) from e

        options = self._montar_options()
        self.log(f"Iniciando o Chrome (timeout de {CHROME_STARTUP_TIMEOUT}s)...")

        executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        future = executor.submit(self._criar_driver_bloqueante, options)

        try:
            self.driver = future.result(timeout=CHROME_STARTUP_TIMEOUT)
        except concurrent.futures.TimeoutError:
            raise RuntimeError(
                f"O Chrome não abriu em {CHROME_STARTUP_TIMEOUT}s. Verifique se os caminhos estão em disco LOCAL."
            )
        except SessionNotCreatedException as e:
            raise RuntimeError(
                "Versão do chromedriver incompatível com o chrome.exe. Baixe as DUAS peças na MESMA versão em https://googlechromelabs.github.io/chrome-for-testing/.\n\n"
                + str(e)
            ) from e
        except WebDriverException as e:
            msg = str(e)
            if any(
                s in msg.lower()
                for s in (
                    "unexpectedly exited",
                    "chrome not reachable",
                    "devtoolsactiveport",
                )
            ):
                dica = (
                    "\n\nCaminhos em DRIVE DE REDE? Copie para disco local."
                    if (
                        _caminho_e_de_rede(self.chrome_binary_path or "")
                        or _caminho_e_de_rede(self.chromedriver_path or "")
                    )
                    else ""
                )
                raise RuntimeError(
                    "O Chrome iniciou e fechou/crashou imediatamente."
                    + dica
                    + "\n\n"
                    + msg
                ) from e
            if "unable to obtain driver" in msg.lower() or "unable to discover" in msg.lower():
                raise RuntimeError(
                    "Selenium não conseguiu localizar/baixar o chromedriver. Informe os caminhos manualmente.\n\n"
                    + msg
                ) from e
            if "cannot find chrome binary" in msg.lower():
                raise RuntimeError(
                    "Chrome não encontrado. Informe o caminho do chrome.exe.\n\n" + msg
                ) from e
            raise RuntimeError(f"Falha ao iniciar o Chrome: {msg}") from e
        finally:
            executor.shutdown(wait=False)

        self.wait = WebDriverWait(
            self.driver, DEFAULT_TIMEOUT, poll_frequency=WAIT_POLL_FREQUENCY
        )
        self.log("Chrome iniciado com sucesso.")
        salvar_config(
            {
                "chrome_binary_path": self.chrome_binary_path or "",
                "chromedriver_path": self.chromedriver_path or "",
            }
        )

    def quit(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None
            self.wait = None

    def _reiniciar_navegador(self):
        """Fecha o navegador atual e cria um novo driver, relogando e voltando para a tela de consulta."""
        self.restarts_feitos += 1
        if self.restarts_feitos > MAX_BROWSER_RESTARTS:
            raise RuntimeError(
                f"Número máximo de reinícios do navegador atingido ({MAX_BROWSER_RESTARTS})."
            )
        self.log(
            f"Tentando reiniciar o navegador (tentativa {self.restarts_feitos}/{MAX_BROWSER_RESTARTS})..."
        )
        try:
            self.quit()
        except Exception:
            pass
        # pequena pausa para garantir que processos mortos sejam liberados
        time.sleep(2)
        self.start_browser()
        self.login()
        self.navegar_para_esteira()
        self.log("Reinício do navegador concluído com sucesso.")

    # ------------------------ FLUXO DE NAVEGAÇÃO ------------------------ #

    def _click(self, by, value, timeout=DEFAULT_TIMEOUT):
        el = WebDriverWait(
            self.driver,
            timeout,
            poll_frequency=WAIT_POLL_FREQUENCY,
        ).until(EC.element_to_be_clickable((by, value)))
        try:
            el.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", el)
        return el

    def _click_el(self, el):
        try:
            el.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", el)

    def _fill(self, by, value, text, timeout=DEFAULT_TIMEOUT):
        el = WebDriverWait(
            self.driver,
            timeout,
            poll_frequency=WAIT_POLL_FREQUENCY,
        ).until(EC.visibility_of_element_located((by, value)))
        el.clear()
        el.send_keys(text)
        return el

    def _switch_to_new_window_if_any(self, previous_handles, timeout=2.5):
        fim = time.time() + timeout
        while time.time() < fim:
            current_handles = self.driver.window_handles
            new_handles = [h for h in current_handles if h not in previous_handles]
            if new_handles:
                self.driver.switch_to.window(new_handles[-1])
                return True
            time.sleep(0.1)
        return False

    def login(self):
        self.log("Abrindo página de login...")
        self.driver.get(LOGIN_URL)
        self._fill(By.ID, "EUsuario_CAMPO", self.usuario)
        self._fill(By.ID, "ESenha_CAMPO", self.senha)
        self._click(By.ID, "lnkEntrar")
        self.log("Login enviado.")
        self.wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//div[contains(@class,'divTit') and normalize-space(text())='NEGOCIAÇÃO']",
                )
            )
        )

    def navegar_para_esteira(self):
        self.log("Abrindo menu NEGOCIAÇÃO...")
        negociacao = self.wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//div[contains(@class,'divTit') and normalize-space(text())='NEGOCIAÇÃO']",
                )
            )
        )
        self._click_el(negociacao)

        self.log("Clicando em Autorizador...")
        previous_handles = self.driver.window_handles
        autorizador = WebDriverWait(
            self.driver,
            MENU_TIMEOUT,
            poll_frequency=WAIT_POLL_FREQUENCY,
        ).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(normalize-space(.),'Autorizador')]")
            )
        )
        self._click_el(autorizador)
        self._switch_to_new_window_if_any(previous_handles)

        self.log(
            "Localizando link EXATO 'Aprovação/Consulta (Canceladas e Integradas)'..."
        )
        href = self._exec_js_retry(
            JS_ACHAR_HREF_APROVACAO_CONSULTA,
            timeout=MENU_TIMEOUT,
            sucesso_check=lambda r: r is not None,
        )

        if not href:
            self.salvar_diagnostico("link_aprovacao_consulta_nao_encontrado")
            raise RuntimeError(
                "Não foi possível localizar o link EXATO 'Aprovação/Consulta (Canceladas e Integradas)'. "
                "Screenshot/HTML salvos em 'diagnosticos'."
            )

        if href == "":
            self.log(
                "Link encontrado, mas é 'javascript:...' - tentando clique direto via JS..."
            )
            clicou = self._exec_js_retry(
                JS_CLICAR_POR_ID_OU_TEXTO,
                args=("WFP2010_PWCNPROPCI", ["Canceladas", "Integradas"]),
                timeout=MENU_TIMEOUT,
                sucesso_check=lambda r: r is True,
            )
            if not clicou:
                self.salvar_diagnostico("clique_aprovacao_consulta_falhou")
                raise RuntimeError(
                    "Encontrou o link mas o clique via JS falhou. Screenshot/HTML salvos em 'diagnosticos'."
                )
        else:
            self.log(f"Navegando diretamente para: {href}")
            self.driver.get(href)

        try:
            self.wait.until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_Cph_AprCons_txtPesquisa_CAMPO")
                )
            )
        except TimeoutException:
            self.salvar_diagnostico("tela_pesquisa_nao_carregou")
            raise RuntimeError(
                "Após navegar para Aprovação/Consulta, o campo de pesquisa não apareceu. "
                "Screenshot/HTML salvos em 'diagnosticos'."
            )

        self.log(f"Tela de pesquisa carregada. URL atual: {self.driver.current_url}")

    def pesquisar_proposta(self, numero: str):
        try:
            self.driver.switch_to.default_content()
        except Exception:
            pass
        campo = self.wait.until(
            EC.visibility_of_element_located(
                (By.ID, "ctl00_Cph_AprCons_txtPesquisa_CAMPO")
            )
        )
        campo.clear()
        campo.send_keys(str(numero))

        botao_pesquisar = self.wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//*[@id='ctl00_Cph_AprCons_btnPesquisar_dvCBtn']//a",
                )
            )
        )
        self._click_el(botao_pesquisar)

    def _linha_do_resultado(self, numero: str):
        try:
            link_numero = self.wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        f"//a[contains(@href,'NrProposta') and normalize-space(text())='{numero}']",
                    )
                )
            )
        except TimeoutException:
            return None
        return link_numero.find_element(By.XPATH, "./ancestor::tr[1]")

    def abrir_detalhe(self, linha):
        link_situacao = linha.find_element(
            By.XPATH, ".//a[contains(@href,'Situacao')]"
        )
        previous_handles = self.driver.window_handles
        self._click_el(link_situacao)
        abriu_nova_janela = self._switch_to_new_window_if_any(
            previous_handles, timeout=2.5
        )
        if abriu_nova_janela:
            self.log(
                "O detalhe abriu em uma NOVA JANELA/ABA - foco trocado para ela."
            )

    def ler_log_observacoes(self) -> str:
        texto = self._exec_js_retry(
            JS_LER_OBSERVACOES, sucesso_check=lambda r: r is not None
        )
        if texto is None:
            self.salvar_diagnostico("campo_observacoes_nao_encontrado")
            raise RuntimeError(
                "Campo de observações (log) não encontrado. Screenshot e HTML salvos em 'diagnosticos'."
            )
        return (texto or "").strip()

    @staticmethod
    def extrair_motivo_cancelamento(log_texto: str):
        m = MOTIVO_REGEX.search(log_texto)
        if not m:
            return None
        resto = m.group(1).strip()
        motivos = re.findall(r"'([^']*)'", resto)
        if motivos:
            motivo_final = motivos[-1].strip()
            if len(motivos) > 1:
                lista_formatada = ", ".join(f"'{m.strip()}'" for m in motivos)
                return f"Motivo(s) do Cancelamento: {lista_formatada}"
            return f"Motivo(s) do Cancelamento: '{motivo_final}'"
        partes = [
            p.strip()
            for p in re.split(r",\s*(?=(?:[^']*'[^']*')*[^']*$)", resto)
            if p.strip()
        ]
        if partes:
            if len(partes) > 1:
                return f"Motivo(s) do Cancelamento: {', '.join(partes)}"
            return f"Motivo(s) do Cancelamento: {partes[0]}"
        aspas = QUOTED_REGEX.search(resto)
        motivo = (
            aspas.group(1).strip()
            if aspas
            else resto.splitlines()[0].strip().rstrip("'").strip()
        )
        return f"Motivo(s) do Cancelamento: '{motivo}'"

    def abrir_atividades_executadas(self):
        clicou = self._exec_js_retry(
            JS_CLICAR_POR_ID_OU_TEXTO,
            args=("ctl00_cph_lnkAtividades", ["Atividades Executadas"]),
            sucesso_check=lambda r: r is True,
        )
        if not clicou:
            self.salvar_diagnostico("link_atividades_nao_encontrado")
            raise RuntimeError(
                "Link 'Atividades Executadas' não encontrado. Screenshot e HTML salvos em 'diagnosticos'."
            )

    @staticmethod
    def _mapear_colunas_atividades(linhas: list):
        idx_desc_atv = None
        idx_usuario_inicial = None
        for linha in linhas[:3]:
            for i, celula in enumerate(linha):
                chave = _normalizar_coluna(celula)
                if chave == "DESCATV":
                    idx_desc_atv = i
                elif chave == "USUARIOINICIAL":
                    idx_usuario_inicial = i
                if idx_desc_atv is not None and idx_usuario_inicial is not None:
                    break
        return idx_desc_atv, idx_usuario_inicial

    @staticmethod
    def _extrair_usuario_a_direita(linha: list, idx_desc_atv=None) -> str:
        indices_cancelada = [
            i for i, c in enumerate(linha) if _normalizar(c) == "CANCELADA"
        ]
        if not indices_cancelada:
            indices_cancelada = [
                i for i, c in enumerate(linha) if "CANCELADA" in _normalizar(c)
            ]
        if not indices_cancelada:
            return ""
        if idx_desc_atv is not None and idx_desc_atv in indices_cancelada:
            idx_cancelada = idx_desc_atv
        else:
            idx_cancelada = indices_cancelada[-1]
        for celula in linha[idx_cancelada + 1 :]:
            candidato = _limpar_texto(celula)
            if _parece_usuario(candidato):
                return candidato
        return ""

    def usuario_do_cancelamento(self) -> str:
        def _tabela_correta(r):
            if r is None:
                return False
            return any(
                _normalizar(celula) == "CANCELADA" for linha in r for celula in linha
            )

        linhas = self._exec_js_retry(
            JS_EXTRAIR_TABELA_ATIVIDADES,
            timeout=8,
            sucesso_check=_tabela_correta,
        )

        if linhas is None:
            self.salvar_diagnostico("tabela_atividades_nao_encontrada")
            raise RuntimeError(
                "Tabela de Atividades Executadas não encontrada. Screenshot e HTML salvos em 'diagnosticos'."
            )

        idx_desc_atv, idx_usuario_inicial = self._mapear_colunas_atividades(linhas)

        linha_cancelada = None
        for linha in linhas:
            if idx_desc_atv is not None and idx_desc_atv < len(linha):
                if _normalizar(linha[idx_desc_atv]) == "CANCELADA":
                    linha_cancelada = linha
                    break

        if linha_cancelada is None:
            for linha in linhas:
                if any(_normalizar(celula) == "CANCELADA" for celula in linha):
                    linha_cancelada = linha
                    break

        if linha_cancelada is None:
            self.log(
                f"Nenhuma linha 'CANCELADA' encontrada. Linhas extraídas: {linhas}"
            )
            return ""

        if (
            idx_usuario_inicial is not None
            and idx_usuario_inicial < len(linha_cancelada)
        ):
            candidato = _limpar_texto(linha_cancelada[idx_usuario_inicial])
            if _parece_usuario(candidato):
                self.log(
                    f"Linha CANCELADA: {linha_cancelada} | Usuário (coluna 'Usuário Inicial'): '{candidato}'"
                )
                return candidato

        usuario = self._extrair_usuario_a_direita(linha_cancelada, idx_desc_atv)
        if usuario:
            self.log(
                f"Linha CANCELADA: {linha_cancelada} | Usuário (à direita de 'CANCELADA'): '{usuario}'"
            )
            return usuario

        candidatos_usuario = [c for c in linha_cancelada if _parece_usuario(c)]
        usuario_encontrado = candidatos_usuario[0] if candidatos_usuario else ""
        self.log(
            f"Linha CANCELADA: {linha_cancelada} | Usuário identificado (fallback por padrão): '{usuario_encontrado}'"
        )
        return usuario_encontrado

    def fechar_modais(self, tentativas=4):
        for _ in range(tentativas):
            fechou = self._exec_js(JS_FECHAR_MODAL)
            if not fechou:
                break
            time.sleep(0.2)
        try:
            self.driver.switch_to.default_content()
        except Exception:
            pass
        try:
            handles = self.driver.window_handles
            if len(handles) > 1:
                for h in handles[1:]:
                    try:
                        self.driver.switch_to.window(h)
                        self.driver.close()
                    except Exception:
                        pass
                self.driver.switch_to.window(handles[0])
        except Exception:
            pass

    def processar_proposta(self, numero: str):
        self.pesquisar_proposta(numero)
        linha = self._linha_do_resultado(numero)
        if linha is None:
            return "PROPOSTA NÃO ENCONTRADA", ""

        self.abrir_detalhe(linha)
        try:
            log_texto = self.ler_log_observacoes()
        except Exception:
            self.fechar_modais()
            raise

        motivo = self.extrair_motivo_cancelamento(log_texto)
        if motivo:
            self.fechar_modais()
            return motivo, log_texto

        try:
            self.abrir_atividades_executadas()
            usuario = self.usuario_do_cancelamento()
        except Exception:
            self.fechar_modais()
            raise

        coluna_b = (
            f"CANCELAMENTO - {usuario}"
            if usuario
            else "CANCELAMENTO - USUÁRIO NÃO IDENTIFICADO"
        )
        self.fechar_modais()
        return coluna_b, log_texto

    # ------------------------ EXECUÇÃO RESILIENTE ------------------------ #

    def _eh_erro_de_sessao(self, e: Exception) -> bool:
        msg = str(e).lower()
        return any(
            chave in msg
            for chave in (
                "chrome not reachable",
                "disconnected: not connected to devtools",
                "no such window",
                "session deleted",
                "invalid session id",
            )
        )

    def executar(self, propostas: list, incremental_callback=None):
        """Executa a automação para a lista de propostas.

        incremental_callback(numero, col_b, col_c, indice, total)
        é chamado a cada proposta.
        """
        resultados = []

        # primeira inicialização
        self.start_browser()
        self.login()
        self.navegar_para_esteira()

        total = len(propostas)
        for i, numero in enumerate(propostas, start=1):
            self.log(f"[{i}/{total}] Processando proposta {numero}...")

            tentativa = 0
            while True:
                tentativa += 1
                try:
                    col_b, col_c = self.processar_proposta(numero)
                    break
                except Exception as e:
                    if self._eh_erro_de_sessao(e) and self.restarts_feitos < MAX_BROWSER_RESTARTS:
                        self.log(
                            f" -> ERRO de sessão na proposta {numero}: {e}. Tentando reiniciar navegador e retomar."
                        )
                        log_para_arquivo(
                            "TRACEBACK SESSAO:\n" + traceback.format_exc()
                        )
                        try:
                            self._reiniciar_navegador()
                        except Exception as e_restart:
                            self.log(
                                f" -> FALHA AO REINICIAR NAVEGADOR: {e_restart}. Marcando proposta como ERRO."
                            )
                            log_para_arquivo(
                                "TRACEBACK RESTART:\n" + traceback.format_exc()
                            )
                            col_b, col_c = f"ERRO: {e}", ""
                            break
                        # após reinício, tenta a proposta de novo (loop while)
                        continue
                    else:
                        self.log(f" -> ERRO na proposta {numero}: {e}")
                        log_para_arquivo(
                            "TRACEBACK PROPOSTA:\n" + traceback.format_exc()
                        )
                        col_b, col_c = f"ERRO: {e}", ""
                        self.fechar_modais()
                        break

            resultados.append((numero, col_b, col_c))
            self.log(f" -> {col_b}")

            if incremental_callback is not None:
                try:
                    incremental_callback(numero, col_b, col_c, i, total)
                except Exception as e_cb:
                    self.log(f"ERRO no salvamento incremental: {e_cb}")

        self.quit()
        return resultados

    def testar_navegador(self):
        self.start_browser()
        self.driver.get(LOGIN_URL)
        self.log(
            "Teste concluído. Verifique se a página carregou corretamente na janela do Chrome."
        )


# --------------------------------------------------------------------------- #
# LEITURA / ESCRITA DE PLANILHA (com salvamento incremental)
# --------------------------------------------------------------------------- #


def carregar_propostas_de_arquivo(caminho: str):
    caminho = Path(caminho)
    propostas = []
    if caminho.suffix.lower() in (".xlsx", ".xlsm"):
        wb = load_workbook(caminho, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):
            valor = row[0]
            if valor is None or str(valor).strip() == "":
                continue
            propostas.append(str(valor).strip())
    else:
        import csv

        with open(caminho, newline="", encoding="utf-8-sig") as f:
            for row in csv.reader(f):
                if row and row[0].strip():
                    propostas.append(row[0].strip())
    return propostas


def preparar_workbook_saida(caminho_saida: str, planilha_original: str = None):
    caminho_saida = Path(caminho_saida)
    if planilha_original and Path(planilha_original).suffix.lower() in (".xlsx", ".xlsm"):
        wb = load_workbook(planilha_original)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value="Proposta")
        ws.cell(row=1, column=2, value="Resultado")
        ws.cell(row=1, column=3, value="Log Completo")
    return wb, ws, caminho_saida


def salvar_resultado_completo(resultados: list, wb, ws, caminho_saida: Path) -> str:
    linha_por_proposta = {}
    for row_idx in range(1, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=1).value
        if val is not None:
            linha_por_proposta[str(val).strip()] = row_idx

    for numero, col_b, col_c in resultados:
        chave = str(numero).strip()
        linha_idx = linha_por_proposta.get(chave)
        if linha_idx is None:
            linha_idx = ws.max_row + 1
        ws.cell(row=linha_idx, column=1, value=numero)
        ws.cell(row=linha_idx, column=2, value=col_b)
        ws.cell(row=linha_idx, column=3, value=col_c)
        ws.cell(row=linha_idx, column=3).alignment = ws.cell(
            row=linha_idx, column=3
        ).alignment.copy(wrap_text=True)

    try:
        wb.save(caminho_saida)
        return str(caminho_saida)
    except PermissionError:
        alternativo = caminho_saida.with_name(
            f"{caminho_saida.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{caminho_saida.suffix}"
        )
        try:
            wb.save(alternativo)
            return str(alternativo)
        except Exception as e2:
            raise RuntimeError(
                f"Não foi possível salvar em '{caminho_saida}' (permissão negada) nem em "
                f"'{alternativo}'. Feche o arquivo se estiver aberto. Erro: {e2}"
            ) from e2


def salvar_incremental(wb, ws, caminho_saida: Path):
    try:
        wb.save(caminho_saida)
    except PermissionError:
        log_para_arquivo(
            f"PERMISSION ERROR ao salvar incremental em '{caminho_saida}'. Arquivo pode estar aberto."
        )
    except Exception as e:
        log_para_arquivo(f"ERRO ao salvar incremental: {e}")


# --------------------------------------------------------------------------- #
# INTERFACE GRÁFICA – mesma da 2.0/2.1
# --------------------------------------------------------------------------- #


class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação Esteira - Consulta de Propostas")
        self.root.geometry("1155x640")
        self.root.minsize(1090, 610)

        self.BG = "#F4F5F9"
        self.SIDEBAR_BG = "#0F172A"
        self.CARD = "#FFFFFF"
        self.TEXT = "#111827"
        self.MUTED = "#9CA3AF"
        self.ACCENT = "#2563EB"
        self.ACCENT_HOVER = "#1D4ED8"
        self.LOG_BG = "#020617"
        self.LOG_FG = "#E5E7EB"

        self.root.configure(bg=self.BG)

        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Root.TFrame", background=self.BG)
        style.configure("Sidebar.TFrame", background=self.SIDEBAR_BG)
        style.configure("Card.TFrame", background=self.CARD)
        style.configure(
            "Card.TLabelframe", background=self.CARD, foreground=self.TEXT
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=self.CARD,
            foreground=self.TEXT,
            font=("Segoe UI", 9, "bold"),
        )
        style.configure(
            "Step.TLabel",
            background=self.SIDEBAR_BG,
            foreground=self.MUTED,
            font=("Segoe UI", 9),
        )
        style.configure(
            "StepActive.TLabel",
            background=self.SIDEBAR_BG,
            foreground="#E5E7EB",
            font=("Segoe UI", 9, "bold"),
        )
        style.configure(
            "StepTitle.TLabel",
            background=self.SIDEBAR_BG,
            foreground="#E5E7EB",
            font=("Segoe UI", 16, "bold"),
        )
        style.configure(
            "TLabel", background=self.CARD, foreground=self.TEXT, font=("Segoe UI", 9)
        )
        style.configure(
            "Muted.TLabel",
            background=self.CARD,
            foreground=self.MUTED,
            font=("Segoe UI", 8),
        )
        style.configure(
            "Title.TLabel",
            background=self.BG,
            foreground=self.TEXT,
            font=("Segoe UI", 17, "bold"),
        )
        style.configure(
            "Sub.TLabel",
            background=self.BG,
            foreground=self.MUTED,
            font=("Segoe UI", 9),
        )
        style.configure(
            "TEntry",
            fieldbackground="#FFFFFF",
            background="#FFFFFF",
            foreground=self.TEXT,
            padding=4,
        )
        style.configure(
            "TButton", font=("Segoe UI", 9, "bold"), padding=(9, 4)
        )
        style.configure(
            "Primary.TButton", background=self.ACCENT, foreground="white"
        )
        style.map(
            "Primary.TButton",
            background=[("active", self.ACCENT_HOVER), ("disabled", "#93C5FD")],
        )
        style.configure(
            "Ghost.TButton", background="#E5E7EB", foreground=self.TEXT
        )
        style.map(
            "Ghost.TButton", background=[("active", "#D1D5DB")]
        )

        self.arquivo_entrada = tk.StringVar()
        self.chromedriver_var = tk.StringVar()
        self.chrome_binary_var = tk.StringVar()
        self.saida_var = tk.StringVar(
            value=str(SCRIPT_DIR / "resultado_propostas.xlsx")
        )
        self.status_deteccao = tk.StringVar(
            value="Buscando chrome.exe e chromedriver.exe automaticamente..."
        )
        self.contagem_var = tk.StringVar(value="Nenhuma proposta carregada")

        root_frame = ttk.Frame(self.root, style="Root.TFrame")
        root_frame.pack(fill="both", expand=True)
        root_frame.grid_columnconfigure(0, weight=0)
        root_frame.grid_columnconfigure(1, weight=1)
        root_frame.grid_rowconfigure(0, weight=1)

        sidebar = ttk.Frame(root_frame, style="Sidebar.TFrame")
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_rowconfigure(11, weight=1)

        ttk.Label(
            sidebar,
            text="Consulta de Propostas",
            style="StepTitle.TLabel",
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(16, 2))
        ttk.Label(
            sidebar,
            text="Fluxo guiado de consulta",
            style="Step.TLabel",
        ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 12))

        for idx, (num, texto) in enumerate(
            [
                ("1", "Acesso ao sistema"),
                ("2", "Navegador"),
                ("3", "Propostas"),
                ("4", "Arquivo de saída"),
                ("5", "Execução & Log"),
            ],
            start=2,
        ):
            ttk.Label(
                sidebar,
                text=f"{num} {texto}",
                style="StepActive.TLabel" if num == "1" else "Step.TLabel",
            ).grid(row=idx, column=0, sticky="w", padx=16, pady=3)

        ttk.Label(
            sidebar,
            text="Projetado por Pablo Di Sisto",
            background=self.SIDEBAR_BG,
            foreground="#CBD5E1",
            font=("Segoe UI Semibold", 10),
        ).grid(row=12, column=0, sticky="w", padx=16, pady=(8, 14))

        main = ttk.Frame(root_frame, style="Root.TFrame")
        main.grid(row=0, column=1, sticky="nsew", padx=(10, 14), pady=(10, 10))
        main.grid_columnconfigure(0, weight=1)
        main.grid_columnconfigure(1, weight=1)
        main.grid_rowconfigure(1, weight=1)

        header = ttk.Frame(main, style="Root.TFrame")
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(
            header,
            text="Painel de consulta de propostas",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Preencha os dados abaixo e acompanhe a execução em tempo real.",
            style="Sub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(1, 3))

        left = ttk.Frame(main, style="Root.TFrame")
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 5))
        left.grid_columnconfigure(0, weight=1)

        right = ttk.Frame(main, style="Root.TFrame")
        right.grid(row=1, column=1, sticky="nsew", padx=(5, 0))
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(0, weight=1)

        frame_login = ttk.Labelframe(
            left, text="1. Acesso ao sistema", style="Card.TLabelframe"
        )
        frame_login.pack(fill="x", pady=(0, 5))
        frame_login.grid_columnconfigure(1, weight=1)
        ttk.Label(frame_login, text="Usuário:").grid(
            row=0, column=0, sticky="w", padx=9, pady=(5, 2)
        )
        self.entry_usuario = ttk.Entry(frame_login)
        self.entry_usuario.grid(
            row=0, column=1, sticky="ew", padx=9, pady=(5, 2)
        )
        ttk.Label(frame_login, text="Senha:").grid(
            row=1, column=0, sticky="w", padx=9, pady=(0, 5)
        )
        self.entry_senha = ttk.Entry(frame_login, show="*")
        self.entry_senha.grid(
            row=1, column=1, sticky="ew", padx=9, pady=(0, 5)
        )

        frame_chrome = ttk.Labelframe(
            left, text="2. Navegador", style="Card.TLabelframe"
        )
        frame_chrome.pack(fill="x", pady=(0, 5))
        frame_chrome.grid_columnconfigure(1, weight=1)
        ttk.Label(frame_chrome, text="Caminho do chrome.exe:").grid(
            row=0, column=0, sticky="w", padx=9, pady=(5, 2)
        )
        ttk.Entry(
            frame_chrome, textvariable=self.chrome_binary_var
        ).grid(row=0, column=1, sticky="ew", padx=9, pady=(5, 2))
        ttk.Button(
            frame_chrome,
            text="Selecionar...",
            style="Ghost.TButton",
            command=self.selecionar_chrome_binary,
        ).grid(row=0, column=2, padx=(0, 9), pady=(5, 2))

        ttk.Label(frame_chrome, text="Caminho do chromedriver.exe:").grid(
            row=1, column=0, sticky="w", padx=9, pady=2
        )
        ttk.Entry(
            frame_chrome, textvariable=self.chromedriver_var
        ).grid(row=1, column=1, sticky="ew", padx=9, pady=2)
        ttk.Button(
            frame_chrome,
            text="Selecionar...",
            style="Ghost.TButton",
            command=self.selecionar_chromedriver,
        ).grid(row=1, column=2, padx=(0, 9), pady=2)

        ttk.Label(
            frame_chrome,
            textvariable=self.status_deteccao,
            style="Muted.TLabel",
        ).grid(row=2, column=0, columnspan=3, sticky="w", padx=9, pady=(1, 2))

        chrome_btns = ttk.Frame(frame_chrome, style="Card.TFrame")
        chrome_btns.grid(row=3, column=0, columnspan=3, sticky="w", padx=9, pady=(0, 5))
        ttk.Button(
            chrome_btns,
            text="Buscar novamente",
            style="Ghost.TButton",
            command=self.buscar_automaticamente,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            chrome_btns,
            text="Testar Chrome",
            style="Ghost.TButton",
            command=self.testar_chrome,
        ).pack(side="left")

        frame_props = ttk.Labelframe(
            left, text="3. Propostas", style="Card.TLabelframe"
        )
        frame_props.pack(fill="both", expand=True, pady=(0, 5))
        frame_props.grid_columnconfigure(0, weight=1)
        frame_props.grid_rowconfigure(1, weight=1)
        ttk.Label(
            frame_props,
            text="Cole os números de proposta (um por linha):",
        ).grid(row=0, column=0, sticky="w", padx=9, pady=(5, 2))
        self.text_propostas = scrolledtext.ScrolledText(
            frame_props,
            height=6,
            bg="#FFFFFF",
            fg=self.TEXT,
            insertbackground="black",
            font=("Consolas", 10),
        )
        self.text_propostas.grid(row=1, column=0, sticky="nsew", padx=9)

        prop_file = ttk.Frame(frame_props)
        prop_file.grid(row=2, column=0, sticky="ew", padx=9, pady=(5, 2))
        prop_file.grid_columnconfigure(1, weight=1)
        ttk.Button(
            prop_file,
            text="Selecionar planilha (.xlsx/.csv)",
            style="Ghost.TButton",
            command=self.selecionar_arquivo,
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            prop_file,
            textvariable=self.arquivo_entrada,
            style="Muted.TLabel",
        ).grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(
            frame_props,
            text=(
                "Se uma planilha for selecionada, os números serão lidos da coluna A."
            ),
            style="Muted.TLabel",
        ).grid(row=3, column=0, sticky="w", padx=9, pady=(0, 5))

        frame_saida = ttk.Labelframe(
            left, text="4. Arquivo de saída", style="Card.TLabelframe"
        )
        frame_saida.pack(fill="x", pady=(0, 5))
        frame_saida.grid_columnconfigure(1, weight=1)
        ttk.Label(frame_saida, text="Salvar resultado em:").grid(
            row=0, column=0, sticky="w", padx=9, pady=(5, 2)
        )
        ttk.Entry(frame_saida, textvariable=self.saida_var).grid(
            row=0, column=1, sticky="ew", padx=9, pady=(5, 2)
        )
        ttk.Button(
            frame_saida,
            text="Escolher...",
            style="Ghost.TButton",
            command=self.escolher_saida,
        ).grid(row=0, column=2, padx=(0, 9), pady=(5, 2))

        action_row = ttk.Frame(left, style="Root.TFrame")
        action_row.pack(fill="x")
        self.botao_iniciar = ttk.Button(
            action_row,
            text="Iniciar automação",
            style="Primary.TButton",
            command=self.iniciar,
        )
        self.botao_iniciar.pack(side="right")
        ttk.Label(
            action_row,
            textvariable=self.contagem_var,
            background=self.BG,
            foreground=self.MUTED,
            font=("Segoe UI", 8),
        ).pack(side="left")

        frame_log = ttk.Labelframe(
            right, text="5. Execução & Log", style="Card.TLabelframe"
        )
        frame_log.grid(row=0, column=0, sticky="nsew")
        frame_log.grid_columnconfigure(0, weight=1)
        frame_log.grid_rowconfigure(0, weight=1)
        self.text_log = scrolledtext.ScrolledText(
            frame_log,
            height=10,
            bg=self.LOG_BG,
            fg=self.LOG_FG,
            insertbackground="white",
            font=("Consolas", 10),
        )
        self.text_log.grid(row=0, column=0, sticky="nsew", padx=9, pady=(5, 3))
        ttk.Label(
            frame_log,
            text="Projetado por Pablo Di Sisto",
            background=self.CARD,
            foreground=self.ACCENT,
            font=("Segoe UI Semibold", 10),
        ).grid(row=1, column=0, sticky="w", padx=9, pady=(0, 5))

        threading.Thread(target=self.buscar_automaticamente, daemon=True).start()

    # ------------------------ Métodos da interface ------------------------ #

    def buscar_automaticamente(self):
        config = carregar_config()
        chrome_path, driver_path = auto_detectar_caminhos(
            config.get("chrome_binary_path", ""),
            config.get("chromedriver_path", ""),
        )

        def _atualizar():
            if chrome_path:
                self.chrome_binary_var.set(chrome_path)
            if driver_path:
                self.chromedriver_var.set(driver_path)
            if chrome_path and driver_path:
                self.status_deteccao.set(
                    f"Detectado: {Path(chrome_path).name} + {Path(driver_path).name}"
                )
            elif chrome_path or driver_path:
                self.status_deteccao.set(
                    "Encontrado parcialmente - confira/complete os caminhos abaixo."
                )
            else:
                self.status_deteccao.set(
                    "Não encontrado automaticamente. Selecione manualmente."
                )

        self.root.after(0, _atualizar)

    def selecionar_chromedriver(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o chromedriver.exe",
            filetypes=[("Executável", "*.exe"), ("Todos", "*.*")],
        )
        if caminho:
            self.chromedriver_var.set(caminho)

    def selecionar_chrome_binary(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o chrome.exe",
            filetypes=[("Executável", "*.exe"), ("Todos", "*.*")],
        )
        if caminho:
            self.chrome_binary_var.set(caminho)

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha de propostas",
            filetypes=[("Planilhas", "*.xlsx *.xlsm *.csv"), ("Todos", "*.*")],
        )
        if caminho:
            self.arquivo_entrada.set(caminho)

    def escolher_saida(self):
        caminho = filedialog.asksaveasfilename(
            title="Salvar resultado como",
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")],
        )
        if caminho:
            self.saida_var.set(caminho)

    def log(self, msg: str):
        self.text_log.configure(state="normal")
        self.text_log.insert("end", msg + "\n")
        self.text_log.see("end")
        self.text_log.configure(state="disabled")

    def coletar_propostas(self):
        propostas = []
        arquivo = self.arquivo_entrada.get().strip()
        if arquivo:
            propostas = carregar_propostas_de_arquivo(arquivo)
        else:
            bruto = self.text_propostas.get("1.0", "end").strip()
            for linha in bruto.splitlines():
                linha = linha.strip().strip(",")
                if linha:
                    propostas.append(linha)
        vistos = set()
        unicos = []
        for p in propostas:
            if p not in vistos:
                vistos.add(p)
                unicos.append(p)
        self.contagem_var.set(
            f"{len(unicos)} proposta(s) pronta(s) para processar"
        )
        return unicos

    def testar_chrome(self):
        self.log("Testando abertura do Chrome...")
        bot = EsteiraBot(
            "teste",
            "teste",
            log_callback=lambda m: self.root.after(0, self.log, m),
            chromedriver_path=self.chromedriver_var.get(),
            chrome_binary_path=self.chrome_binary_var.get(),
        )

        def _run():
            try:
                bot.testar_navegador()
            except Exception as e:
                self.root.after(0, self.log, f"ERRO NO TESTE:\n{e}")
                log_para_arquivo(
                    "TRACEBACK COMPLETO:\n" + traceback.format_exc()
                )
                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        "Erro ao abrir o Chrome", "Veja o log para detalhes."
                    ),
                )
            finally:
                bot.quit()

        threading.Thread(target=_run, daemon=True).start()

    def iniciar(self):
        usuario = self.entry_usuario.get().strip()
        senha = self.entry_senha.get().strip()
        if not usuario or not senha:
            messagebox.showwarning("Atenção", "Informe usuário e senha.")
            return

        propostas = self.coletar_propostas()
        if not propostas:
            messagebox.showwarning("Atenção", "Informe ao menos uma proposta.")
            return

        saida = self.saida_var.get().strip()
        if not saida:
            messagebox.showwarning("Atenção", "Informe o arquivo de saída.")
            return

        self.botao_iniciar.configure(state="disabled")
        self.log(
            f"Iniciando automação para {len(propostas)} proposta(s)..."
        )

        thread = threading.Thread(
            target=self._executar_em_thread,
            args=(
                usuario,
                senha,
                propostas,
                saida,
                self.arquivo_entrada.get().strip() or None,
            ),
            daemon=True,
        )
        thread.start()

    def _executar_em_thread(
        self, usuario, senha, propostas, saida, arquivo_original
    ):
        bot = EsteiraBot(
            usuario,
            senha,
            log_callback=lambda m: self.root.after(0, self.log, m),
            chromedriver_path=self.chromedriver_var.get(),
            chrome_binary_path=self.chrome_binary_var.get(),
        )

        wb, ws, caminho_saida = preparar_workbook_saida(
            saida, planilha_original=arquivo_original
        )

        resultados = []

        def _incremental(numero, col_b, col_c, indice, total):
            chave = str(numero).strip()
            linha_existente = None
            for row_idx in range(1, ws.max_row + 1):
                val = ws.cell(row=row_idx, column=1).value
                if val is not None and str(val).strip() == chave:
                    linha_existente = row_idx
                    break
            if linha_existente is None:
                linha_existente = ws.max_row + 1

            ws.cell(row=linha_existente, column=1, value=numero)
            ws.cell(row=linha_existente, column=2, value=col_b)
            ws.cell(row=linha_existente, column=3, value=col_c)
            ws.cell(row=linha_existente, column=3).alignment = ws.cell(
                row=linha_existente, column=3
            ).alignment.copy(wrap_text=True)

            resultados.append((numero, col_b, col_c))

            # salvamento incremental a cada 10 propostas (ajustável)
            if indice % 10 == 0 or indice == total:
                salvar_incremental(wb, ws, caminho_saida)

        try:
            bot.executar(propostas, incremental_callback=_incremental)
            caminho_final = salvar_resultado_completo(
                resultados, wb, ws, caminho_saida
            )
            self.root.after(
                0,
                self.log,
                f"Concluído! Resultado salvo em: {caminho_final}",
            )
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Sucesso",
                    f"Automação concluída.\nArquivo:\n{caminho_final}",
                ),
            )
        except Exception as e:
            self.root.after(0, self.log, f"ERRO FATAL:\n{e}")
            log_para_arquivo(
                "TRACEBACK COMPLETO:\n" + traceback.format_exc()
            )
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "Erro", "A automação falhou. Veja o log."
                ),
            )
        finally:
            self.root.after(
                0, lambda: self.botao_iniciar.configure(state="normal")
            )


def main():
    root = tk.Tk()
    app = AppGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
