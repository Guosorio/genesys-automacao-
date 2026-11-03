import csv
import io
import os
import re
import shutil
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

BASE = Path(__file__).resolve().parent
load_dotenv(BASE / ".env", override=True)

URL = "https://login.mypurecloud.com/?rid=cvPX_YK-DPxcWAI3GC2usMRu5fCI_1EOCb9L0q-qDlI#/authenticate-adv/org/americanet"
LOGIN = (os.getenv("GENESYS_LOGIN") or "").strip()
SENHA = (os.getenv("GENESYS_SENHA") or "").strip()
HEADLESS = os.getenv("HEADLESS", "0") == "1"

PASTA_DOWNLOAD = BASE / "downloads"
# Perfil em pasta LOCAL: o Chromium trava se o perfil ficar em pasta de rede (\\servidor\...)
PASTA_SESSAO = Path(os.getenv("LOCALAPPDATA") or Path.home()) / "genesys_sessao_browser"
PASTA_DOWNLOAD.mkdir(exist_ok=True)

TIMEOUT = 30_000  # ms

# ───────────────────────── Excel de destino (aba Base) ─────────────────────────
# Caminho completo do Excel que recebe os dados. Pode ser sobrescrito no .env (EXCEL_PAUSAS=...).
# Troque o nome do arquivo quando virar o mês.
EXCEL_PAUSAS = Path(os.getenv("EXCEL_PAUSAS") or (
    r"\\172.22.0.101\Planejamento2\10. Gustavo\Automação Genesys\downloads\Relátorio Genesys"
    r"\Pausas Vero - Outubro.xlsx"
))
ABA_BASE = os.getenv("ABA_BASE", "Base")
PASTA_BACKUP = BASE / "backup_excel"   # cópia de segurança do Excel antes de cada atualização
MANTER_BACKUPS = 5
LINHA_LIMITE_DASH = 15_000             # a aba Dash só lê a Base até esta linha

# O CSV do Genesys tem 18 colunas e entra nas colunas B:S da Base (a coluna A é a fórmula =INT(B)).
# A ORDEM importa: a Dash lê "Atendimento Chat" na coluna S. Se o Genesys mudar o layout, o script para.
CABECALHO_CSV = [
    "Início do intervalo", "Fim do intervalo", "Intervalo concluído", "Filtros", "ID de agente",
    "Nome do agente", "Qualificações", "Conectado", "Na fila", "Ocioso", "Disponível", "Ausente",
    "Intervalo", "Refeição", "Não respondendo", "Fora da fila", "Ocupado", "Ocupado: Atendimento Chat",
]

# Layout de exportação (tela "Exportar" > Formato)
LAYOUT = {
    "Formato de duração": "HH:MM:SS",
    "Delimitador": "Ponto e vírgula",
}
INCLUIR_LINHA_RESUMO = False
EXCLUIR_LINHAS_VAZIAS = True

# Intervalo entre o clique em Atualizar e o início da exportação
INTERVALO_APOS_ATUALIZAR_S = 5

# Antes desta hora (horário local) o script extrai o dia ANTERIOR; a partir dela, o dia de HOJE.
# Ex.: rodando às 8h -> ontem; rodando às 10h ou depois -> hoje.
HORA_CORTE = int(os.getenv("HORA_CORTE", "10"))

MESES = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5,
    "junho": 6, "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10,
    "novembro": 11, "dezembro": 12,
}
# Data exibida no topo central do Espaço de trabalho de análises: "27 de outubro de 2026"
PADRAO_DATA = re.compile(r"\b(\d{1,2})\s+de\s+([a-zçã]+)\s+de\s+(\d{4})\b", re.I)

PADRAO_PULAR = re.compile(r"(ignorar|pular)\s+(a\s+)?configura|skip\s+setup", re.I)
PADRAO_COLLABORATE = re.compile(r"Collaborate\s*/\s*Communicate", re.I)

XPATH_ALTERNAR_PAINEL = "/html/body/div/div[2]/div/div/div/div/div[1]/div[3]/span[4]/button"
XPATH_ATUALIZAR = "//*[starts-with(@id,'ember')]/div/div[1]/div[3]/span[1]/button"
XPATH_FORMATO = (
    "/html/body/div/div[2]/div/div/div/div/div[2]/div[2]/aside/div[2]/"
    "div/div[5]/gux-accordion/gux-accordion-section/h2"
)
XPATH_DOWNLOAD = (
    "/html/body/div[2]/div[1]/div[2]/aside/div/div[2]/section/div/"
    "div[2]/div/div[10]/div[2]/div[1]/div/div[4]/div[1]/button"
)


# ───────────────────────── utilidades ─────────────────────────

def ve(loc):
    """True se o elemento existe e está visível AGORA (sem esperar)."""
    try:
        return loc.count() > 0 and loc.first.is_visible()
    except Exception:
        return False


def tirar_print(page, arquivo):
    try:
        page.screenshot(path=str(BASE / arquivo), full_page=True)
    except Exception as e:
        print(f"(não consegui salvar {arquivo}: {type(e).__name__})")


def falhar(page, arquivo, msg):
    tirar_print(page, arquivo)
    raise RuntimeError(f"{msg} Veja {arquivo}.")


def esperar(page, *criadores, timeout_s=30, habilitado=False):
    """
    Procura em TODOS os frames o primeiro elemento visível, testando os criadores
    (funções frame -> locator) em ordem de prioridade.
    Retorna (frame, elemento) ou (None, None) se estourar o tempo.
    """
    fim = time.time() + timeout_s
    while time.time() < fim:
        for frame in page.frames:
            for criar in criadores:
                try:
                    loc = criar(frame)
                    if ve(loc) and (not habilitado or loc.first.is_enabled()):
                        return frame, loc.first
                except Exception:
                    pass
        page.wait_for_timeout(400)
    return None, None


def fechar_notificacoes(page, max_s=6):
    """Fecha os avisos flutuantes do Genesys; sai cedo quando não há mais nenhum."""
    seletor = (
        'button.messenger-close, [class*="messenger-close"], '
        'button[aria-label*="Fechar notificação" i], '
        'button[aria-label*="Fechar notificacao" i], '
        'button[aria-label*="close notification" i]'
    )
    fim = time.time() + max_s
    fechadas, ultimo = 0, time.time()
    while time.time() < fim:
        achou = False
        for frame in page.frames:
            try:
                botoes = frame.locator(seletor)
                for i in range(botoes.count()):
                    b = botoes.nth(i)
                    if b.is_visible() and b.is_enabled():
                        b.click(timeout=2_000)
                        fechadas += 1
                        achou = True
            except Exception:
                pass
        if achou:
            ultimo = time.time()
        elif fechadas and time.time() - ultimo > 1.5:
            break
        page.wait_for_timeout(400)
    print(f"Notificações fechadas: {fechadas}.")


# ───────────────────────── login ─────────────────────────

def no_app(page):
    """True quando já está no app do Genesys (apps.<região>.com)."""
    return (urlparse(page.url).hostname or "").startswith("apps.")


def clicar_logo_microsoft(page):
    estrategias = [
        lambda: page.get_by_role("link", name="Microsoft"),
        lambda: page.get_by_role("button", name="Microsoft"),
        lambda: page.get_by_alt_text("Microsoft"),
        lambda: page.get_by_text("Microsoft", exact=True),
        lambda: page.locator("img[src*='microsoft' i], [class*='microsoft' i], [id*='microsoft' i]"),
        lambda: page.locator("img[src*='azure' i], [class*='azure' i], [id*='azure' i]"),
    ]
    fim = time.time() + 5
    while time.time() < fim:
        for criar in estrategias:
            try:
                alvo = criar()
                if ve(alvo):
                    alvo.first.click()
                    return
            except Exception:
                pass
        page.wait_for_timeout(300)

    # Fallback: o logo fica centralizado ~40px abaixo do título "Faça logon com sua conta"
    caixa = page.get_by_text("Faça logon com sua conta").first.bounding_box()
    if not caixa:
        falhar(page, "erro_login.png", "Não consegui clicar no logo da Microsoft.")
    page.mouse.click(caixa["x"] + caixa["width"] / 2, caixa["y"] + caixa["height"] + 40)


def digitar_senha(page, campo):
    """Digita a senha e CONFERE o campo antes de enviar (a tela da Microsoft pode apagá-lo)."""
    page.wait_for_timeout(800)  # animação da tela
    for _ in range(3):
        campo.click()
        campo.fill("")
        campo.press_sequentially(SENHA, delay=40)
        if campo.input_value() == SENHA:
            return
        page.wait_for_timeout(500)
    falhar(page, "erro_login.png", "Não consegui preencher a senha (nada foi enviado).")


def login_microsoft(page, clicar_logo=True, limite_s=90):
    """Trata cada tela do login pelo que está visível (funciona com ou sem sessão salva)."""
    if clicar_logo:
        clicar_logo_microsoft(page)

    senha_enviada = False
    fim = time.time() + limite_s
    while time.time() < fim:
        if no_app(page):
            return

        # O input de senha "fantasma" da tela de e-mail (moveOffScreen) precisa ser ignorado
        campo_senha = page.locator(
            "input[type='password']:not(.moveOffScreen):not([aria-hidden='true']), "
            "input[name='passwd']:not(.moveOffScreen):not([aria-hidden='true'])"
        )
        campo_email = page.locator("input[name='loginfmt'], input[type='email']")

        if ve(page.get_by_text(PADRAO_COLLABORATE)):
            print("Tela de aplicativos -> Collaborate/Communicate")
            page.get_by_text(PADRAO_COLLABORATE).first.click()
            page.wait_for_timeout(2500)

        elif ve(page.get_by_text("Escolha uma conta")):
            conta = page.get_by_text(LOGIN)
            (conta if ve(conta) else page.get_by_text("Use outra conta")).first.click()
            page.wait_for_timeout(800)

        elif ve(campo_senha) or ve(page.get_by_text("Insira a senha")):
            if senha_enviada:  # a tela voltou: senha recusada. Para antes de bloquear a conta
                falhar(page, "erro_login.png", "A Microsoft recusou a senha.")
            senha_enviada = True
            campo_senha.first.wait_for(state="visible", timeout=10_000)
            digitar_senha(page, campo_senha.first)
            page.keyboard.press("Enter")
            page.wait_for_timeout(2000)

        elif ve(campo_email):
            campo_email.first.fill(LOGIN)
            page.keyboard.press("Enter")
            page.wait_for_timeout(1000)

        elif ve(page.get_by_text("Continuar conectado")) or ve(page.get_by_text("Manter-me conectado")):
            page.locator("#idSIButton9, input[value='Sim']").first.click()
            page.wait_for_timeout(1000)

        elif ve(page.get_by_text(PADRAO_PULAR)):
            print("Tela de segurança -> pulando configuração")
            page.get_by_text(PADRAO_PULAR).first.click()
            page.wait_for_timeout(1000)

        elif ve(page.get_by_text("manter sua conta segura")) and not ve(page.get_by_text("número de telefone")):
            page.get_by_role("button", name="Avançar").first.click()
            page.wait_for_timeout(1000)

        else:
            page.wait_for_timeout(300)

    falhar(page, "erro_login.png", f"Login não concluiu em {limite_s}s (URL: {page.url}).")


def entrar(page):
    """Abre o site e conclui o login (ou só a escolha do app, se a sessão já estiver ativa)."""
    page.goto(URL, wait_until="domcontentloaded")
    fim = time.time() + 25
    while time.time() < fim:
        if no_app(page):
            return
        if ve(page.get_by_text("Faça logon com sua conta")):
            return login_microsoft(page)
        if ve(page.get_by_text(PADRAO_COLLABORATE)):
            return login_microsoft(page, clicar_logo=False)
        page.wait_for_timeout(300)
    falhar(page, "erro_login.png", "Nenhuma tela de login/aplicativos apareceu.")


# ───────────────────────── layout de exportação ─────────────────────────

SEL_DROPDOWN = "select, [role='combobox'], button[aria-haspopup], [aria-haspopup='listbox'], gux-dropdown"
SEL_CHECKBOX = "input[type='checkbox'], [role='checkbox']"


def achar_controle(painel, rotulo, seletores, exato):
    """Acha o controle pelo label; se não estiver associado, procura no mesmo bloco do texto."""
    campo = painel.get_by_label(rotulo, exact=False).first
    try:
        campo.wait_for(state="visible", timeout=1500)
        return campo
    except PWTimeout:
        pass
    label = painel.get_by_text(rotulo, exact=exato).first
    label.wait_for(state="visible", timeout=10_000)
    for nivel in range(1, 6):
        controles = label.locator("xpath=" + "/.." * nivel).locator(seletores)
        if ve(controles):
            return controles.first
    raise RuntimeError(f"Não encontrei o controle de '{rotulo}'.")


def escolher_opcao(frame, painel, rotulo, valor):
    try:
        achar_controle(painel, rotulo, SEL_DROPDOWN, exato=True).click()
        opcao = frame.get_by_role("option", name=valor, exact=False).first
        opcao.wait_for(state="visible", timeout=10_000)
        opcao.click()
        print(f"Formato: {rotulo} = {valor}")
    except Exception as e:
        falhar(frame.page, "erro_formato.png", f"Falha ao definir '{rotulo}' = '{valor}' ({type(e).__name__}).")


def ajustar_checkbox(painel, rotulo, marcado):
    caixa = achar_controle(painel, rotulo, SEL_CHECKBOX, exato=False)
    if caixa.is_checked() != marcado:
        caixa.set_checked(marcado)
    print(f"Checkbox: {rotulo} = {marcado}")


def expandir_secao_formato(frame, painel):
    campo = painel.get_by_text("Formato de duração").first
    if ve(campo):
        return
    cabecalhos = [
        frame.locator(f"xpath={XPATH_FORMATO}"),
        painel.get_by_text("Formato", exact=True),
        frame.get_by_text("Formato", exact=True),
    ]
    for _ in range(3):
        for cab in cabecalhos:
            try:
                if ve(cab):
                    cab.first.scroll_into_view_if_needed()
                    cab.first.click(timeout=5_000)
                    campo.wait_for(state="visible", timeout=8_000)
                    return
            except Exception:
                pass
        frame.page.wait_for_timeout(750)
    falhar(frame.page, "erro_formato.png", "Não consegui expandir a seção Formato.")


def configurar_layout(frame, painel):
    for rotulo, valor in LAYOUT.items():
        escolher_opcao(frame, painel, rotulo, valor)
    ajustar_checkbox(painel, "Incluir linha de resumo", INCLUIR_LINHA_RESUMO)
    ajustar_checkbox(painel, "Excluir linhas vazias", EXCLUIR_LINHAS_VAZIAS)


# ───────────────────────── exportação ─────────────────────────

def abrir_workspace_analises(page):
    menu = page.locator(
        "[aria-label*='menu' i], [title*='menu' i], [data-testid*='menu' i], button:has-text('Menu')"
    ).first
    try:
        menu.wait_for(state="visible", timeout=90_000)
    except PWTimeout:
        print("URL atual:", page.url)
        falhar(page, "erro_menu.png", "O botão Menu não apareceu.")
    menu.click()
    page.get_by_text("Análises").first.click()
    page.get_by_text("Espaço de trabalho de análises").first.click()
    page.wait_for_timeout(1000)
    fechar_notificacoes(page, max_s=10)


def ler_data_painel(page, timeout_s=30):
    """Lê a data exibida no topo central do painel. Retorna (elemento, date) ou (None, None)."""
    _, el = esperar(page, lambda f: f.get_by_text(PADRAO_DATA), timeout_s=timeout_s)
    if not el:
        return None, None
    try:
        m = PADRAO_DATA.search(el.inner_text())
        dia, mes, ano = int(m.group(1)), MESES[m.group(2).lower()], int(m.group(3))
        return el, datetime(ano, mes, dia).date()
    except Exception:
        return None, None


def data_alvo():
    """Dia a extrair: ontem antes de HORA_CORTE, hoje a partir de HORA_CORTE."""
    agora = datetime.now()
    return agora.date() if agora.hour >= HORA_CORTE else agora.date() - timedelta(days=1)


def clicar_seta(page, el_data, data_antes, direita=True):
    """Clica na seta ao lado da data ('>' avança, '<' volta) e confirma que a data mudou."""
    if direita:
        candidatos = [
            lambda: el_data.locator("xpath=following::button[1]"),  # 1º botão depois da data
            lambda: el_data.locator("xpath=ancestor::*[position()<=3]").locator(
                "button[aria-label*='próxim' i], button[aria-label*='proxim' i], "
                "button[aria-label*='next' i], button[title*='próxim' i], button[title*='next' i]"
            ),
        ]
    else:
        candidatos = [
            lambda: el_data.locator("xpath=preceding::button[1]"),  # último botão antes da data
            lambda: el_data.locator("xpath=ancestor::*[position()<=3]").locator(
                "button[aria-label*='anterior' i], button[aria-label*='previous' i], "
                "button[aria-label*='prev' i], button[title*='anterior' i], button[title*='previous' i]"
            ),
        ]
    for criar in candidatos:
        try:
            seta = criar().first
            seta.click(timeout=5_000)
        except Exception:
            continue
        fim = time.time() + 6
        while time.time() < fim:
            _, nova = ler_data_painel(page, timeout_s=1)
            if nova and nova != data_antes:
                return True
            page.wait_for_timeout(400)
    return False


def ajustar_data_painel(page):
    """Leva o painel até o dia alvo (ontem antes das HORA_CORTE h; hoje depois), usando as setas."""
    alvo = data_alvo()
    el, atual = ler_data_painel(page)
    if not atual:
        falhar(page, "erro_data.png", "Não consegui ler a data do painel de análises.")
    print(
        f"Agora: {datetime.now():%H:%M} (corte {HORA_CORTE}h) -> extraindo {alvo:%d/%m/%Y} | "
        f"data no Genesys: {atual:%d/%m/%Y}"
    )

    passos = 0
    while atual != alvo:
        passos += 1
        if passos > 62:
            falhar(page, "erro_data.png", "Não consegui chegar na data alvo em 62 cliques.")
        direita = atual < alvo
        if not clicar_seta(page, el, atual, direita=direita):
            lado = "direita" if direita else "esquerda"
            falhar(page, "erro_data.png", f"Cliquei na seta {lado}, mas a data não mudou.")
        page.wait_for_timeout(1000)
        el, atual = ler_data_painel(page)
        if not atual:
            falhar(page, "erro_data.png", "Perdi a leitura da data após clicar na seta.")
        print(f"Data do painel: {atual:%d/%m/%Y}")

    print("Data do painel confere com a data alvo.")


def solicitar_exportacao(page):
    # Passo 7: garante que o painel está no dia alvo antes de atualizar
    ajustar_data_painel(page)

    # Passo 8: Atualizar (ao lado de "Ao vivo")
    _, botao = esperar(page, lambda f: f.locator(f"xpath={XPATH_ATUALIZAR}"), timeout_s=90, habilitado=True)
    if not botao:
        falhar(page, "erro_atualizar.png", "O botão Atualizar não ficou clicável em 90s.")
    botao.scroll_into_view_if_needed()
    botao.click(timeout=5_000)
    print(f"Atualizando... aguardando {INTERVALO_APOS_ATUALIZAR_S}s antes de exportar.")
    page.wait_for_timeout(INTERVALO_APOS_ATUALIZAR_S * 1000)

    # Passo 9: seta "Alternar painel de exportações"
    nome = re.compile(r"alternar painel de exporta", re.I)
    _, botao = esperar(
        page,
        lambda f: f.get_by_role("button", name=nome),
        lambda f: f.locator("button[aria-label*='exporta' i], button[title*='exporta' i]"),
        lambda f: f.locator(f"xpath={XPATH_ALTERNAR_PAINEL}"),
        timeout_s=30,
    )
    if not botao:
        falhar(page, "erro_exportacoes.png", "Não achei 'Alternar painel de exportações'.")
    botao.click()
    page.wait_for_timeout(800)

    # Passo 10: relatório "Detalhes da duração do status do agente"
    _, item = esperar(
        page, lambda f: f.get_by_text("Detalhes da duração do status do agente"), timeout_s=30
    )
    if not item:
        falhar(page, "erro_relatorio.png", "Não achei o relatório no iframe de Análises.")
    item.click()

    # Passo 11: painel Formato -> layout -> Exportar
    frame, _ = esperar(page, lambda f: f.get_by_text("Formato"), timeout_s=60)
    if not frame:
        falhar(page, "erro_formato.png", "O painel 'Formato' não apareceu.")
    painel = frame.get_by_role("dialog").filter(has_text="Formato")
    if painel.count() == 0:
        painel = frame
    expandir_secao_formato(frame, painel)
    configurar_layout(frame, painel)

    exportar = painel.get_by_role("button", name="Exportar", exact=True).last
    try:
        exportar.wait_for(state="visible", timeout=15_000)
    except PWTimeout:
        falhar(page, "erro_formato.png", "Não achei o botão Exportar no painel.")
    exportar.click()
    print("Exportação solicitada.")
    page.wait_for_timeout(2000)


def abrir_exportacoes_na_caixa_entrada(page):
    inbox = page.locator(
        '[data-test-id="command-bar-inbox"], '
        'button[aria-label*="Caixa de entrada" i], button[title*="Caixa de entrada" i]'
    ).first
    inbox.wait_for(state="visible", timeout=30_000)
    inbox.click()

    padrao = re.compile(r"exporta(?:ções|coes|tions)?", re.I)
    _, alvo = esperar(
        page,
        lambda f: f.get_by_role("button", name=padrao),
        lambda f: f.get_by_role("tab", name=padrao),
        lambda f: f.get_by_text(padrao),
        timeout_s=30,
    )
    if not alvo:
        falhar(page, "erro_caixa_entrada.png", "Não achei 'Exportações' na Caixa de entrada.")
    alvo.click()


def baixar_exportacao(page):
    """Espera a seta de download (item de exportação pronto) e salva o arquivo."""
    page.wait_for_timeout(5_000)  # tempo para a exportação ser gerada e listada
    _, botao = esperar(
        page, lambda f: f.locator(f"xpath={XPATH_DOWNLOAD}"), timeout_s=120, habilitado=True
    )
    if not botao:
        falhar(page, "erro_download_exportacao.png", "A seta de download não apareceu em 120s.")
    with page.expect_download(timeout=120_000) as dl:
        botao.click()
    download = dl.value
    nome = download.suggested_filename or f"status_agente_{datetime.now():%Y-%m-%d_%H%M}.csv"
    destino = PASTA_DOWNLOAD / nome
    download.save_as(destino)
    return destino


# ───────────────────────── atualização do Excel ─────────────────────────

FORMATOS_DATA = ("%d/%m/%y %H:%M", "%d/%m/%Y %H:%M", "%d/%m/%y %H:%M:%S", "%d/%m/%Y %H:%M:%S")
BOOLEANOS = {"FALSO": False, "VERDADEIRO": True, "FALSE": False, "TRUE": True}
XL_CALC_MANUAL = -4135
XL_UP = -4162


def _data_hora(txt):
    for fmt in FORMATOS_DATA:
        try:
            return datetime.strptime(txt.strip(), fmt)
        except ValueError:
            pass
    raise ValueError(f"Data/hora fora do padrão esperado: {txt!r}")


def ler_csv_exportacao(caminho):
    """
    Lê o CSV do Genesys (separador ';') e devolve as linhas já tipadas, prontas para B:S da Base:
    datas -> datetime, "Intervalo concluído" -> bool, vazio -> None, durações -> texto como vieram.
    """
    caminho = Path(caminho)
    try:
        texto = caminho.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        texto = caminho.read_text(encoding="cp1252")
    linhas = list(csv.reader(io.StringIO(texto, newline=""), delimiter=";"))
    if not linhas:
        raise RuntimeError(f"O CSV {caminho.name} está vazio.")

    esperado = [c.casefold() for c in CABECALHO_CSV]
    recebido = [c.strip().casefold() for c in linhas[0]]
    if recebido != esperado:
        raise RuntimeError(
            "O cabeçalho do CSV mudou (colunas diferentes ou em outra ordem). "
            f"Esperado {len(esperado)} colunas, veio {len(recebido)}: {linhas[0]}. "
            "A aba Base NÃO foi alterada."
        )

    dados = []
    for n, lin in enumerate(linhas[1:], start=2):
        if not any(c.strip() for c in lin):
            continue
        if len(lin) != len(CABECALHO_CSV):
            raise RuntimeError(f"Linha {n} do CSV tem {len(lin)} colunas (esperado {len(CABECALHO_CSV)}).")
        concluido = BOOLEANOS.get(lin[2].strip().upper(), lin[2] or None)
        resto = [c if c != "" else None for c in lin[3:]]
        dados.append([_data_hora(lin[0]), _data_hora(lin[1]), concluido, *resto])

    if not dados:
        raise RuntimeError(
            f"O CSV {caminho.name} não tem linhas de dados. A aba Base NÃO foi alterada "
            "(para não apagar o dia sem ter o que colocar no lugar)."
        )
    return dados


def agrupar_blocos(numeros):
    """[5, 6, 7, 10] -> [(5, 7), (10, 10)]"""
    blocos = []
    for r in sorted(numeros):
        if blocos and r == blocos[-1][1] + 1:
            blocos[-1][1] = r
        else:
            blocos.append([r, r])
    return [tuple(b) for b in blocos]


def _e_do_dia(valor, dias):
    return isinstance(valor, datetime) and valor.date() in dias


def atualizar_base_excel(csv_path):
    """
    Na aba Base do Excel: apaga as linhas do(s) dia(s) que estão no CSV e insere as linhas do CSV.
    Usa o próprio Excel (xlwings) para NÃO perder gráfico, imagem, validações e fórmulas da Dash.
    Se qualquer conferência falhar, o arquivo não é salvo.
    """
    import xlwings as xw  # só esta etapa precisa do Excel instalado

    destino = Path(EXCEL_PAUSAS)
    if not destino.exists():
        raise FileNotFoundError(f"Excel não encontrado: {destino}. Ajuste EXCEL_PAUSAS no .env.")

    dados = ler_csv_exportacao(csv_path)
    dias = {linha[0].date() for linha in dados}
    dias_txt = ", ".join(f"{d:%d/%m/%Y}" for d in sorted(dias))
    print(f"CSV: {len(dados)} linhas do dia {dias_txt}.")

    # Backup rotativo antes de mexer
    PASTA_BACKUP.mkdir(exist_ok=True)
    copia = PASTA_BACKUP / f"{destino.stem}_{datetime.now():%Y%m%d_%H%M%S}{destino.suffix}"
    shutil.copy2(destino, copia)
    for antigo in sorted(PASTA_BACKUP.glob(f"{destino.stem}_*{destino.suffix}"))[:-MANTER_BACKUPS]:
        antigo.unlink()

    app = xw.App(visible=False, add_book=False)
    wb = None
    try:
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(str(destino), update_links=False)
        if wb.api.ReadOnly:
            raise RuntimeError(
                "O Excel abriu o arquivo como somente leitura (provavelmente está aberto em outra janela "
                "ou por outra pessoa). Feche-o e rode de novo."
            )
        ws = wb.sheets[ABA_BASE]
        api = ws.api

        calc_original = app.api.Calculation
        app.api.Calculation = XL_CALC_MANUAL  # evita recalcular a Dash a cada alteração

        ultima = api.Cells(api.Rows.Count, 2).End(XL_UP).Row
        if ultima < 2:
            ultima = 1

        # 1) remove o(s) dia(s) do CSV (limpa o conteúdo e, se o dia não for o último, sobe o que vem depois;
        #    não exclui linhas, para não encolher o intervalo $2:$15000 usado pelas fórmulas da Dash)
        existentes = ws.range(f"B2:B{ultima}").options(ndim=1).value if ultima >= 2 else []
        linhas_do_dia = [i + 2 for i, v in enumerate(existentes) if _e_do_dia(v, dias)]
        for ini, fim in reversed(agrupar_blocos(linhas_do_dia)):
            n = fim - ini + 1
            if fim < ultima:
                api.Range(f"A{fim + 1}:S{ultima}").Cut(api.Range(f"A{ini}"))
            api.Range(f"A{ultima - n + 1}:S{ultima}").ClearContents()
            ultima -= n
        print(f"Base: {len(linhas_do_dia)} linhas antigas do dia removidas.")

        # 2) insere as linhas novas no final
        primeira = ultima + 1
        ultima_nova = primeira + len(dados) - 1
        api.Range(f"A{primeira}:A{ultima_nova}").NumberFormat = api.Range("A2").NumberFormat
        api.Range(f"B{primeira}:C{ultima_nova}").NumberFormat = api.Range("B2").NumberFormat
        api.Range(f"E{primeira}:S{ultima_nova}").NumberFormat = "@"  # texto: evita o Excel converter " 03:04:42.589" em hora
        ws.range(f"B{primeira}").value = dados
        api.Range(f"A{primeira}:A{ultima_nova}").Formula = f"=INT(B{primeira})"

        # 3) conferência antes de salvar
        todas = ws.range(f"B2:B{ultima_nova}").options(ndim=1).value
        total_dia = sum(1 for v in todas if _e_do_dia(v, dias))
        if total_dia != len(dados):
            raise RuntimeError(
                f"Conferência falhou: a Base ficou com {total_dia} linhas do dia {dias_txt}, "
                f"esperado {len(dados)}. Nada foi salvo."
            )
        if ultima_nova > LINHA_LIMITE_DASH:
            print(f"ATENÇÃO: a Base passou da linha {LINHA_LIMITE_DASH}; a Dash não lê além dela.")

        app.api.Calculation = calc_original  # volta o modo de cálculo original (recalcula a Dash)
        wb.save()
        print(f"Excel atualizado: {len(dados)} linhas inseridas (linhas {primeira} a {ultima_nova}) em {destino.name}.")
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            app.quit()


# ───────────────────────── principal ─────────────────────────

def main():
    args = sys.argv[1:]

    # Só atualiza o Excel, sem abrir o Genesys: python script.py --so-excel [caminho_do_csv]
    # (sem caminho, usa o CSV mais recente da pasta downloads)
    if "--so-excel" in args:
        extras = [a for a in args if a != "--so-excel"]
        if extras:
            arquivo = Path(extras[0])
        else:
            csvs = sorted(PASTA_DOWNLOAD.glob("*.csv"), key=lambda p: p.stat().st_mtime)
            if not csvs:
                raise SystemExit(f"Nenhum CSV em {PASTA_DOWNLOAD}.")
            arquivo = csvs[-1]
        print(f"Usando o CSV: {arquivo}")
        atualizar_base_excel(arquivo)
        return

    if not LOGIN or not SENHA:
        raise SystemExit("Defina GENESYS_LOGIN e GENESYS_SENHA no arquivo .env")

    arquivo = None
    with sync_playwright() as p:
        contexto = p.chromium.launch_persistent_context(
            user_data_dir=str(PASTA_SESSAO),
            headless=HEADLESS,
            accept_downloads=True,
            args=["--disable-infobars", "--no-first-run", "--no-default-browser-check"],
        )
        contexto.set_default_timeout(TIMEOUT)
        page = contexto.pages[0] if contexto.pages else contexto.new_page()
        try:
            entrar(page)
            page.wait_for_timeout(1000)
            fechar_notificacoes(page, max_s=2)

            abrir_workspace_analises(page)
            solicitar_exportacao(page)
            abrir_exportacoes_na_caixa_entrada(page)
            arquivo = baixar_exportacao(page)
            print(f"Relatório salvo em: {arquivo.resolve()}")
        except Exception:
            tirar_print(page, "erro_geral.png")
            raise
        finally:
            contexto.close()

    # Navegador já fechado: agora atualiza o Excel
    atualizar_base_excel(arquivo)


if __name__ == "__main__":
    main()