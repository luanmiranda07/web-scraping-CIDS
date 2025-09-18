from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, ElementClickInterceptedException
import time
import random

# ===== Configurações de espera =====
WAIT_SHORT = 10        # cliques/cookies
WAIT_LONG  = 60        # carregamentos de páginas/tabelas
POLL_SLEEP = 0.25      # intervalo do polling leve
POST_CLICK_SLEEP = 0.35  # pausa pós-clique para animações

opts = Options()
# opts.add_argument("--headless=new")  # ative se quiser rodar sem abrir janela
opts.add_argument("--disable-notifications")
opts.add_argument("--headless=new")
opts.add_argument("--disable-gpu")
opts.add_argument("--no-sandbox")
opts.add_argument("--window-size=1366,768")  # ajuda no headless

driver = webdriver.Chrome(options=opts)
wait = WebDriverWait(driver, WAIT_LONG)

def switch_into_categorias(driver, wait):
    """
    Garante que estamos no contexto (main ou iframe) onde #tbCategorias existe.
    Depois de chamar, o driver fica dentro do frame correto (se houver).
    """
    def _in_categorias(drv):
        drv.switch_to.default_content()
        if drv.find_elements(By.ID, "tbCategorias"):
            return True
        frames = drv.find_elements(By.TAG_NAME, "iframe")
        for f in frames:
            drv.switch_to.frame(f)
            if drv.find_elements(By.ID, "tbCategorias"):
                return True
            drv.switch_to.default_content()
        return False

    wait.until(_in_categorias)

    driver.switch_to.default_content()
    if driver.find_elements(By.ID, "tbCategorias"):
        return
    for f in driver.find_elements(By.TAG_NAME, "iframe"):
        driver.switch_to.frame(f)
        if driver.find_elements(By.ID, "tbCategorias"):
            return
        driver.switch_to.default_content()

def scroll_center(elem):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)

def safe_click(elem):
    scroll_center(elem)
    try:
        elem.click()
    except Exception:
        driver.execute_script("arguments[0].click();", elem)

def click_and_wait(clickable, locator_to_wait, max_tries=3, base_sleep=POST_CLICK_SLEEP):
    """
    Clica em `clickable` e espera `locator_to_wait` aparecer.
    Tenta com backoff para lidar com latência/overlay.
    """
    for attempt in range(1, max_tries + 1):
        scroll_center(clickable)
        try:
            clickable.click()
        except (ElementClickInterceptedException, StaleElementReferenceException, Exception):
            driver.execute_script("arguments[0].click();", clickable)

        # micro-pausa para permitir render/transição
        time.sleep(base_sleep + (attempt - 1) * 0.25 + random.uniform(0, 0.15))
        try:
            WebDriverWait(driver, WAIT_LONG).until(EC.presence_of_element_located(locator_to_wait))
            return True
        except TimeoutException:
            if attempt == max_tries:
                return False
            time.sleep(0.4 + 0.2 * attempt)
    return False

def click_voltar():
    """Clica no botão Voltar da tela de detalhes (button id=btnVoltarTbListCategorias).
       Se não achar, usa driver.back() como fallback e espera a tabela principal.
    """
    try:
        btn_voltar = WebDriverWait(driver, WAIT_LONG).until(
            EC.element_to_be_clickable((By.ID, "btnVoltarTbListCategorias"))
        )
        ok = click_and_wait(btn_voltar, (By.ID, "tbCategorias"), max_tries=2)
        if not ok:
            driver.back()
            WebDriverWait(driver, WAIT_LONG).until(EC.presence_of_element_located((By.ID, "tbCategorias")))
    except TimeoutException:
        driver.back()
        WebDriverWait(driver, WAIT_LONG).until(EC.presence_of_element_located((By.ID, "tbCategorias")))

def go_next_page() -> bool:
    """Tenta ir para a próxima página de categorias. Retorna True se conseguiu, False se não há próxima."""
    switch_into_categorias(driver, wait)

    candidatos = [
        (By.XPATH, "//a[contains(., 'Próxima') or contains(., 'Proxima') or contains(., 'Next')]"),
        (By.XPATH, "//button[contains(., 'Próxima') or contains(., 'Proxima') or contains(., 'Next')]"),
        (By.CSS_SELECTOR, "[aria-label='Próxima página'], [aria-label='Proxima página'], [aria-label='Next']"),
        (By.CSS_SELECTOR, ".paginate_button.next, .pagination .next a, .pagination li.next a"),
        (By.XPATH, "//a[.//svg or .//i][contains(@class,'next') or contains(@aria-label,'Próxima') or contains(@aria-label,'Next')]"),
    ]

    for by, sel in candidatos:
        try:
            btns = driver.find_elements(by, sel)
            btns = [b for b in btns if b.is_displayed() and b.is_enabled()]
            if not btns:
                continue

            for btn in btns:
                cls = (btn.get_attribute("class") or "").lower()
                aria_disabled = (btn.get_attribute("aria-disabled") or "").lower()
                if "disabled" in cls or aria_disabled == "true":
                    continue

                linhas_antes = driver.find_elements(By.CSS_SELECTOR, "#tbCategorias > tbody > tr")
                num_antes = len(linhas_antes)

                safe_click(btn)

                try:
                    WebDriverWait(driver, WAIT_LONG).until(EC.presence_of_element_located((By.ID, "tbCategorias")))
                    for _ in range(40):  # até ~10s
                        time.sleep(POLL_SLEEP)
                        linhas_depois = driver.find_elements(By.CSS_SELECTOR, "#tbCategorias > tbody > tr")
                        if len(linhas_depois) != num_antes or linhas_depois != linhas_antes:
                            return True
                    # mesmo número; pode ser última página
                    return False
                except TimeoutException:
                    switch_into_categorias(driver, wait)
                    return True
        except Exception:
            continue
    return False

# ===== NOVO: setar 100 por página =====
def set_page_size_100():
    """Seleciona 100 resultados por página no seletor #tbCategorias_length > label > select e
    espera a tabela redesenhar (mais linhas na página)."""
    switch_into_categorias(driver, wait)

    # Espera o select aparecer
    select_el = WebDriverWait(driver, WAIT_LONG).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#tbCategorias_length > label > select"))
    )

    # Quantas linhas existem antes da mudança?
    linhas_antes = driver.find_elements(By.CSS_SELECTOR, "#tbCategorias > tbody > tr")
    num_antes = len(linhas_antes)

    try:
        # Tenta pelo Select padrão
        sel = Select(select_el)
        # Tente por value (geralmente "100") e se falhar, por texto
        try:
            sel.select_by_value("100")
        except Exception:
            sel.select_by_visible_text("100")
    except Exception:
        # Fallback JS caso o Select falhe por overlay/estilo custom
        driver.execute_script("""
            const s = document.querySelector('#tbCategorias_length > label > select');
            if (s) { s.value = '100'; s.dispatchEvent(new Event('change', {bubbles: true})); }
        """)

    # Aguardar o redraw da tabela (número de linhas aumentar ou mudar)
    for _ in range(60):  # até ~15s (60 * 0.25)
        time.sleep(POLL_SLEEP)
        linhas_depois = driver.find_elements(By.CSS_SELECTOR, "#tbCategorias > tbody > tr")
        if len(linhas_depois) > num_antes or linhas_depois != linhas_antes:
            break

try:
    driver.get("https://www.cremesp.org.br/?siteAcao=cid10")

    # Aceitar cookies se aparecer
    try:
        btn_cookie = WebDriverWait(driver, WAIT_SHORT).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Ciente') or contains(., 'OK')]"))
        )
        safe_click(btn_cookie)
    except Exception:
        pass

    # Entrar no contexto correto da tabela
    switch_into_categorias(driver, wait)

    # >>>>>> AQUI: setar 100 por página uma única vez
    set_page_size_100()

    pagina = 1
    while True:
        switch_into_categorias(driver, wait)

        linhas = driver.find_elements(By.CSS_SELECTOR, "#tbCategorias > tbody > tr")
        if not linhas:
            break

        print(f"\n=== Página {pagina} | Categorias visíveis: {len(linhas)} ===")

        # Percorre TODAS as linhas desta página
        for i in range(len(linhas)):
            switch_into_categorias(driver, wait)
            linhas = driver.find_elements(By.CSS_SELECTOR, "#tbCategorias > tbody > tr")
            linha = linhas[i]

            tds = linha.find_elements(By.TAG_NAME, "td")
            codigo = tds[0].text.strip() if len(tds) > 0 else ""
            descricao = tds[1].text.strip() if len(tds) > 1 else ""
            print(f"\nCategoria: {codigo} - {descricao}")

            # Clicar no botão 'olho' (3ª coluna). Evitar clicar no SVG/path.
            try:
                botao = linha.find_element(By.CSS_SELECTOR, "td:nth-child(3) button")
            except Exception:
                botao = linha.find_element(By.XPATH, ".//td[3]//button | .//td[3]//a")

            # Clique com retry e espera pela TELA DE DETALHE (sinal: botão Voltar)
            ok = click_and_wait(botao, (By.ID, "btnVoltarTbListCategorias"), max_tries=3)
            if not ok:
                driver.execute_script("arguments[0].click();", botao)
                WebDriverWait(driver, WAIT_LONG).until(EC.presence_of_element_located((By.ID, "btnVoltarTbListCategorias")))

            # Agora buscar a tabela (algumas categorias podem não ter linhas)
            tabela = driver.find_elements(By.ID, "tabela_body")
            if not tabela:
                tabela = driver.find_elements(By.CSS_SELECTOR, "[id*='tabela_body']")

            detalhas = []
            if tabela:
                # Polling leve até virem linhas (ou aceitar vazio)
                for _ in range(int(8 / POLL_SLEEP)):  # ~8s total
                    detalhas = driver.find_elements(By.CSS_SELECTOR, "[id*='tabela_body'] > tr")
                    if detalhas:
                        break
                    time.sleep(POLL_SLEEP)

            # Raspar os CIDs (se houver)
            for row in detalhas:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 2:
                    cid_codigo = cols[0].text.strip()
                    cid_desc = cols[1].text.strip()
                    print(f"   {cid_codigo} - {cid_desc}")

            # Voltar para a lista de categorias (com retry + espera pela lista)
            click_voltar()
            switch_into_categorias(driver, wait)

            time.sleep(0.3)  # Pausa curta entre categorias

        # Tentar ir para a próxima página; se não tiver, encerra
        if go_next_page():
            pagina += 1
            time.sleep(0.5)  # Pausa curta entre páginas
            continue
        else:
            break

finally:
    driver.quit()
