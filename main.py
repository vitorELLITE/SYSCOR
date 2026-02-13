import time
import shutil
import os
import requests
import traceback
from pathlib import Path
from urllib.parse import urlparse, parse_qs, urlencode
from datetime import datetime, date, timedelta

# --- BIBLIOTECAS SELENIUM ---
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service # Importante para Chromium
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# CONFIGURAÇÕES
# ==========================================
SYSCOR_USER = "ELLITERELATORIOS"
SYSCOR_PASS = "Santos1414"
EMAIL_SPO = "spupdate@ellitecelular.com"
SENHA_SPO = "B/244938822972uz"
SHAREPOINT_BASE_URL = "https://vivoellitecelular.sharepoint.com"

# Caminhos internos do Docker
BASE = Path("/app")
PASTA_TEMP = BASE / "downloads"
PASTA_TEMP.mkdir(parents=True, exist_ok=True)

# Intervalo do Loop (30 minutos = 1800 segundos)
INTERVALO_SEGUNDOS = 1800 

# ==========================================
# FUNÇÕES DE APOIO
# ==========================================
def limpar_pasta_seguro(caminho: Path):
    if not caminho.exists():
        return
    for item in caminho.iterdir():
        try:
            if item.is_file(): item.unlink()
        except: pass

def link_ate_hoje_capado_no_trimestre(link_base: str) -> str:
    hoje = date.today()
    q = (hoje.month - 1) // 3
    inicio = date(hoje.year, q * 3 + 1, 1)
    fim_mes = q * 3 + 3
    if fim_mes == 12: fim = date(hoje.year, 12, 31)
    else: fim = date(hoje.year, fim_mes + 1, 1) - timedelta(days=1)
    fim_final = min(hoje, fim)
    parts = urlparse(link_base)
    q = parse_qs(parts.query, keep_blank_values=True)
    q['data_inicio'] = [inicio.strftime("%d/%m/%Y")]
    q['data_fim'] = [fim_final.strftime("%d/%m/%Y")]
    return parts._replace(query=urlencode(q, doseq=True)).geturl()

def esperar_download_robusto(download_dir: Path, lista_arquivos_antes: set, timeout=300) -> Path:
    limite = time.time() + timeout
    print(f"    ... Aguardando arquivo ...")
    while time.time() < limite:
        try: atuais = set(download_dir.glob("*"))
        except: continue
        novos = atuais - lista_arquivos_antes
        candidatos = []
        for p in novos:
            # Aceita extensões válidas e ignora temporários
            if p.suffix.lower() in ['.xls', '.xlsx', '.csv'] and not p.name.endswith('.crdownload'):
                if p.exists() and p.stat().st_size > 0:
                    candidatos.append(p)
        if candidatos:
            time.sleep(3) 
            if candidatos[0].exists(): return candidatos[0]
        time.sleep(1)
    raise Exception("Timeout: Nenhum arquivo válido apareceu.")

def upload_via_api_backend(driver, local_path: Path, server_relative_url: str):
    print(f"    [API] Enviando para: {server_relative_url}")
    
    # Pega cookies do Selenium para autenticar na API
    selenium_cookies = driver.get_cookies()
    session = requests.Session()
    for cookie in selenium_cookies:
        session.cookies.set(cookie['name'], cookie['value'])
    
    # Header essencial para evitar bloqueio
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    })

    try:
        # 1. Pega Token de Segurança
        context_url = f"{SHAREPOINT_BASE_URL}/sites/BI/_api/contextinfo"
        resp_digest = session.post(context_url, headers={"Accept": "application/json;odata=verbose"})
        
        if resp_digest.status_code != 200:
            print(f"    [ERRO API] Token negado: {resp_digest.status_code}")
            return False
            
        digest_value = resp_digest.json()['d']['GetContextWebInformation']['FormDigestValue']
        
        # 2. Upload do Arquivo
        file_name = local_path.name
        upload_url = f"{SHAREPOINT_BASE_URL}/sites/BI/_api/web/GetFolderByServerRelativeUrl('{server_relative_url}')/Files/add(url='{file_name}', overwrite=true)"
        
        with open(local_path, 'rb') as f:
            file_content = f.read()
            
        resp_upload = session.post(upload_url, data=file_content, headers={
            "X-RequestDigest": digest_value,
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/octet-stream"
        })
        
        if resp_upload.status_code in [200, 201]:
            print("    [SUCESSO] Upload via API concluído!")
            return True
        else:
            print(f"    [ERRO API] Falha: {resp_upload.status_code} - {resp_upload.text}")
            return False
    except Exception as e:
        print(f"    [ERRO FATAL API] {e}")
        return False

# ==========================================
# LÓGICA DO ROBÔ (CICLO)
# ==========================================
def executar_ciclo():
    print(f"\n[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] >>> INICIANDO TAREFA <<<")
    
    agora = datetime.now()
    trimestre = (agora.month - 1) // 3 + 1
    sufixo = f"{agora.year} {trimestre} Tri Rev. GO_MS"

    # LINKS ORIGINAIS SYSCOR
    LINK_SERVICOS = "https://syscor.before.com.br/_sys/relServicosVisaoExcel.php?fi_id=66,235,225,236,222,220,218,200,219,223,221,201,226,228,224,237,217,229,231,230,232,233,4,215,213,7,19,211,210,1,216,212,234,3,199,154,208,214,32,2,5,6,63,149,116,67,35,68,202,44,207&servicos=ALTA%20PRE,ALTA%20PRE%20DADOS,ALTA%20CONTROLE,ALTA%20SMARTVIVO%20CONTROLE,REATIVACAO%20CONTROLE,ALTA%20POS,ALTA%20POS%20SMARTPHONE,ALTA%20POS%20FIXO,ALTA%20INTERNET%20BOX,REATIVACAO%20POS,ALTA%20LINHA%20FIXA,REATIVACAO%20FIXA,ALTA%20INTERNET%20FIXA,ALTA%20VIVO%20TV,ALTA%20INTERNET%20MOVEL,ALTA%20SVA%20FIXO,ALTA%20SVA,ALTA%20SVA%20PRE,UPGRADE%20CONTROLE-POS%20PAGO%20FULL,MIGRACAO%20PRE%20CONTROLE,MIGRACAO%20PRE%20SMARTVIVO%20CONTROLE,MIGRACAO%20PRE%20POS,MIGRACAO%20PRE%20SMARTPHONE,MIGRACAO%20POS%20CONTROLE,DOWNGRADE%20INTERNET%20BOX,DOWNGRADE%20INTERNET%20MOVEL,DOWNGRADE%20CONTROLE,DOWNGRADE%20POS%20PAGO,DOWNGRADE%20POS%20PAGO%20FIXO,FIDELIZACAO%20CONTROLE,FIDELIZACAO%20POS,FIDELIZACAO%20POS%20SMARTPHONE,UPGRADE%20CONTROLE%20FULL,UPGRADE%20POS%20PAGO%20FULL,UPGRADE%20INTERNET%20BOX,UPGRADE%20INTERNET%20MOVEL,UPGRADE%20POS%20PAGO%20FIXO,UPGRADE%20FIXA,DOWNGRADE%20PRE,SEGURO,PORTABILIDADE%20CONTROLE,PORTABILIDADE%20POS,PORTABILIDADE%20FIXA,MIGRACAO%20TECNOLOGICA&indicador=2&num_protocolo_ativo=1&modo=1&data_inicio=01/02/2026&data_fim=12/02/2026&contadigital=0&campos_analitico=vendedor,cli_nome,ve_data_ins,vsv_data,ve_hora,nota_num,nota_num_serie,vsv_id,dependente,tipoTroca,tipoHabilitacao,fidelizado,celular,num_port,nome_operadora,sistema_origem,serial_celular,serial,promocao,recarga,recargaValor,recargaDesconto,num_protocolo,dia_venc,conta_digital,valorReceita,servico_remuneracao,valor_remuneracao,remuneracao_zerada,vsv_data_instalacao_servico,num_protocolo_status,vsv_servico_instalado,num_protocolo_ged,num_protocolo_data,vsv_zerar_rem_motivo,fixa_disponivel,oferta_fixa,vsv_linha_dependente,multiplicador,uf_sigla,num_protocolo_data_cancelamento,us_lider_nome&excel=1&considerarData=0&rel=1"
    LINK_PRODUTOS = "https://syscor.before.com.br/_sys/relVendaProdutoVisaoExcel.php?fi_id=66,235,225,222,220,218,200,219,223,221,201,226,228,224,237,217,4,215,213,7,211,210,1,216,212,234,3,199,154,32&removerSimcardDoado=1&modo=1&excel=1&produto_tipo=5,1,7,2,99&campos_analitico=ven_id,pm_nome,produto_nome,modelo_dpgc,plano_nome,serial,for_nome,valor_compra,valor_overlay,valor_desconto,valor_adicional,valor_venda,valor_lucro,cor,qtde_vendida,data_compra,ve_data,us_nome,promocao,ve_hora,valor_comissao,valor_price,nomeFilialCompra,cli_nome,ncm,sku,cupom_fiscal_numero,data_emissao,chave_acesso,vl_icms,vl_icms_st,vl_icms_fcpst,vl_ipi,vl_pis,vl_cofins,vl_custo_liquido,vl_vivo_renova,vl_desc_fabricante,vc_prom_id,num_acesso,uf_sigla,us_autorizacao,tipo_produto,categoria_tipo,subcategoria,multiplicador,valorMultiplicadorVenda,condicaoPagamento,valorProduto,vecg_cod,fv_nome,vc_ordem_renova,vc_voucher_venda_assistida&data_inicio=01/01/2026&data_fim=31/03/2026&agrupar=1&custoLiquidoPorMediaOuUltimaCompra=1&rel=1"
    LINK_USUARIOS = "https://syscor.before.com.br/_sys/relUsuarioVisao.php?fi_id=66,235,225,227,236,222,220,218,200,219,223,221,201,226,228,224,237,217,229,231,230,232,233,4,215,213,7,59,19,211,210,1,216,212,234,3,199,154,208,214,32&fi_id_acesso=66,235,225,227,236,222,220,218,200,219,223,221,201,226,228,224,237,217,229,231,230,232,233,4,215,213,7,59,19,211,210,1,216,212,234,3,199,154,208,214,32&modo=0&visualizar=0&us_campos=us_id,us_nome,us_login,us_ativo,us_cpf,us_rg,us_orgao_expedidor,us_estado_civil_id,us_rua,us_num,us_bairro,us_cep,uf_sigla,cid_nome,us_email,us_cel,us_data_nasc,us_data_adm,us_data_dem,us_compl,us_rem,us_ultimo_salario_alt,us_inss,us_comissao,ban_nome,us_ag,us_cc,us_dados_adicionais,data_inclusao,us_inclusao,data_alteracao,us_alteracao,us_acesso,fi_nome_acesso,us_data_experiencia,us_data_ferias,us_data_afastado,fure_nome,us_cbo,us_login_telefonica,dependentes,us_irrf,us_pis,fi_id,us_sexo,us_lider,us_acesso_ip,us_acesso_export_rel_sis,us_acesso_visualizacao_dados_sensiveis_rel,us_suspenso,us_suspenso_motivo,us_prestadora_nome,us_prestadora_cnpj&sms=0&whats=0&excel=1&rel=1"

    regras = [
        {
            "nome": f"relatorio_servico_analitico {sufixo}.csv",
            "link_sys": LINK_SERVICOS,
            "url_relativa": "/sites/BI/Servicos",
            "tipo": "servicos"
        },
        {
            "nome": f"relatorio_venda_produto {sufixo}.csv",
            "link_sys": LINK_PRODUTOS,
            "url_relativa": "/sites/BI/Produtos",
            "tipo": "produtos"
        },
        {
            "nome": f"relatorio_usuario {sufixo}.csv",
            "link_sys": LINK_USUARIOS,
            "url_relativa": "/sites/BI/DadosConfigurao",
            "tipo": "usuarios"
        },
    ]

    limpar_pasta_seguro(PASTA_TEMP)

    # Configuração do Chromium para Docker
    chrome_options = Options()
    # Aponta para o binário do Chromium instalado no Dockerfile
    chrome_options.binary_location = "/usr/bin/chromium" 
    
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

    # Inicia o Service apontando para o driver do sistema
    webdriver_service = Service("/usr/bin/chromedriver")

    prefs = {
        "download.default_directory": str(PASTA_TEMP),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = None
    try:
        print(">>> Iniciando Chromium (Docker)...")
        driver = webdriver.Chrome(service=webdriver_service, options=chrome_options)
        
        # Força download dir
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": str(PASTA_TEMP)})

        # Ajusta datas apenas para serviços
        for r in regras:
            if r.get("tipo") == "servicos":
                r["link_sys"] = link_ate_hoje_capado_no_trimestre(r["link_sys"])

        # 1. LOGIN SHAREPOINT
        print(">>> (1/3) Login SharePoint...")
        driver.get("https://vivoellitecelular.sharepoint.com/sites/BI")
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "loginfmt"))).send_keys(EMAIL_SPO + Keys.ENTER)
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "passwd"))).send_keys(SENHA_SPO + Keys.ENTER)
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        except: pass 
        
        print("    [Login] Aguardando autenticação completa (15s)...")
        time.sleep(15) 
        
        # 2. DOWNLOAD SYSCOR
        print(">>> (2/3) Baixando do Syscor...")
        driver.get("https://syscor.before.com.br/_sys/?logUfIdSession=12")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "us_email"))).send_keys(SYSCOR_USER)
        driver.find_element(By.ID, "us_senha").send_keys(SYSCOR_PASS)
        driver.find_element(By.CSS_SELECTOR, "input.submit.entrar").click()
        time.sleep(5)

        arquivos_prontos = []

        for r in regras:
            print(f"\n    -> Solicitando: {r['nome']}")
            snapshot_antes = set(PASTA_TEMP.glob("*"))
            driver.get(r["link_sys"])
            try:
                baixado = esperar_download_robusto(PASTA_TEMP, snapshot_antes, 1200 if r['tipo']=='servicos' else 300)
                destino = PASTA_TEMP / r["nome"]
                if destino.exists():
                    try: destino.unlink()
                    except: pass
                shutil.move(str(baixado), str(destino))
                arquivos_prontos.append(destino)
                print(f"       [OK] Baixado.")
            except Exception as e:
                print(f"       [ERRO] {e}")
                arquivos_prontos.append(None)

        # 3. UPLOAD API
        print("\n>>> (3/3) Upload API...")
        # Refresh para cookies
        driver.get("https://vivoellitecelular.sharepoint.com/sites/BI")
        time.sleep(5)
        
        for i, r in enumerate(regras):
            arquivo = arquivos_prontos[i]
            if arquivo and arquivo.exists():
                upload_via_api_backend(driver, arquivo, r['url_relativa'])
            else:
                print(f"    [PULAR] Arquivo não encontrado.")

    except Exception as e:
        print(f"ERRO NO CICLO: {e}")
        traceback.print_exc()
    finally:
        if driver:
            driver.quit()
        print(f"[{datetime.now().strftime('%H:%M:%S')}] >>> TAREFA CONCLUÍDA. <<<")

# ==========================================
# LOOP INFINITO
# ==========================================
if __name__ == "__main__":
    print(f"Iniciando Bot Syscor -> SharePoint.")
    print(f"Modo: Loop Infinito. Intervalo: {INTERVALO_SEGUNDOS}s ({int(INTERVALO_SEGUNDOS/60)} min)")
    
    while True:
        executar_ciclo()
        print(f"\n--- Dormindo por {int(INTERVALO_SEGUNDOS/60)} minutos... ---")
        time.sleep(INTERVALO_SEGUNDOS)
