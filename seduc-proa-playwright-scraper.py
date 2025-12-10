import os
import re
import time
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

MODO_DEBUG = False

# --- 1. CONFIGURAÇÕES INICIAIS ---

CONFIG_DIRETORIOS = {
    "CAMINHO_ENV": r"C:\Users\joao-dasilva\PycharmProjects\sibelle-automation\scrapper\.env",
    "DIRETORIO_ATUAL": os.getcwd(),
    "PASTA_RAIZ_DADOS": os.path.join(os.getcwd(), "_dados_teste_proa"),
    "DB_PATH": os.path.join(os.getcwd(), "_dados_teste_proa", "controle_processos_baixados.csv"),
}
#  --- 2. Database ---
COLUMNS = {
    "proa_notificatorio", "num_pg", "assunto", "data_abertura", "grupo_origem", "grupo_portador", "ultima_analise_feita", "ultimo_download_feito"
}

# --- 3. SELETORES ATUALIZADOS ---
CONFIG_LOGIN = {
    "ORGANIZACAO": '//*[@id="organizacao"]',
    "USUARIO": '//*[@id="matricula"]',
    "SENHA": '//*[@id="senha"]',
    "BOTAO_LOGIN": '//*[@id="btnLogonOrganizacao"]',
}

CONFIG_CAIXAS = [
    {
        "nome_interno": "DEPARTAMENTO DE OBRAS ESCOLARES - NOTIFICATORIO",       # Nome para salvar no CSV/Log
        "filtro_grupo_label": "DMOE-NOT", # O texto exato para buscar no Dropdown do site
        "pasta_destino": r"C:\Users\joao-dasilva\PycharmProjects\sibelle-automation\scrapper\_dados_teste_proa\PDF_NOT",     # Onde salvar os arquivos desta caixa
        "guia_tabela": "DMOE-NOT"
    },
    {
        "nome_interno": "DEPARTAMENTO DE OBRAS ESCOLARES - MINISTÉRIO PÚBLICO",
        "filtro_grupo_label": "DMOE-MP",
        "pasta_destino": r"C:\Users\joao-dasilva\PycharmProjects\sibelle-automation\scrapper\_dados_teste_proa\PDF_MP",
        "guia_tabela": "DMOE-MP"
},
]

PAINEL_POS_LOGIN = {
    "MINHAS_ATIVIDADES": '[aria-label="Minhas Atividades"]',
    "MEUS_PRAZOS": '[aria-label="Meus Prazos"]',
    "MINHAS_TAREFAS": '[aria-label="Minhas Tarefas"]',
    "NOVO_PROCESSO": '[aria-label="Novo Processo"]',
    "PESQUISA_POR_CONTEUDO": '[aria-label="Pesquisa por Conteúdo"]',
    "PESQUISA_AVANCADA": '//*[@id="panelListaProcessos"]/div/span[5]'
}
# --- MAPEAMENTO DOS CAMPOS DO FORMULÁRIO ---
CONFIG_FILTROS = {

    "situacao_todos": "//label[@for='form:situacaoProcesso:2']",     # 1. Situação: "Todos" (Radio Button)
    "tipo_orgao_portador": "//label[@for='form:tpOrgao:2']",         # 2. Tipo de Órgão: "Portador" (Radio Button)
    "assunto_todos": "//label[@for='form:indAtivoAssuntoDoc:2']",    # 3. Assunto: "Todos" (Radio Button - Cuidado: Fica dentro do bloco Assunto)
    "verificar_orgao_texto": "//*[@id='form:orgaoED_label']",        # 4. Verificar se Órgão é "SE" (Texto estático) Pegue o .inner_text() deste elemento para validar se é == "SE"
    "grupo_trigger": "//*[@id='form:grupoOrganizacaoEDOrigem']",     # 5. Mudar Grupo (Dropdown)

    # Passo B: Clicar no item da lista (Use .format(seu_label_config))
    # O ID com '_panel' é a lista flutuante segura do PrimeFaces
    "grupo_opcao": "//div[@id='form:grupoOrganizacaoEDOrigem_panel']//li[contains(text(), '{}')]",

    # Botão para aplicar tudo
    "btn_pesquisar": "//button[span[text()='Pesquisar']]"
}


class ProaBot:
    def __init__(self):
        self.playwright = None
        self.CONFIG_LOGIN = CONFIG_LOGIN
        self.CONFIG_CAIXAS = CONFIG_CAIXAS
        self.PAINEL_POS_LOGIN = PAINEL_POS_LOGIN
        self.CONFIG_FILTROS = CONFIG_FILTROS
        self.CONFIG_DIRETORIOS = CONFIG_DIRETORIOS

        self.pasta_raiz = os.path.join(os.getcwd(), "_dados_teste_proa")
        self.db_path = os.path.join(self.pasta_raiz, "controle_processos_baixados.csv")

        self.browser = None
        self.context = None
        self.page = None
        self.base_url = "https://secweb.procergs.com.br/pra-aj4/mod-processo/processoAdministrativo-list.xhtml?is-cobrado=false"

    def iniciar(self):

        print(">>> Verificando estrutura de pastas <<<")
        if os.path.exists(CONFIG_DIRETORIOS["CAMINHO_ENV"]):
            load_dotenv(CONFIG_DIRETORIOS["CAMINHO_ENV"])
            print(">>> .env carregada! <<<")
        else:
            print(f"ERRO: .env não encontrado em {CONFIG_DIRETORIOS["CAMINHO_ENV"]}")
            exit()

        # 1. Cria a pasta raiz (_dados_teste_proa) se não existir
        if not os.path.exists(CONFIG_DIRETORIOS["PASTA_RAIZ_DADOS"]):
            os.makedirs(CONFIG_DIRETORIOS["PASTA_RAIZ_DADOS"])
            print(f"[CRIADO] Raiz: {CONFIG_DIRETORIOS["PASTA_RAIZ_DADOS"]}")
        else:
            print(f"[OK] Raiz: {CONFIG_DIRETORIOS["PASTA_RAIZ_DADOS"]}")

        # 2. Loop para verificar/criar as pastas de cada caixa na CONFIG
        for caixa in CONFIG_CAIXAS:
            caminho_pasta = caixa["pasta_destino"]

            # O parâmetro exist_ok=True faz com que ele não dê erro se a pasta já existir
            os.makedirs(caminho_pasta, exist_ok=True)

            # Apenas para feedback visual no console:
            if os.path.exists(caminho_pasta):
                print(f"   -> [OK] Pasta verificada: {caixa['guia_tabela']}")
            else:
                print(f"   -> [CRIADO] Pasta nova: {caixa['guia_tabela']}")

        print("-> [OK] Ambiente pronto\n")
        print(">>> Abrindo browser <<<")
        playwright = sync_playwright().start()                   # ← start manual
        self.browser = playwright.chromium.launch(headless=not MODO_DEBUG)
        self.context = self.browser.new_context(
            viewport={'width': 1366, 'height': 768},
            accept_downloads=True
        )
        self.page = self.context.new_page()
        if MODO_DEBUG == True:
            input("Enter para começar: ")
        return self

    def login(self):
        print(">>> INICIANDO O LOGIN <<<")
        organizacao = os.getenv("PROA_ORG")
        usuario = os.getenv("PROA_USER")
        senha = os.getenv("PROA_PASS")
        try:
            try:
                print("1. Acessando página de Login...")
                self.page.goto(
                    "https://secweb.procergs.com.br/pra-aj4/mod-processo/processoAdministrativo-list.xhtml?is-cobrado=false")
                print("   Tela carregada.")
                # Espera o campo de organização aparecer
                self.page.wait_for_selector(CONFIG_LOGIN["ORGANIZACAO"], timeout=10000)

            except Exception as e:
                print(f"Não foi possivel carregar o site {e}")
                self.page.screenshot(path="erro_ao_carregar_site.png")

            print("   Preenchendo...")
            # Preenchimento
            # Dica: Adicionei delays pequenos entre digitações para parecer
            self.page.wait_for_selector(CONFIG_LOGIN["ORGANIZACAO"], state="visible", timeout=20000)
            self.page.fill(CONFIG_LOGIN["ORGANIZACAO"], organizacao);

            self.page.wait_for_selector(CONFIG_LOGIN["USUARIO"], state="visible", timeout=20000)
            self.page.fill(CONFIG_LOGIN["USUARIO"], usuario);

            self.page.wait_for_selector(CONFIG_LOGIN["SENHA"], state="visible", timeout=20000)
            self.page.fill(CONFIG_LOGIN["SENHA"], senha);

            print("   Clicando em entrar...")
            self.page.click(CONFIG_LOGIN["BOTAO_LOGIN"], force=True)

            # --- AQUI ESTÁ O SEGREDO DO "RAPIDO DEMAIS" ---

            if MODO_DEBUG == True:
                input("Enter para começar: ")
        except Exception as e:
            print(f"Erro no Login: {e}")
            return
    #--------------VERIFICANDO-O-LOGIN------------------------------
        try:
            print("   [AGUARDANDO] Login enviado. Aguardando site carregar...")
            time.sleep(0.5)
            if any(self.page.is_visible(sel) for sel in PAINEL_POS_LOGIN.values()):
                print("   -> Logado! (Carregamento imediato)")
            else:
                print("   -> Painel não encontrado de imediato. Aguardando mais um pouco...")
                self.page.wait_for_selector(PAINEL_POS_LOGIN["PESQUISA_AVANCADA"], state="visible", timeout=20000)
                print("   -> Logado! (Carregamento finalizado)")

        except Exception as e:
            print(f"❌ ERRO CRÍTICO: Não foi possível confirmar o login. O site pode estar fora do ar ou a senha incorreta.")
            print(f"Detalhe do erro: {e}")
            self.page.screenshot(path="erro_login_timeout.png")
            raise e
        return self

    def abrir_pesquisa_avancada(self):
        selector = PAINEL_POS_LOGIN["PESQUISA_AVANCADA"]
        try:
            if  self.page.wait_for_selector(selector, state="visible", timeout=10000):  # 10s timeout
                print("2. Abrindo Pesquisa Avançada...")
                self.page.click(selector)

            time.sleep(2)
            if any(self.page.is_visible(sel) for sel in CONFIG_FILTROS.values()):
                print("   -> Sucesso: Entrou na Pesquisa Avançada!")
            else:
                raise Exception("Falha na verificação: Elementos internos da Pesquisa Avançada não encontrados.")
        except Exception as e:
            print(f"Não foi possível acessar o painel de pesquisa avançada: {e}")
            self.page.screenshot(path="erro_pesquisa_avancada.png")
        return self


    def aplicar_filtro(self, nome_filtro):
        print(">>> Iniciando aplicação dos filtros <<<")

        # Defina as chaves dos filtros iniciais para clareza e reutilização
        filtros_iniciais_chaves = ["situacao_todos", "tipo_orgao_portador", "assunto_todos"]


        try:
            print("   Verificando visibilidade dos filtros iniciais...")

            # Checa se QUALQUER um dos filtros iniciais está visível para confirmar que estamos na página correta
            if any(self.page.is_visible(CONFIG_FILTROS[chave]) for chave in filtros_iniciais_chaves):
                print("   -> Filtros iniciais visíveis.")

                # Aplicar cada filtro inicial com wait individual para robustez
                for chave in filtros_iniciais_chaves:
                    selector = CONFIG_FILTROS[chave]
                    if self.page.is_visible(selector):
                        self.page.click(selector)
                        self.page.wait_for_timeout(500)
                        print(f"      Seletor clicado: {chave} ...")
                    else:
                        print(f"   [Aviso] O filtro '{chave}' não estava visível e foi pulado.")

                print("   -> Filtros iniciais aplicados com sucesso!")
            else:
                raise Exception("Painel de Pesquisa Avançada parece não estar ativo (nenhum filtro encontrado).")

            # Validação do Órgão Atual
            print("   -> Validando órgão atual...")

            selector_orgao = CONFIG_FILTROS["verificar_orgao_texto"]
            self.page.wait_for_selector(selector_orgao, state="visible", timeout=5000)
            orgao_atual = self.page.inner_text(selector_orgao).strip()

            if orgao_atual != "SE":
                raise ValueError(f"ERRO CRÍTICO: O robô esperava estar no órgão 'SE', mas está em '{orgao_atual}'.")

            print("   -> Validação de Órgão (SE): OK")

            config_encontrada = None

            # Procura na lista CONFIG_CAIXAS qual dicionário tem esse rótulo
            for item in CONFIG_CAIXAS:
                if item['filtro_grupo_label'] == nome_filtro:  # Compara com "DMOE-NOT"
                    config_encontrada = item
                    break

            if not config_encontrada:
                raise ValueError(f"Filtro '{nome_filtro}' não encontrado no CONFIG_CAIXAS")

            # Agora define a variável alvo baseada no que achou
            nome_grupo_alvo = config_encontrada['filtro_grupo_label']

            print(f">>> Iniciando aplicação dos filtros para: {nome_grupo_alvo} <<<")

            # A: Clicar no trigger para abrir o combo (dropdown)
            # Nota: CONFIG_FILTROS parece ser outro dicionário global, se funcionar, mantenha.
            selector_trigger = CONFIG_FILTROS["grupo_trigger"]
            self.page.wait_for_selector(selector_trigger, state="visible", timeout=5000)
            self.page.click(selector_trigger)

            # B: Montar e aguardar o seletor dinâmico da opção do grupo
            seletor_item_grupo = CONFIG_FILTROS["grupo_opcao"].format(nome_grupo_alvo)
            self.page.wait_for_selector(seletor_item_grupo, state="visible",
                                        timeout=5000)

            # C: Clicar na opção
            self.page.click(seletor_item_grupo)

            print(f">>> Grupo '{nome_grupo_alvo}' selecionado com sucesso.")

            #Pesquisa
            print(">>> Clicando em Pesquisar...")
            self.page.click(CONFIG_FILTROS["btn_pesquisar"])

            # E: Verificar se a pesquisa retornou resultados
            # Estratégia: Esperamos aparecer a linha com data-ri="0" dentro do corpo da tabela.
            # Se aparecer, significa que a tabela NÃO está vazia.
            seletor_primeira_linha = "tbody[id*='lista_data'] tr[data-ri='0']"

            try:
                # Espera até 10 segundos para ver se a primeira linha aparece
                self.page.wait_for_selector(seletor_primeira_linha, state="visible", timeout=10000)

                print(">>> SUCESSO: A pesquisa retornou itens na lista.")
                pesquisa_com_sucesso = True

            except Exception as e:
                # Se der timeout, significa que a linha 0 nunca apareceu (tabela vazia ou erro)
                print(">>> AVISO: A pesquisa não retornou nenhum item ou demorou demais.")
                pesquisa_com_sucesso = False

                # Opcional: Tirar um print para debug se falhar
                self.page.screenshot(path="debug_erro_pesquisa.png")

        except Exception as e:
            print(f"Não foi possível aplicar os filtros: {e}")
            # Captura de tela para debug em caso de erro
            self.page.screenshot(path="erro_aplicar_filtros.png")
            raise  # Re-levanta o erro para propagação, se necessário
        return self

    def carregar_banco_dados(self):
        db_path = self.db_path  # assumindo que você tem self.db_path no __init__ como os.path.join(self.pasta_raiz, "controle_processos.csv")
        if os.path.exists(db_path):
            return pd.read_csv(db_path, dtype=str)
        return pd.DataFrame(columns=[
            "proa_notificatorio", "num_pg", "assunto", "data_abertura",
            "grupo_origem", "grupo_portador", "ultima_analise_feita", "ultimo_download_feito"
        ])

    def salvar_banco_dados(self, novos_dados: dict):
        df = self.carregar_banco_dados()

        # Converte os novos dados pra um DF temporário (pra 1 linha só)
        df_novo = pd.DataFrame([novos_dados])

        # Checa se o proa_notificatorio já existe (update se sim, append se não)
        if 'proa_notificatorio' in df.columns and df_novo['proa_notificatorio'].iloc[0] in df[
            'proa_notificatorio'].values:
            # Update: merge mantendo as colunas existentes e atualizando as novas
            df = df.merge(df_novo, on='proa_notificatorio', how='left', suffixes=('', '_new'))
            for col in df_novo.columns:
                if col != 'proa_notificatorio':
                    df[col] = df[col + '_new'].combine_first(df[col])
            df = df.drop(columns=[col for col in df.columns if col.endswith('_new')])
        else:
            # Novo: só append
            df = pd.concat([df, df_novo], ignore_index=True)

        # Salva de volta
        df.to_csv(self.db_path, index=False)
        print(f"DB atualizado: {len(df)} linhas totais.")

    def limpar_nome_arquivo(self, texto):
        return re.sub(r'[\\/*?:"<>|]', "", texto)

    def preparando_lista(self):
        # Seletores
        SELETOR_COLUNA_DATA = "th[id='form:lista:j_idt543']"
        # O estado final que queremos no HTML
        ESTADO_DESEJADO = "descending"

        print("\n>>> FASE 1.5: ORDENANDO LISTA E AJUSTANDO TAMANHO <<<")

        # -----------------------------------------------------------
        # PARTE A: ORDENAÇÃO (Com paciência para o site carregar)
        # -----------------------------------------------------------

        # Tenta no máximo 4 vezes (Padrão -> Asc -> Desc -> Padrão -> Asc...)
        for i in range(1, 5):
            # 1. Pega o elemento
            coluna = self.page.locator(SELETOR_COLUNA_DATA)

            # 2. Lê o estado ATUAL (pode ser None, 'ascending' ou 'descending')
            estado_atual = coluna.get_attribute("aria-sort")
            print(f"   [Tentativa {i}] Estado atual da coluna: {estado_atual}")

            # 3. VERIFICAÇÃO: Se já estiver certo, para tudo!
            if estado_atual == ESTADO_DESEJADO:
                print("   ✅ Ordenação Descendente alcançada com sucesso!")
                break

            # 4. AÇÃO: Clica para mudar o estado
            print("   -> Clicando para alterar ordem...")
            coluna.click()

            # 5. ESPERA INTELIGENTE (O Segredo)
            # Espera 2 segundos fixos para o AJAX processar e a tabela recarregar
            self.page.wait_for_timeout(2000)

            # Opcional: Se o site for muito lento, descomente a linha abaixo
            # self.page.wait_for_load_state("networkidle")

        # Se saiu do loop e ainda não está certo, aí sim gera erro
        estado_final = self.page.locator(SELETOR_COLUNA_DATA).get_attribute("aria-sort")
        if estado_final != ESTADO_DESEJADO:
            # Tira um print para você ver como ficou antes de quebrar
            self.page.screenshot(path="erro_ordenacao.png")
            raise Exception(f"ERRO: Ordenação falhou. Esperado: descending, Obtido: {estado_final}")

        print("--- Ordenação Concluída ---\n")

        # -----------------------------------------------------------
        # PARTE B: MAXIMIZAR LINHAS POR PÁGINA (PARA 100)
        # -----------------------------------------------------------
        SELETOR_LINHAS_POR_PAGINA = "select[name='form:lista_rppDD']"
        VALOR_DESEJADO = "100"

        try:
            print(f">>> Alterando a exibição para {VALOR_DESEJADO} linhas...")
            self.page.select_option(SELETOR_LINHAS_POR_PAGINA, value=VALOR_DESEJADO)

            # Espera novamente a tabela recarregar após mudar o tamanho
            self.page.wait_for_timeout(2000)
            print("   ✅ Exibição alterada.")

        except Exception as e:
            print(f"AVISO: Não foi possível alterar para 100 linhas: {e}")

        return self

    def coletar_links(self):
        print("\n=======================================================")
        print(">>> FASE 1: COLETANDO LINKS (COM PAGINAÇÃO AUTOMÁTICA) <<<")
        print("=======================================================")

        self.links_para_processar = []
        SELETOR_PROXIMA_PAGINA = "a.ui-paginator-next:not(.ui-state-disabled)"
        BASE_URL_DUPLICADA = "/pra-aj4/pra-aj4/"
        BASE_URL_CORRETA = "/pra-aj4/"

        # Pega todas as linhas do corpo da tabela
        SELETOR_LINHAS = "tbody[id*='lista_data'] tr"
        # Seletor do link dentro da linha
        SELETOR_LINK = 'td a[href*="processoAdministrativo-form.xhtml"]'

        pagina_atual = 1
        ultimo_primeiro_processo = None  # Variável para verificar se a página mudou

        while True:
            print(f"\n--- Processando Página {pagina_atual} ---")

            # 1. Aguarda as linhas aparecerem (garante que não está vazio)
            try:
                self.page.wait_for_selector(SELETOR_LINHAS, timeout=10000)
            except:
                print("   AVISO: Timeout aguardando linhas. Lista pode ter acabado.")
                break

            # 2. Captura todas as linhas visíveis
            linhas = self.page.locator(SELETOR_LINHAS).all()
            qtd_linhas = len(linhas)
            print(f"   -> Encontradas {qtd_linhas} linhas.")

            if qtd_linhas == 0:
                break

            # --- TRAVA DE SEGURANÇA: Verifica se não estamos lendo a página anterior ---
            # Pega o texto do primeiro processo da lista atual
            try:
                primeiro_processo_atual = linhas[0].locator(SELETOR_LINK).first.inner_text().strip()

                if primeiro_processo_atual == ultimo_primeiro_processo:
                    print("   ⚠️ ALERTA: A página parece não ter atualizado (os dados são iguais aos da anterior).")
                    # Opcional: break ou continue com retry. Aqui vamos assumir que acabou.

                ultimo_primeiro_processo = primeiro_processo_atual
            except:
                pass
                # ---------------------------------------------------------------------------

            # 3. Coleta os dados
            for linha in linhas:
                link_loc = linha.locator(SELETOR_LINK).first

                if link_loc.count() > 0:
                    href = link_loc.get_attribute('href')
                    numero = link_loc.inner_text().strip()

                    if href:
                        # Correção da URL
                        base_url = self.page.url.split('/mod-processo/')[0]
                        link_completo = base_url + href
                        if BASE_URL_DUPLICADA in link_completo:
                            link_completo = link_completo.replace(BASE_URL_DUPLICADA, BASE_URL_CORRETA, 1)

                        self.links_para_processar.append({
                            'numero': numero,
                            'link_detalhe': link_completo
                        })

            print(f"   -> Acumulado: {len(self.links_para_processar)} links.")

            # 4. Paginação
            botao_prox = self.page.locator(SELETOR_PROXIMA_PAGINA)

            if botao_prox.is_visible():
                print("   -> Clicando em 'Próxima'...")
                botao_prox.click()

                # O PULO DO GATO: Espera o primeiro item MUDAR ou a tabela recarregar
                # Como PrimeFaces é chato, um sleep generoso é a solução mais segura hoje
                self.page.wait_for_timeout(3000)

                pagina_atual += 1
            else:
                print("   -> Fim da paginação (botão Próximo sumiu ou desabilitado).")
                break

        print(f"Total FINAL coletado: {len(self.links_para_processar)}")
        return self

    def processar_downloads(self, filtro_label):
        print("\n=======================================================")
        print(f">>> FASE 2: DOWNLOAD COM FALLBACK INTELIGENTE ({filtro_label}) <<<")
        print("=======================================================")

        # 1. Configuração Inicial
        config_alvo = next((item for item in self.CONFIG_CAIXAS if item['filtro_grupo_label'] == filtro_label), None)
        if not config_alvo:
            print(f"❌ ERRO: Configuração não encontrada para '{filtro_label}'.")
            return self

        pasta_destino = config_alvo['pasta_destino']
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino, exist_ok=True)

        if not self.links_para_processar:
            print("AVISO: Lista vazia.")
            return self

        # Seletores
        SELETOR_VISUALIZAR = "#visualizar-processo"
        SELETOR_BOTAO_DOWNLOAD = "button.download"

        # URL Base para o Fallback
        URL_BASE_DIRECT = "https://secweb.procergs.com.br/pra-aj4/visualizaProcesso?numero={}&download=true"

        erros_consecutivos = 0
        LIMITE_ERROS = 5
        total = len(self.links_para_processar)

        for i, item in enumerate(self.links_para_processar):

            # --- PARADA DE EMERGÊNCIA ---
            if erros_consecutivos >= LIMITE_ERROS:
                print(f"🚨 5 Falhas Totais Consecutivas. Parando execução.")
                self.fechar()
                return self

            numero_processo = item['numero']
            url_visual = item['link_detalhe']

            # Preparação dos nomes (Necessário para ambos os métodos)
            numero_limpo = re.sub(r'[/\.-]', '', numero_processo)
            nome_arquivo = f"Processo_Administrativo_{numero_limpo}.pdf"
            caminho_final = os.path.join(pasta_destino, nome_arquivo)

            # Verifica se já existe
            if os.path.exists(caminho_final):
                print(f"[{i + 1}/{total}] ⏩ Já existe: {nome_arquivo}")
                continue

            print(f"\n[{i + 1}/{total}] Processando: {numero_processo}")

            download_sucesso = False  # Flag de controle

            # =================================================================
            # TENTATIVA 1: MÉTODO VISUAL (Navegação + Clique)
            # =================================================================
            try:
                self.page.goto(url_visual, timeout=60000)

                # (Aqui entraria sua extração de dados futura)

                # Garante modal aberto
                try:
                    if self.page.is_visible(SELETOR_VISUALIZAR):
                        self.page.click(SELETOR_VISUALIZAR)
                        self.page.wait_for_timeout(1000)
                except:
                    pass

                # Tenta baixar clicando
                with self.page.expect_download(timeout=30000) as download_info:
                    self.page.click(SELETOR_BOTAO_DOWNLOAD, force=True)

                download = download_info.value
                download.save_as(caminho_final)

                print(f"   ✅ Sucesso (Via Visual): {nome_arquivo}")
                download_sucesso = True
                erros_consecutivos = 0  # Zera contador se der certo

            except Exception as e:
                print(f"   ⚠️ Falha Visual ({e}). Iniciando Fallback...")

                # =================================================================
                # TENTATIVA 2: FALLBACK (Link Direto + Retry)
                # =================================================================
                link_direto = URL_BASE_DIRECT.format(numero_limpo)

                # Loop de tentativas do fallback (Tenta agora, e se falhar, tenta denovo)
                for tentativa in range(1, 3):
                    try:
                        print(f"      -> 🔄 Fallback (Link Direto) - Tentativa {tentativa}...")

                        with self.page.expect_download(timeout=45000) as download_info:
                            try:
                                self.page.goto(link_direto)
                            except:
                                pass  # Ignora erro de navegação se o download começar

                        download = download_info.value
                        download.save_as(caminho_final)

                        print(f"      ✅ Recuperado com Sucesso (Via Fallback)!")
                        download_sucesso = True
                        erros_consecutivos = 0
                        break  # Sai do loop de tentativas se conseguir

                    except Exception as e_fallback:
                        if tentativa == 1:
                            print("      ❌ Falhou. Aguardando 5s para tentar novamente...")
                            time.sleep(5)  # Espera um tempinho como solicitado
                        else:
                            print(f"      ❌ Falhou na segunda tentativa: {e_fallback}")

            # =================================================================
            # RESULTADO FINAL DO ITEM
            # =================================================================
            if not download_sucesso:
                print(f"   ❌❌ FALHA TOTAL no processo {numero_processo}")
                # AQUI ENTRARIA O CÓDIGO PARA SALVAR NA PLANILHA DE CONTROLE (FUTURO)

                erros_consecutivos += 1

        print("\n>>> Fim dos downloads. <<<")
        return self

    def fechar(self):
        if MODO_DEBUG:
            print("========================================================")
            print("   PAUSA PARA DEBUG: O navegador permanecerá aberto.")
            print("   Inspecione a página (F12) se necessário.")
            print("   👉 Pressione [ENTER] aqui no console para fechar...")
            print("========================================================")
            input()

            try:
                if self.context:
                    self.context.close()
                if self.browser:
                    self.browser.close()
                if self.playwright:
                    self.playwright.stop()
                print(">>> Recursos liberados e browser fechado. <<<")

            except Exception as e:
                print(f"Erro ao tentar fechar o navegador: {e}")

if __name__ == "__main__":
    bot = ProaBot()
    filtro = "DMOE-MP"
    try:
        (
        bot.iniciar()
            .login()
            .abrir_pesquisa_avancada()
            .aplicar_filtro(filtro)
            .preparando_lista()
            .coletar_links()
            .processar_downloads(filtro)
        )
    except Exception as e:
        import traceback;
        traceback.print_exc()
    finally:
        bot.fechar()
