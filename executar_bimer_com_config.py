"""
Script para executar automação do Bimer usando configurações do config.yaml
Execute este script DENTRO da VM onde o Bimer está aberto
Utiliza as mesmas ações do testar_login_bimer.py mas lê configurações do YAML
"""
import time
import pyautogui
import logging
import yaml
from pathlib import Path
from datetime import datetime

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger(__name__)

# Configurações de segurança do pyautogui
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

# ============================================================
# CARREGAR CONFIGURAÇÕES DO YAML
# ============================================================

def carregar_config():
    """Carrega configurações do arquivo config.yaml"""
    config_path = Path(__file__).parent / "config.yaml"
    
    if not config_path.exists():
        logger.error(f"❌ Arquivo config.yaml não encontrado em: {config_path}")
        return None
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        logger.info(f"✓ Configurações carregadas de: {config_path}")
        return config
    except Exception as e:
        logger.error(f"❌ Erro ao carregar config.yaml: {e}")
        return None

# Carregar configurações
config = carregar_config()
if not config:
    logger.error("❌ Não foi possível carregar as configurações. Abortando.")
    exit(1)

# Extrair configurações do Bimer
bimer_config = config.get('bimmer', {})
login_config = bimer_config.get('login', {})
ui_elements = bimer_config.get('ui_elements', {})

# Credenciais
SENHA_BIMER = login_config.get('password', 'Rpa@@2025')

# Coordenadas de login (valores padrão se não estiverem no config)
DROPDOWN_AMBIENTE_X = ui_elements.get('dropdown_ambiente_x', 866)
DROPDOWN_AMBIENTE_Y = ui_elements.get('dropdown_ambiente_y', 579)
AMBIENTE_TESTE_X = ui_elements.get('ambiente_teste_x', 974)
AMBIENTE_TESTE_Y = ui_elements.get('ambiente_teste_y', 677)
CAMPO_SENHA_X = ui_elements.get('campo_senha_x', 904)
CAMPO_SENHA_Y = ui_elements.get('campo_senha_y', 520)
BOTAO_ENTRAR_X = ui_elements.get('botao_entrar_x', ui_elements.get('entrar_bimer_x', 953))
BOTAO_ENTRAR_Y = ui_elements.get('botao_entrar_y', ui_elements.get('entrar_bimer_y', 645))
FECHAR_MODAL_X = ui_elements.get('fechar_modal_x', 1391)
FECHAR_MODAL_Y = ui_elements.get('fechar_modal_y', 192)

# Coordenadas para quando HOUVER títulos para processar
MARCAR_TODOS_TITULOS_X = ui_elements.get('marcar_todos_titulos_x', 619)
MARCAR_TODOS_TITULOS_Y = ui_elements.get('marcar_todos_titulos_y', 335)
CAMPO_NOME_ARQUIVO_X = ui_elements.get('campo_nome_arquivo_x', 804)
CAMPO_NOME_ARQUIVO_Y = ui_elements.get('campo_nome_arquivo_y', 732)
BOTAO_GERAR_ARQUIVO_X = ui_elements.get('botao_gerar_arquivo_x', 1312)
BOTAO_GERAR_ARQUIVO_Y = ui_elements.get('botao_gerar_arquivo_y', 838)
BOTAO_SIM_CONFIRMACAO_X = ui_elements.get('botao_sim_confirmacao_x', 919)
BOTAO_SIM_CONFIRMACAO_Y = ui_elements.get('botao_sim_confirmacao_y', 581)
BOTAO_OK_OPERACAO_CONCLUIDA_X = ui_elements.get('botao_ok_operacao_concluida_x', 949)
BOTAO_OK_OPERACAO_CONCLUIDA_Y = ui_elements.get('botao_ok_operacao_concluida_y', 559)

# ============================================================
# FUNÇÕES DE DATA E FERIADOS
# ============================================================

# Lista de feriados nacionais fixos (adicione feriados móveis manualmente)
FERIADOS_NACIONAIS = [
    "01/01",  # Ano Novo
    "21/04",  # Tiradentes
    "01/05",  # Dia do Trabalho
    "07/09",  # Independência do Brasil
    "12/10",  # Nossa Senhora Aparecida
    "02/11",  # Finados
    "15/11",  # Proclamação da República
    "20/11",  # Consciência Negra
    "25/12",  # Natal
]

# Feriados móveis 2025 (atualizar anualmente)
FERIADOS_MOVEIS_2025 = [
    "03/03",  # Carnaval
    "04/03",  # Carnaval
    "18/04",  # Sexta-feira Santa
    "30/05",  # Corpus Christi
]

def eh_feriado(data):
    """Verifica se a data é feriado nacional"""
    dia_mes = data.strftime("%d/%m")
    data_completa = data.strftime("%d/%m")
    
    # Verifica feriados fixos
    if dia_mes in FERIADOS_NACIONAIS:
        return True
    
    # Verifica feriados móveis do ano atual
    if data.year == 2025 and data_completa in FERIADOS_MOVEIS_2025:
        return True
    
    return False

def eh_dia_util(data):
    """Verifica se a data é dia útil (não é fim de semana nem feriado)"""
    # 5 = Sábado, 6 = Domingo
    if data.weekday() >= 5:
        return False
    
    if eh_feriado(data):
        return False
    
    return True

def obter_ultimo_dia_util():
    """Retorna o último dia útil antes de hoje"""
    from datetime import timedelta
    
    hoje = datetime.now()
    data = hoje - timedelta(days=1)
    
    # Volta até encontrar um dia útil
    while not eh_dia_util(data):
        data = data - timedelta(days=1)
    
    return data

def obter_periodo_busca():
    """
    Retorna o período de busca (data_inicio, data_fim) considerando:
    - Se hoje é dia útil: busca apenas hoje
    - Se hoje é fim de semana/feriado: busca desde o último dia útil até hoje
    """
    from datetime import timedelta
    
    hoje = datetime.now()
    
    # Se hoje é dia útil, busca apenas hoje
    if eh_dia_util(hoje):
        # Verifica se ontem foi dia útil
        ontem = hoje - timedelta(days=1)
        if eh_dia_util(ontem):
            # Ontem foi dia útil, busca apenas hoje
            return hoje, hoje
        else:
            # Ontem não foi dia útil, busca desde o último dia útil
            ultimo_dia_util = obter_ultimo_dia_util()
            return ultimo_dia_util, hoje
    else:
        # Hoje não é dia útil, não deveria executar, mas se executar:
        # Busca desde o último dia útil até hoje
        ultimo_dia_util = obter_ultimo_dia_util()
        return ultimo_dia_util, hoje

def obter_data_atual():
    """Retorna a data atual no formato dd/mm/aaaa"""
    return datetime.now().strftime("%d/%m/%Y")

def obter_data_inicio_busca():
    """Retorna a data de início da busca no formato dd/mm/aaaa"""
    data_inicio, _ = obter_periodo_busca()
    return data_inicio.strftime("%d/%m/%Y")

def obter_data_fim_busca():
    """Retorna a data de fim da busca no formato dd/mm/aaaa"""
    _, data_fim = obter_periodo_busca()
    return data_fim.strftime("%d/%m/%Y")

def obter_caminho_completo_arquivo_remessa(empresa_numero):
    r"""
    Retorna o caminho completo do arquivo de remessa: C:\TEMP\RPA\REMDDMMAAAA_EMP##.TXT
    Em caso de feriado/fim de semana, usa a data de início da busca (último dia útil)
    """
    hoje = datetime.now()
    
    # Se hoje não é dia útil, usar a data de início da busca (último dia útil)
    if not eh_dia_util(hoje):
        data_inicio, _ = obter_periodo_busca()
        data_str = data_inicio.strftime('%d%m%Y')
    else:
        data_str = hoje.strftime('%d%m%Y')
    
    # Formatar número da empresa com zeros à esquerda (1 -> 01, 20 -> 20)
    emp_formatado = str(empresa_numero).zfill(2)
    nome_arquivo = f"REM{data_str}_EMP{emp_formatado}.TXT"
    
    return f"C:\\TEMP\\RPA\\{nome_arquivo}"

# ============================================================
# FUNÇÕES DE PROCESSAMENTO DE TÍTULOS
# ============================================================

# Variável global para rastrear resultados por empresa
resultados_empresas = {
    "1": {"tem_titulos": False, "arquivo_gerado": False},
    "2": {"tem_titulos": False, "arquivo_gerado": False},
    "20": {"tem_titulos": False, "arquivo_gerado": False}
}

def aguardar(segundos, mensagem=""):
    """Aguarda um tempo especificado"""
    if mensagem:
        logger.info(f"⏳ {mensagem}")
    time.sleep(segundos)

def verificar_e_processar_titulos(empresa_numero):
    """
    Verifica se há títulos para processar e executa as ações apropriadas.
    Estratégia DEFINITIVA: Usar input() para perguntar ao usuário
    - Aguarda 5 segundos para o usuário ver a tela
    - Pergunta se há títulos
    - Se SIM → Marcar todos e gerar arquivo
    - Se NÃO → Clicar OK e fechar modal
    """
    logger.info(f"   🔍 Verificando se há títulos para Empresa {empresa_numero}...")
    aguardar(5.0, "Aguardando tela carregar (5s)...")
    
    # Perguntar ao usuário se há títulos
    logger.info("")
    logger.info("=" * 70)
    logger.info(f"⚠️  ATENÇÃO: Verifique a tela do Bimer!")
    logger.info(f"   Empresa {empresa_numero} - Há títulos para processar?")
    logger.info("=" * 70)
    
    resposta = input("Digite 'S' se HÁ títulos ou 'N' se NÃO HÁ títulos: ").strip().upper()
    
    if resposta == 'S':
        # HÁ TÍTULOS - Processar
        logger.info(f"   ✅ TÍTULOS ENCONTRADOS para Empresa {empresa_numero}!")
        logger.info(f"   → Marcando todos os títulos...")
        
        # Marcar todos
        pyautogui.click(MARCAR_TODOS_TITULOS_X, MARCAR_TODOS_TITULOS_Y)
        aguardar(1.0)
        
        # Definir caminho completo do arquivo
        caminho_completo = obter_caminho_completo_arquivo_remessa(empresa_numero)
        logger.info(f"   → Definindo caminho do arquivo: {caminho_completo}")
        
        # Clicar no campo de arquivo
        pyautogui.click(CAMPO_NOME_ARQUIVO_X, CAMPO_NOME_ARQUIVO_Y)
        aguardar(0.5)
        
        # Limpar e digitar caminho
        pyautogui.hotkey('ctrl', 'a')
        aguardar(0.2)
        pyautogui.press('backspace')
        aguardar(0.3)
        pyautogui.typewrite(caminho_completo, interval=0.05)
        aguardar(0.5)
        
        # Gerar arquivo
        logger.info(f"   → Clicando no botão Gerar arquivo...")
        pyautogui.click(BOTAO_GERAR_ARQUIVO_X, BOTAO_GERAR_ARQUIVO_Y)
        aguardar(2.0)
        
        # Confirmar
        logger.info(f"   → Confirmando geração (Sim)...")
        pyautogui.click(BOTAO_SIM_CONFIRMACAO_X, BOTAO_SIM_CONFIRMACAO_Y)
        aguardar(3.0)
        
        # OK final
        logger.info(f"   → Confirmando operação concluída (OK)...")
        pyautogui.click(BOTAO_OK_OPERACAO_CONCLUIDA_X, BOTAO_OK_OPERACAO_CONCLUIDA_Y)
        aguardar(1.0)
        
        # Atualizar resultado
        resultados_empresas[empresa_numero]["tem_titulos"] = True
        resultados_empresas[empresa_numero]["arquivo_gerado"] = True
        
        logger.info(f"   ✅ Arquivo gerado com sucesso!")
        logger.info(f"   📁 Local: {caminho_completo}")
        return True
        
    else:
        # NÃO HÁ TÍTULOS - Fechar modal
        logger.info(f"   ⚠️  Nenhum título encontrado para Empresa {empresa_numero}")
        logger.info(f"   → Clicando em OK no modal 'Sem títulos'...")
        
        # Clicar OK
        pyautogui.click(945, 553)
        aguardar(1.0)
        
        # Fechar modal de remessa
        logger.info(f"   → Fechando modal de remessa...")
        pyautogui.click(1458, 186)
        aguardar(1.0)
        
        # Atualizar resultado
        resultados_empresas[empresa_numero]["tem_titulos"] = False
        resultados_empresas[empresa_numero]["arquivo_gerado"] = False
        
        logger.info(f"   ✓ Modal fechado - continuando para próxima empresa...")
        return False

# Sequência de cliques pós-login (MESMA do testar_login_bimer.py)
CLIQUES_POS_LOGIN = [
    # PRIMEIRA EMPRESA - Busca inicial
    ("Financeiro (menu lateral)", 106, 308, None),
    ("A Pagar", 89, 349, None),
    ("Fechar modal A Pagar", 1389, 192, None),
    ("Campo Empresa - Definir para 1", 106, 209, "1"),
    ("Ferramentas", 196, 63, None),
    ("Gerar Arquivo Remessa", 377, 112, None),
    ("Uma Conta", 449, 163, None),
    ("Campo Número da Conta", 648, 360, "14"),
    ("Campo Layout do Arquivo", 649, 403, "12"),
    ("Data Vencimento Programado (início)", 859, 550, "DATA_INICIO"),
    ("Data Vencimento Programado (fim)", 980, 551, "DATA_FIM"),
    ("Botão Ambos", 1032, 679, None),
    ("Botão Avançar - Filtros", 1171, 836, None),
    ("Botão Avançar - Formas de Pagamento", 1171, 836, None),
    ("Botão Avançar - Naturezas de Lançamento", 1171, 836, None),
    ("Botão Avançar - Pessoas", 1171, 836, None),
    ("Botão Avançar - Mapas de Carregamento", 1171, 836, None),
    ("VERIFICAR_TITULOS_EMPRESA_1", None, None, None),  # Ação condicional (já fecha modal se necessário)
    
    # TROCAR PARA SEGUNDA EMPRESA
    ("Campo Empresa - Trocar para 2", 106, 209, "2"),
    
    # SEGUNDA EMPRESA - Repetir busca
    ("Ferramentas - Empresa 2", 196, 63, None),
    ("Gerar Arquivo Remessa - Empresa 2", 377, 112, None),
    ("Uma Conta - Empresa 2", 449, 163, None),
    ("Campo Número da Conta - Empresa 2", 648, 360, "14"),
    ("Campo Layout do Arquivo - Empresa 2", 649, 403, "12"),
    ("Data Vencimento Programado (início) - Empresa 2", 859, 550, "DATA_INICIO"),
    ("Data Vencimento Programado (fim) - Empresa 2", 980, 551, "DATA_FIM"),
    ("Botão Ambos - Empresa 2", 1032, 679, None),
    ("Botão Avançar - Filtros - Empresa 2", 1171, 836, None),
    ("Botão Avançar - Formas de Pagamento - Empresa 2", 1171, 836, None),
    ("Botão Avançar - Naturezas de Lançamento - Empresa 2", 1171, 836, None),
    ("Botão Avançar - Pessoas - Empresa 2", 1171, 836, None),
    ("Botão Avançar - Mapas de Carregamento - Empresa 2", 1171, 836, None),
    ("VERIFICAR_TITULOS_EMPRESA_2", None, None, None),  # Ação condicional (já fecha modal se necessário)
    
    # TROCAR PARA TERCEIRA EMPRESA
    ("Campo Empresa - Trocar para 20", 106, 209, "20"),
    
    # TERCEIRA EMPRESA (20) - Repetir busca
    ("Ferramentas - Empresa 20", 196, 63, None),
    ("Gerar Arquivo Remessa - Empresa 20", 377, 112, None),
    ("Uma Conta - Empresa 20", 449, 163, None),
    ("Campo Número da Conta - Empresa 20", 648, 360, "14"),
    ("Campo Layout do Arquivo - Empresa 20", 649, 403, "12"),
    ("Data Vencimento Programado (início) - Empresa 20", 859, 550, "DATA_INICIO"),
    ("Data Vencimento Programado (fim) - Empresa 20", 980, 551, "DATA_FIM"),
    ("Botão Ambos - Empresa 20", 1032, 679, None),
    ("Botão Avançar - Filtros - Empresa 20", 1171, 836, None),
    ("Botão Avançar - Formas de Pagamento - Empresa 20", 1171, 836, None),
    ("Botão Avançar - Naturezas de Lançamento - Empresa 20", 1171, 836, None),
    ("Botão Avançar - Pessoas - Empresa 20", 1171, 836, None),
    ("Botão Avançar - Mapas de Carregamento - Empresa 20", 1171, 836, None),
    ("VERIFICAR_TITULOS_EMPRESA_20", None, None, None),  # Ação condicional (já fecha modal se necessário)
    
    # FINALIZAR SISTEMA
    ("Fechar Bimer", 1904, 6, None),
    ("Confirmar Fechar Sistema", 871, 552, None),
]

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def copiar_para_clipboard(texto):
    """Copia texto para área de transferência"""
    try:
        import pyperclip
        pyperclip.copy(texto)
        return True
    except ImportError:
        logger.warning("⚠️  pyperclip não instalado - usando digitação direta")
        return False

def aguardar(segundos, mensagem=""):
    """Aguarda um tempo específico com mensagem opcional"""
    if mensagem:
        logger.info(f"⏳ {mensagem}")
    time.sleep(segundos)

# ============================================================
# FUNÇÃO PRINCIPAL DE LOGIN
# ============================================================

def executar_login_bimer():
    """
    Executa o processo completo de login no Bimer
    """
    try:
        logger.info("")
        logger.info("=" * 70)
        logger.info("🤖 INICIANDO AUTOMAÇÃO DO BIMER (COM CONFIG.YAML)")
        logger.info("=" * 70)
        logger.info("")
        logger.info(f"📋 Configurações:")
        logger.info(f"   • Senha: {'*' * len(SENHA_BIMER)}")
        logger.info(f"   • Botão Entrar: ({BOTAO_ENTRAR_X}, {BOTAO_ENTRAR_Y})")
        logger.info(f"   • Fechar Modal: ({FECHAR_MODAL_X}, {FECHAR_MODAL_Y})")
        logger.info("")
        
        # ========================================
        # PASSO 1: Selecionar ambiente TESTE
        # ========================================
        logger.info("=" * 70)
        logger.info("[PASSO 1/4] SELECIONANDO AMBIENTE TESTE")
        logger.info("=" * 70)
        
        logger.info(f"→ Clicando no dropdown de ambiente em ({DROPDOWN_AMBIENTE_X}, {DROPDOWN_AMBIENTE_Y})")
        pyautogui.click(DROPDOWN_AMBIENTE_X, DROPDOWN_AMBIENTE_Y)
        aguardar(0.5)
        logger.info("✓ Dropdown aberto")
        
        logger.info(f"→ Selecionando TESTE em ({AMBIENTE_TESTE_X}, {AMBIENTE_TESTE_Y})")
        pyautogui.click(AMBIENTE_TESTE_X, AMBIENTE_TESTE_Y)
        aguardar(0.5)
        logger.info("✓ Ambiente TESTE selecionado")
        
        # ========================================
        # PASSO 2: Preencher senha
        # ========================================
        logger.info("")
        logger.info("=" * 70)
        logger.info("[PASSO 2/4] PREENCHENDO SENHA")
        logger.info("=" * 70)
        
        logger.info(f"→ Clicando no campo de senha em ({CAMPO_SENHA_X}, {CAMPO_SENHA_Y})")
        pyautogui.click(CAMPO_SENHA_X, CAMPO_SENHA_Y)
        aguardar(0.3)
        
        # Tentar usar clipboard primeiro (mais rápido e confiável)
        if copiar_para_clipboard(SENHA_BIMER):
            logger.info("→ Colando senha via clipboard (Ctrl+V)")
            pyautogui.hotkey('ctrl', 'v')
        else:
            logger.info("→ Digitando senha caractere por caractere")
            pyautogui.write(SENHA_BIMER, interval=0.1)
        
        aguardar(0.5)
        logger.info("✓ Senha preenchida")
        
        # ========================================
        # PASSO 3: Clicar em Entrar
        # ========================================
        logger.info("")
        logger.info("=" * 70)
        logger.info("[PASSO 3/4] CLICANDO NO BOTÃO ENTRAR")
        logger.info("=" * 70)
        
        if not BOTAO_ENTRAR_X or not BOTAO_ENTRAR_Y:
            logger.warning("⚠️  Coordenadas do botão Entrar não configuradas. Usando tecla Enter.")
            pyautogui.press('enter')
            aguardar(8.0, "Aguardando sistema abrir e carregar (8s)...")
        else:
            logger.info(f"→ Clicando em Entrar em ({BOTAO_ENTRAR_X}, {BOTAO_ENTRAR_Y})")
            pyautogui.click(BOTAO_ENTRAR_X, BOTAO_ENTRAR_Y)
            aguardar(8.0, "Aguardando sistema abrir e carregar (8s)...")
            logger.info("✓ Botão Entrar clicado")
        
        # ========================================
        # PASSO 4: Pós-login (modal e cliques)
        # ========================================
        logger.info("")
        logger.info("=" * 70)
        logger.info("[PASSO 4/4] AÇÕES PÓS-LOGIN")
        logger.info("=" * 70)
        
        aguardar(2.0, "Aguardando modal inicial aparecer...")
        
        # Fechar modal automático se configurado
        if FECHAR_MODAL_X and FECHAR_MODAL_Y:
            logger.info(f"→ Tentando fechar modal em ({FECHAR_MODAL_X}, {FECHAR_MODAL_Y})")
            try:
                pyautogui.click(FECHAR_MODAL_X, FECHAR_MODAL_Y)
                aguardar(0.5)
                logger.info("✓ Modal fechado")
            except Exception as e:
                logger.warning(f"⚠️  Falha ao fechar modal: {e}")
        else:
            logger.info("ℹ️  Coordenadas do modal não configuradas - pulando")
        
        # Executar sequência de cliques pós-login
        if CLIQUES_POS_LOGIN:
            logger.info(f"→ Executando {len(CLIQUES_POS_LOGIN)} cliques pós-login...")
            for i, (nome, x, y, acao) in enumerate(CLIQUES_POS_LOGIN, 1):
                # Verificar se é uma ação condicional
                if nome.startswith("VERIFICAR_TITULOS_"):
                    empresa_num = nome.replace("VERIFICAR_TITULOS_EMPRESA_", "")
                    logger.info(f"  [{i}/{len(CLIQUES_POS_LOGIN)}] {nome}")
                    verificar_e_processar_titulos(empresa_num)
                    continue
                
                logger.info(f"  [{i}/{len(CLIQUES_POS_LOGIN)}] {nome} em ({x}, {y})")
                pyautogui.click(x, y)
                
                # Tempos de espera específicos por tipo de ação
                if "a pagar" in nome.lower() and "fechar" not in nome.lower() and "voltar" not in nome.lower():
                    aguardar(20.0, "Aguardando tela A Pagar carregar completamente (20s)...")
                elif "a pagar - voltar" in nome.lower():
                    aguardar(1.5, "Voltando para A Pagar...")
                elif "empresa" in nome.lower() and ("trocar" in nome.lower() or "definir" in nome.lower()):
                    aguardar(0.5, "Configurando empresa...")
                elif "modal" in nome.lower() or "fechar" in nome.lower():
                    aguardar(0.8, "Aguardando modal fechar...")
                elif "menu" in nome.lower() or "financeiro" in nome.lower():
                    aguardar(1.0, "Aguardando menu expandir...")
                elif "ferramentas" in nome.lower():
                    aguardar(0.8, "Aguardando menu abrir...")
                elif "remessa" in nome.lower():
                    aguardar(1.0, "Aguardando submenu...")
                elif "uma conta" in nome.lower():
                    aguardar(1.5, "Aguardando modal de remessa abrir...")
                elif "campo" in nome.lower() or "data" in nome.lower():
                    aguardar(0.3, "Campo focado...")
                elif "avançar" in nome.lower():
                    aguardar(1.5, "Aguardando próxima tela carregar...")
                elif "sem títulos" in nome.lower() or "ok -" in nome.lower():
                    if "empresa 20" in nome.lower():
                        empresa_num = "20"
                    elif "empresa 2" in nome.lower():
                        empresa_num = "2"
                    else:
                        empresa_num = "1"
                    logger.warning(f"⚠️  EMPRESA {empresa_num}: Não há títulos para processamento")
                    aguardar(0.5, "Confirmando mensagem...")
                elif "botão" in nome.lower() or "ambos" in nome.lower():
                    aguardar(0.5, "Clique processado...")
                elif "bimer" in nome.lower() and "fechar" in nome.lower():
                    aguardar(1.0, "Fechando Bimer...")
                elif "confirmar" in nome.lower():
                    aguardar(0.5, "Confirmando...")
                else:
                    aguardar(0.5)
                
                # Se houver ação adicional (ex: digitar)
                if acao:
                    # Substituir marcadores de data pelos valores reais
                    if acao == "DATA_ATUAL":
                        valor_digitar = obter_data_atual()
                    elif acao == "DATA_INICIO":
                        valor_digitar = obter_data_inicio_busca()
                    elif acao == "DATA_FIM":
                        valor_digitar = obter_data_fim_busca()
                    else:
                        valor_digitar = acao
                    
                    logger.info(f"       → Digitando: {valor_digitar}")
                    
                    if isinstance(valor_digitar, str):
                        # Para campo de empresa: SEMPRE limpar antes de digitar
                        if "empresa" in nome.lower() and ("trocar" in nome.lower() or "definir" in nome.lower()):
                            logger.info(f"       → Limpando campo de empresa (tinha valor anterior)")
                            # Clicar no campo para garantir foco
                            pyautogui.click(x, y)
                            aguardar(0.3)
                            # Fazer triplo clique manualmente (3 cliques rápidos para selecionar tudo)
                            pyautogui.click(x, y)
                            aguardar(0.05)
                            pyautogui.click(x, y)
                            aguardar(0.05)
                            pyautogui.click(x, y)
                            aguardar(0.3)
                            # Apagar múltiplas vezes para garantir que o campo está limpo
                            for _ in range(10):  # Apagar até 10 caracteres
                                pyautogui.press('backspace')
                                aguardar(0.05)
                            aguardar(0.3)
                            # Digitar o novo número
                            logger.info(f"       → Digitando: {valor_digitar}")
                            pyautogui.write(valor_digitar, interval=0.15)
                            aguardar(0.5)
                            # Pressionar Enter para confirmar
                            logger.info(f"       → Confirmando com Enter")
                            pyautogui.press('enter')
                            aguardar(1.0)
                        else:
                            # Para outros campos: usar Ctrl+A
                            pyautogui.hotkey('ctrl', 'a')
                            aguardar(0.2)
                            pyautogui.press('delete')
                            aguardar(0.2)
                            pyautogui.write(valor_digitar, interval=0.1)
                            aguardar(0.3)
                            # Pressionar Enter para confirmar
                            logger.info(f"       → Confirmando com Enter")
                            pyautogui.press('enter')
                            aguardar(0.5)
            logger.info("✓ Sequência de cliques concluída")
        else:
            logger.info("ℹ️  Nenhum clique pós-login configurado")
        
        # ========================================
        # CONCLUSÃO
        # ========================================
        logger.info("")
        logger.info("=" * 70)
        logger.info("✅ PROCESSO COMPLETO EXECUTADO COM SUCESSO!")
        logger.info("=" * 70)
        logger.info("")
        # Obter informações do período de busca
        data_inicio = obter_data_inicio_busca()
        data_fim = obter_data_fim_busca()
        hoje = datetime.now()
        eh_util = eh_dia_util(hoje)
        
        logger.info("📊 RESUMO DA EXECUÇÃO:")
        logger.info("   ✓ Login realizado")
        logger.info(f"   📅 Data de execução: {obter_data_atual()} ({'Dia útil' if eh_util else 'Fim de semana/Feriado'})")
        logger.info(f"   🔍 Período de busca: {data_inicio} até {data_fim}")
        logger.info("")
        logger.info("   🏢 EMPRESA 1:")
        logger.info("      ✓ Navegação: Financeiro → A Pagar")
        logger.info("      ✓ Filtros: Conta 14, Layout 12")
        logger.info(f"      ✓ Período: {data_inicio} até {data_fim}")
        if resultados_empresas["1"]["tem_titulos"]:
            logger.info("      ✅ Resultado: Títulos encontrados e arquivo gerado!")
        else:
            logger.info("      ⚠️  Resultado: Sem títulos para processamento")
        logger.info("")
        logger.info("   🏢 EMPRESA 2:")
        logger.info("      ✓ Troca de empresa realizada")
        logger.info("      ✓ Filtros: Conta 14, Layout 12")
        logger.info(f"      ✓ Período: {data_inicio} até {data_fim}")
        if resultados_empresas["2"]["tem_titulos"]:
            logger.info("      ✅ Resultado: Títulos encontrados e arquivo gerado!")
        else:
            logger.info("      ⚠️  Resultado: Sem títulos para processamento")
        logger.info("")
        logger.info("   🏢 EMPRESA 20:")
        logger.info("      ✓ Troca de empresa realizada")
        logger.info("      ✓ Filtros: Conta 14, Layout 12")
        logger.info(f"      ✓ Período: {data_inicio} até {data_fim}")
        if resultados_empresas["20"]["tem_titulos"]:
            logger.info("      ✅ Resultado: Títulos encontrados e arquivo gerado!")
        else:
            logger.info("      ⚠️  Resultado: Sem títulos para processamento")
        logger.info("")
        
        # Resumo de arquivos gerados
        total_arquivos = sum(1 for emp in resultados_empresas.values() if emp["arquivo_gerado"])
        if total_arquivos > 0:
            logger.info(f"   📁 Total de arquivos gerados: {total_arquivos}")
            logger.info("")
        logger.info("   ✓ Sistema fechado corretamente")
        logger.info("")
        
        return True
        
    except KeyboardInterrupt:
        logger.info("")
        logger.warning("⚠️  Execução interrompida pelo usuário (Ctrl+C)")
        return False
        
    except Exception as e:
        logger.error("")
        logger.error("=" * 70)
        logger.error("❌ ERRO DURANTE A EXECUÇÃO")
        logger.error("=" * 70)
        logger.error(f"Erro: {str(e)}")
        logger.error("")
        logger.error("Dicas:")
        logger.error("  • Verifique se o Bimer está aberto")
        logger.error("  • Confirme se as coordenadas estão corretas")
        logger.error("  • Certifique-se de que a tela de login está visível")
        logger.error("")
        return False

# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    logger.info("")
    logger.info("╔" + "═" * 68 + "╗")
    logger.info("║" + " " * 15 + "RPA BIMER - TESTE DE LOGIN (CONFIG)" + " " * 16 + "║")
    logger.info("╚" + "═" * 68 + "╝")
    logger.info("")
    
    # Exibir informações de período de busca
    data_inicio = obter_data_inicio_busca()
    data_fim = obter_data_fim_busca()
    hoje = datetime.now()
    eh_util = eh_dia_util(hoje)
    
    logger.info("📅 INFORMAÇÕES DE DATA:")
    logger.info(f"   • Hoje: {obter_data_atual()} ({'Dia útil' if eh_util else 'Fim de semana/Feriado'})")
    logger.info(f"   • Período de busca: {data_inicio} até {data_fim}")
    if data_inicio != data_fim:
        logger.info(f"   ⚠️  Buscando múltiplos dias (incluindo dias não úteis anteriores)")
    logger.info("")
    
    sucesso = executar_login_bimer()
    
    if sucesso:
        logger.info("✅ Script finalizado com sucesso!")
    else:
        logger.info("⚠️  Script finalizado com erros")
    
    logger.info("")
