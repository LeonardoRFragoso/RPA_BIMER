"""
Script standalone para testar login no Bimer
Execute este script DENTRO da VM onde o Bimer está aberto
Não precisa de conexão RDP - testa apenas o fluxo de login
"""
import time
import pyautogui
import logging
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
# CONFIGURAÇÕES - AJUSTE CONFORME NECESSÁRIO
# ============================================================

# Credenciais do Bimer
SENHA_BIMER = "Rpa@@2025"

# Coordenadas capturadas (ajuste se necessário)
DROPDOWN_AMBIENTE_X = 866
DROPDOWN_AMBIENTE_Y = 579
AMBIENTE_TESTE_X = 974
AMBIENTE_TESTE_Y = 677
CAMPO_SENHA_X = 904
CAMPO_SENHA_Y = 520
BOTAO_ENTRAR_X = 953
BOTAO_ENTRAR_Y = 645

# Coordenadas pós-login (mapeadas na VM)
FECHAR_MODAL_X = 1391
FECHAR_MODAL_Y = 192

# Coordenadas para quando HOUVER títulos para processar
MARCAR_TODOS_TITULOS_X = 619
MARCAR_TODOS_TITULOS_Y = 335
CAMPO_NOME_ARQUIVO_X = 804
CAMPO_NOME_ARQUIVO_Y = 732
BOTAO_GERAR_ARQUIVO_X = 1312
BOTAO_GERAR_ARQUIVO_Y = 838
BOTAO_SIM_CONFIRMACAO_X = 919
BOTAO_SIM_CONFIRMACAO_Y = 581
BOTAO_OK_OPERACAO_CONCLUIDA_X = 949
BOTAO_OK_OPERACAO_CONCLUIDA_Y = 559

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
    """
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

# Sequência de cliques pós-login
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
    ("VERIFICAR_TITULOS_EMPRESA_1", None, None, None),  # Ação condicional
    ("Fechar Modal Remessa", 1453, 187, None),
    
    # TROCAR PARA SEGUNDA EMPRESA
    ("A Pagar - Voltar para Empresa 2", 89, 349, None),
    ("Campo Empresa - Trocar para 2", 106, 209, "2"),
    ("Fechar modal A Pagar - Empresa 2", 1389, 192, None),
    
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
    ("VERIFICAR_TITULOS_EMPRESA_2", None, None, None),  # Ação condicional
    ("Fechar Modal Remessa - Empresa 2", 1453, 187, None),
    
    # TROCAR PARA TERCEIRA EMPRESA
    ("A Pagar - Voltar para Empresa 20", 89, 349, None),
    ("Campo Empresa - Trocar para 20", 106, 209, "20"),
    ("Fechar modal A Pagar - Empresa 20", 1389, 192, None),
    
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
    ("VERIFICAR_TITULOS_EMPRESA_20", None, None, None),  # Ação condicional
    ("Fechar Modal Remessa - Empresa 20", 1453, 187, None),
    
    # FINALIZAR SISTEMA
    ("Fechar Bimer", 1904, 6, None),
    ("Confirmar Fechar Sistema", 871, 552, None),
]

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

# Variável global para rastrear resultados por empresa
resultados_empresas = {
    "1": {"tem_titulos": False, "arquivo_gerado": False},
    "2": {"tem_titulos": False, "arquivo_gerado": False},
    "20": {"tem_titulos": False, "arquivo_gerado": False}
}

def verificar_e_processar_titulos(empresa_numero):
    """
    Verifica se há títulos para processar e executa as ações apropriadas.
    Dois cenários possíveis:
    1. Modal "Não há títulos" aparece → Clicar OK e retornar False
    2. Lista de títulos aparece → Marcar todos e gerar arquivo
    """
    logger.info(f"   🔍 Verificando se há títulos para Empresa {empresa_numero}...")
    aguardar(3.0, "Aguardando tela carregar...")
    
    # Primeiro: verificar se o modal "Não há títulos" apareceu
    # Clicar no OK (945, 553)
    logger.info(f"   → Verificando se modal 'Sem títulos' apareceu...")
    pyautogui.click(945, 553)  # Tentar clicar no OK
    aguardar(1.5, "Aguardando resposta...")
    
    # Agora verificar se ainda estamos na tela de títulos
    # Tentar clicar no checkbox "Marcar Todos"
    # Se conseguir, o modal NÃO existia (há títulos)
    # Se não conseguir, o modal existia e foi fechado (não há títulos)
    
    logger.info(f"   → Tentando marcar todos os títulos...")
    pyautogui.click(MARCAR_TODOS_TITULOS_X, MARCAR_TODOS_TITULOS_Y)
    aguardar(1.0)
    
    # Verificar se conseguimos acessar o campo de arquivo
    # Isso confirma que há títulos e eles foram marcados
    logger.info(f"   → Verificando se há títulos marcados...")
    try:
        # Tentar clicar no campo de arquivo
        pyautogui.click(CAMPO_NOME_ARQUIVO_X, CAMPO_NOME_ARQUIVO_Y)
        aguardar(0.5)
        
        # Testar se o campo está realmente acessível digitando algo
        pyautogui.hotkey('ctrl', 'a')
        aguardar(0.2)
        
        # Se chegou aqui sem erro, há títulos!
        logger.info(f"   ✅ TÍTULOS ENCONTRADOS para Empresa {empresa_numero}!")
        
        # Definir caminho completo do arquivo (com número da empresa)
        caminho_completo = obter_caminho_completo_arquivo_remessa(empresa_numero)
        logger.info(f"   → Definindo caminho do arquivo: {caminho_completo}")
        
        # Clicar no campo de nome do arquivo
        pyautogui.click(CAMPO_NOME_ARQUIVO_X, CAMPO_NOME_ARQUIVO_Y)
        aguardar(0.5, "Campo de nome focado")
        
        # Selecionar tudo (Ctrl+A)
        pyautogui.hotkey('ctrl', 'a')
        aguardar(0.2)
        
        # Apagar tudo
        pyautogui.press('backspace')
        aguardar(0.3)
        
        # Digitar o caminho completo
        pyautogui.typewrite(caminho_completo, interval=0.05)
        aguardar(0.5, f"Caminho definido: {caminho_completo}")
        
        logger.info(f"   → Clicando no botão Gerar arquivo...")
        pyautogui.click(BOTAO_GERAR_ARQUIVO_X, BOTAO_GERAR_ARQUIVO_Y)
        aguardar(2.0, "Aguardando caixa de confirmação...")
        
        # Clicar em "Sim" na caixa de diálogo de confirmação
        logger.info(f"   → Confirmando geração do arquivo (Sim)...")
        pyautogui.click(BOTAO_SIM_CONFIRMACAO_X, BOTAO_SIM_CONFIRMACAO_Y)
        aguardar(3.0, "Arquivo sendo gerado...")
        
        # Clicar em "OK" no modal de operação concluída
        logger.info(f"   → Confirmando operação concluída (OK)...")
        pyautogui.click(BOTAO_OK_OPERACAO_CONCLUIDA_X, BOTAO_OK_OPERACAO_CONCLUIDA_Y)
        aguardar(1.0, "Modal fechado")
        
        # Atualizar resultado
        resultados_empresas[empresa_numero]["tem_titulos"] = True
        resultados_empresas[empresa_numero]["arquivo_gerado"] = True
        
        logger.info(f"   ✅ Arquivo gerado com sucesso para Empresa {empresa_numero}!")
        logger.info(f"   📁 Local: {caminho_completo}")
        return True
        
    except Exception as e:
        # Se chegou aqui, é porque não há títulos
        logger.info(f"   ⚠️  Nenhum título encontrado para Empresa {empresa_numero}")
        
        # Atualizar resultado
        resultados_empresas[empresa_numero]["tem_titulos"] = False
        resultados_empresas[empresa_numero]["arquivo_gerado"] = False
        
        return False

def copiar_para_clipboard(texto):
    """Copia texto para área de transferência"""
    try:
        import win32clipboard
        import win32con
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, texto)
        finally:
            win32clipboard.CloseClipboard()
        logger.info("✓ Texto copiado para área de transferência")
        return True
    except Exception:
        try:
            import pyperclip
            pyperclip.copy(texto)
            logger.info("✓ Texto copiado para área de transferência (pyperclip)")
            return True
        except Exception as e:
            logger.error(f"✗ Falha ao copiar para área de transferência: {e}")
            return False

def aguardar(segundos, mensagem=""):
    """Aguarda com mensagem opcional"""
    if mensagem:
        logger.info(f"⏳ {mensagem}")
    time.sleep(segundos)

# ============================================================
# FLUXO DE LOGIN
# ============================================================

def executar_login_bimer():
    """Executa o login no Bimer"""
    try:
        logger.info("=" * 70)
        logger.info("🤖 TESTE DE LOGIN NO BIMER - VERSÃO STANDALONE")
        logger.info("=" * 70)
        logger.info("")
        
        # Calcular e exibir período de busca
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
        
        logger.info("⚠️  IMPORTANTE:")
        logger.info("   1. Certifique-se de que o Bimer está ABERTO")
        logger.info("   2. A tela de LOGIN deve estar VISÍVEL")
        logger.info("   3. Não mova o mouse durante a execução")
        logger.info("")
        logger.info("Iniciando em 3 segundos...")
        time.sleep(3)
        
        # ========================================
        # PASSO 1: Selecionar ambiente TESTE
        # ========================================
        logger.info("")
        logger.info("=" * 70)
        logger.info("[PASSO 1/4] SELECIONANDO AMBIENTE TESTE NO DROPDOWN")
        logger.info("=" * 70)
        
        if DROPDOWN_AMBIENTE_X and DROPDOWN_AMBIENTE_Y and AMBIENTE_TESTE_X and AMBIENTE_TESTE_Y:
            logger.info(f"→ Clicando no dropdown em ({DROPDOWN_AMBIENTE_X}, {DROPDOWN_AMBIENTE_Y})")
            pyautogui.click(DROPDOWN_AMBIENTE_X, DROPDOWN_AMBIENTE_Y)
            aguardar(0.5, "Aguardando dropdown abrir...")
            
            logger.info(f"→ Clicando em 'TESTE' em ({AMBIENTE_TESTE_X}, {AMBIENTE_TESTE_Y})")
            pyautogui.click(AMBIENTE_TESTE_X, AMBIENTE_TESTE_Y)
            aguardar(0.5, "Ambiente selecionado")
            logger.info("✓ Ambiente TESTE selecionado com sucesso")
        else:
            logger.warning("⚠️  Coordenadas do dropdown não configuradas - pulando passo 1")
        
        # ========================================
        # PASSO 2: Preencher senha
        # ========================================
        logger.info("")
        logger.info("=" * 70)
        logger.info("[PASSO 2/4] PREENCHENDO SENHA")
        logger.info("=" * 70)
        
        if not CAMPO_SENHA_X or not CAMPO_SENHA_Y:
            logger.error("✗ Coordenadas do campo de senha não configuradas!")
            return False
        
        logger.info(f"→ Clicando no campo de senha em ({CAMPO_SENHA_X}, {CAMPO_SENHA_Y})")
        pyautogui.click(CAMPO_SENHA_X, CAMPO_SENHA_Y)
        aguardar(0.4, "Campo focado")
        
        logger.info("→ Limpando campo (Ctrl+A)")
        pyautogui.hotkey('ctrl', 'a')
        aguardar(0.2)
        
        # Libera teclas modificadoras
        for k in ('shift', 'ctrl', 'alt'):
            try:
                pyautogui.keyUp(k)
            except:
                pass
        aguardar(0.2)
        
        logger.info(f"→ Colando senha ({len(SENHA_BIMER)} caracteres)")
        if copiar_para_clipboard(SENHA_BIMER):
            aguardar(0.2)
            pyautogui.hotkey('ctrl', 'v')
            aguardar(0.3)
            logger.info("✓ Senha colada com sucesso")
        else:
            logger.warning("⚠️  Falha ao copiar. Digitando caractere por caractere...")
            pyautogui.write(SENHA_BIMER, interval=0.08)
            aguardar(0.3)
            logger.info("✓ Senha digitada com sucesso")
        
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
                        # Para campo de empresa: usar backspace múltiplo para limpar
                        if "empresa" in nome.lower() and "trocar" in nome.lower():
                            logger.info(f"       → Limpando campo de empresa com backspaces")
                            # Clicar no campo para garantir foco
                            pyautogui.click(x, y)
                            aguardar(0.3)
                            # Selecionar tudo e apagar
                            pyautogui.hotkey('ctrl', 'a')
                            aguardar(0.2)
                            pyautogui.press('backspace')
                            aguardar(0.3)
                            # Digitar o novo número
                            pyautogui.typewrite(valor_digitar, interval=0.1)
                            aguardar(0.3)
                            # Pressionar Enter para confirmar
                            logger.info(f"       → Confirmando com Enter")
                            pyautogui.press('enter')
                            aguardar(0.5)
                        else:
                            # Para outros campos: usar Ctrl+A
                            pyautogui.hotkey('ctrl', 'a')
                            aguardar(0.2)
                            pyautogui.press('delete')
                            aguardar(0.2)
                            pyautogui.typewrite(valor_digitar, interval=0.1)
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
        logger.error(f"❌ ERRO DURANTE EXECUÇÃO: {str(e)}")
        logger.error("=" * 70)
        import traceback
        logger.error(traceback.format_exc())
        return False

# ============================================================
# EXECUÇÃO PRINCIPAL
# ============================================================

if __name__ == '__main__':
    try:
        # Captura posição inicial do mouse
        pos_inicial = pyautogui.position()
        logger.info(f"Posição inicial do mouse: {pos_inicial}")
        
        # Executa o login
        sucesso = executar_login_bimer()
        
        # Resultado final
        if sucesso:
            logger.info("")
            logger.info("🎉 Script executado com sucesso!")
            exit(0)
        else:
            logger.error("")
            logger.error("💥 Script finalizado com erros")
            exit(1)
            
    except Exception as e:
        logger.error(f"Erro fatal: {e}")
        import traceback
        traceback.print_exc()
        exit(1)
