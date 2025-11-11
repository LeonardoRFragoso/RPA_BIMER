"""
Script standalone para testar login no Bimer
Execute este script DENTRO da VM onde o Bimer está aberto
Não precisa de conexão RDP - testa apenas o fluxo de login
"""
import time
import pyautogui
import logging
from pathlib import Path

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
BOTAO_ENTRAR_X = 511
BOTAO_ENTRAR_Y = 682

# Coordenadas pós-login (se mapeadas)
FECHAR_MODAL_X = None  # Configure se necessário
FECHAR_MODAL_Y = None

# Sequência de cliques pós-login
CLIQUES_POS_LOGIN = [
    # ("nome_elemento", x, y, "ação_opcional")
    # Exemplo: ("menu_financeiro", 100, 200, None)
]

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

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
            aguardar(2.0, "Aguardando login processar...")
        else:
            logger.info(f"→ Clicando em Entrar em ({BOTAO_ENTRAR_X}, {BOTAO_ENTRAR_Y})")
            pyautogui.click(BOTAO_ENTRAR_X, BOTAO_ENTRAR_Y)
            aguardar(2.0, "Aguardando login processar...")
            logger.info("✓ Botão Entrar clicado")
        
        # ========================================
        # PASSO 4: Pós-login (modal e cliques)
        # ========================================
        logger.info("")
        logger.info("=" * 70)
        logger.info("[PASSO 4/4] AÇÕES PÓS-LOGIN")
        logger.info("=" * 70)
        
        aguardar(1.0, "Aguardando sistema processar...")
        
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
                logger.info(f"  [{i}/{len(CLIQUES_POS_LOGIN)}] {nome} em ({x}, {y})")
                pyautogui.click(x, y)
                aguardar(0.4)
                
                # Se houver ação adicional (ex: digitar)
                if acao:
                    logger.info(f"       → Executando: {acao}")
                    if isinstance(acao, str):
                        pyautogui.write(acao, interval=0.05)
                        aguardar(0.3)
            logger.info("✓ Sequência de cliques concluída")
        else:
            logger.info("ℹ️  Nenhum clique pós-login configurado")
        
        # ========================================
        # CONCLUSÃO
        # ========================================
        logger.info("")
        logger.info("=" * 70)
        logger.info("✅ TESTE DE LOGIN CONCLUÍDO COM SUCESSO!")
        logger.info("=" * 70)
        logger.info("")
        logger.info("📋 Próximos passos:")
        logger.info("   1. Verifique se o login foi bem-sucedido")
        logger.info("   2. Capture coordenadas de novos elementos se necessário")
        logger.info("   3. Adicione cliques pós-login em CLIQUES_POS_LOGIN")
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
