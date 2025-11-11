"""
Script principal para executar o bot do Bimmer
Primeiro conecta à VM via RDP, depois executa a automação dentro da VM
"""
import sys
import time
from pathlib import Path

# Adiciona o diretório raiz ao path
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

from src.bot import BimmerBot
from src.rdp_connection import RDPConnection
from src.rdp_auto_fill import RDPAutoFill
from src.logger import logger
import pyautogui
import pygetwindow as gw

def aguardar_janela_rdp(timeout=30):
    """Aguarda a janela RDP aparecer e foca nela"""
    try:
        logger.info("Aguardando janela RDP aparecer...")
        inicio = time.time()
        
        while time.time() - inicio < timeout:
            try:
                # Procura por janelas RDP
                windows = gw.getWindowsWithTitle("Conexão de Área de Trabalho Remota")
                if not windows:
                    windows = gw.getWindowsWithTitle("Remote Desktop Connection")
                if not windows:
                    # Tenta encontrar pela porta ou IP
                    windows = [w for w in gw.getAllWindows() if "144.22.232.212" in w.title or "50491" in w.title]
                
                if windows:
                    window = windows[0]
                    if window.visible:
                        logger.info("Janela RDP encontrada!")
                        window.activate()
                        time.sleep(2)  # Aguarda a janela ficar ativa
                        return True
            except Exception as e:
                logger.debug(f"Erro ao procurar janela RDP: {e}")
            
            time.sleep(1)
        
        logger.warning("Janela RDP não encontrada no tempo esperado")
        return False
        
    except Exception as e:
        logger.error(f"Erro ao aguardar janela RDP: {str(e)}")
        return False

def main():
    """Função principal"""
    try:
        logger.info("=" * 50)
        logger.info("RPA Bimmer - Iniciando execução")
        logger.info("=" * 50)
        
        # Passo 1: Conectar à VM via RDP
        logger.info("Passo 1: Conectando à VM via RDP...")
        rdp = RDPConnection()
        
        if not rdp.conectar_com_senha():
            logger.warning("Tentando método alternativo de conexão...")
            if not rdp.conectar_rdp(usar_arquivo=True):
                logger.error("Não foi possível conectar à VM via RDP")
                logger.info("Por favor, conecte manualmente à VM e execute o bot novamente")
                sys.exit(1)
        
        logger.info("Conexão RDP iniciada, aguardando janela abrir...")
        time.sleep(2)  # Aguarda a janela RDP abrir
        
        # Passo 2: Preencher credenciais do Windows de forma objetiva (apenas senha + OK)
        logger.info("Passo 2: Preenchendo credenciais do Windows (senha + OK)...")
        try:
            auto_fill = RDPAutoFill()
            if auto_fill.preencher_automaticamente(timeout=15):
                logger.info("✓ Credenciais do Windows preenchidas.")
            else:
                logger.warning("Tela de credenciais não encontrada ou já autenticada.")
        except Exception as e:
            logger.warning(f"Falha ao preencher credenciais do Windows: {e}")
            logger.info("Se a tela de credenciais aparecer, insira manualmente e clique em OK.")
        
        # Aguarda um pouco para a conexão ser estabelecida
        logger.info("Aguardando conexão RDP ser estabelecida...")
        time.sleep(5)  # Aguarda mais tempo para a conexão ser estabelecida
        
        # Verifica se a conexão foi estabelecida verificando se a janela RDP está aberta
        logger.info("Verificando se a conexão RDP foi estabelecida...")
        if not aguardar_janela_rdp(timeout=10):
            logger.warning("Janela RDP não encontrada - conexão pode não ter sido estabelecida")
            logger.info("Por favor, verifique se a conexão RDP foi estabelecida corretamente")
            logger.info("Se a senha estava incorreta, corrija no config.yaml e tente novamente")
        
        # Passo 3: Aguardar e focar na janela RDP
        logger.info("Passo 3: Aguardando janela RDP...")
        if not aguardar_janela_rdp(timeout=30):
            logger.warning("Janela RDP não encontrada automaticamente")
            logger.info("Por favor, certifique-se de que a janela RDP está aberta e ativa")
            logger.info("Aguardando 10 segundos para você focar na janela RDP...")
            time.sleep(10)
        
        # Passo 4: Garantir que a janela RDP está focada
        logger.info("Passo 4: Garantindo que a janela RDP está focada...")
        try:
            # Tenta focar na janela RDP novamente
            windows = gw.getWindowsWithTitle("Conexão de Área de Trabalho Remota")
            if not windows:
                windows = gw.getWindowsWithTitle("Remote Desktop Connection")
            if not windows:
                windows = [w for w in gw.getAllWindows() if "144.22.232.212" in w.title or "50491" in w.title]
            
            if windows:
                window = windows[0]
                window.activate()
                logger.info("Janela RDP focada com sucesso")
                time.sleep(2)
            else:
                logger.warning("Não foi possível encontrar a janela RDP para focar")
        except Exception as e:
            logger.warning(f"Erro ao focar janela RDP: {e}")
        
        # Passo 5: Garantir que a janela RDP está focada antes de executar o bot
        logger.info("Passo 5: Garantindo que a janela RDP está focada...")
        logger.info("IMPORTANTE: A janela RDP DEVE estar focada para a automação funcionar!")
        logger.info("A automação será executada dentro da sessão RDP.")
        
        # Foca na janela RDP novamente antes de executar o bot
        try:
            windows = gw.getWindowsWithTitle("Conexão de Área de Trabalho Remota")
            if not windows:
                windows = gw.getWindowsWithTitle("Remote Desktop Connection")
            if not windows:
                windows = [w for w in gw.getAllWindows() if "144.22.232.212" in w.title or "50491" in w.title]
            
            if windows:
                window = windows[0]
                window.activate()
                logger.info("Janela RDP focada com sucesso antes de executar o bot")
                time.sleep(3)  # Aguarda a janela ficar completamente ativa
            else:
                logger.warning("Não foi possível encontrar a janela RDP para focar")
                logger.info("Aguardando 5 segundos para você focar manualmente na janela RDP...")
                time.sleep(5)
        except Exception as e:
            logger.warning(f"Erro ao focar janela RDP: {e}")
            logger.info("Aguardando 5 segundos para você focar manualmente na janela RDP...")
            time.sleep(5)
        
        # Pausa adicional para aguardar NF Easy abrir e fechar antes de iniciar o Bimer
        logger.info("=" * 60)
        logger.info("Aguardando 40 segundos para estabilizar a sessão (NF Easy abrir/fechar)...")
        logger.info("=" * 60)
        time.sleep(40)
        
        # Passo 6: Executar o bot dentro da sessão RDP
        logger.info("=" * 60)
        logger.info("Passo 6: Executando bot dentro da VM...")
        logger.info("=" * 60)
        logger.info("IMPORTANTE: Certifique-se de que a janela RDP está ativa e focada!")
        logger.info("A automação será executada dentro da sessão RDP.")
        logger.info("Aguardando 2 segundos antes de iniciar a automação...")
        time.sleep(2)
        
        # Cria e executa o bot
        # Nota: Quando a janela RDP está focada, o pyautogui executará dentro da sessão remota
        bot = BimmerBot()
        sucesso = bot.executar()
        
        if sucesso:
            logger.info("Execução concluída com sucesso")
            sys.exit(0)
        else:
            logger.error("Execução concluída com erros")
            sys.exit(1)
            
    except KeyboardInterrupt:
        logger.info("Execução interrompida pelo usuário")
        sys.exit(0)
        
    except Exception as e:
        logger.error(f"Erro fatal: {str(e)}", exc_info=True)
        sys.exit(1)

if __name__ == '__main__':
    main()

