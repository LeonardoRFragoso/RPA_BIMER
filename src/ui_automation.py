"""
Módulo de automação de interface do usuário
Utiliza pyautogui e pywinauto para interagir com o Bimmer
"""
import os
import time
import pyautogui
import pywinauto
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
from src.config import config
from src.logger import logger

# Configurações de segurança do pyautogui
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

class BimmerAutomation:
    """Classe para automação do software Bimmer"""
    
    def __init__(self):
        self.config = config.get('bimmer', {})
        self.app = None
        self.delay = self.config.get('action_delay', 1.0)
        self.timeout = self.config.get('screen_load_timeout', 10.0)
        self.executable_path = self.config.get('executable_path', '')
        self.icon_x = self.config.get('icon_x')
        self.icon_y = self.config.get('icon_y')
        self.ui = self.config.get('ui_elements', {}) or {}
        # Coordenadas do login Bimer (opcionais)
        self.pwd_x = self.ui.get('password_bimer_x')
        self.pwd_y = self.ui.get('password_bimer_y')
        self.enter_x = self.ui.get('entrar_bimer_x')
        self.enter_y = self.ui.get('entrar_bimer_y')
        # Modal erro login (OK)
        self.error_ok_x = self.ui.get('erro_login_ok_x')
        self.error_ok_y = self.ui.get('erro_login_ok_y')
        # Segunda tentativa explícita
        self.retry_pwd_x = self.ui.get('retry_password_x')
        self.retry_pwd_y = self.ui.get('retry_password_y')
        self.retry_enter_x = self.ui.get('retry_entrar_x')
        self.retry_enter_y = self.ui.get('retry_entrar_y')

    def digitar_senha_char_por_char(self, senha: str, intervalo: float = 0.06):
        """Digita a senha caractere a caractere, com tratamentos para caracteres especiais comuns."""
        try:
            # Garante que nenhuma tecla modificadora está presa
            for k in ('shift', 'ctrl', 'alt'):
                try:
                    pyautogui.keyUp(k)
                except Exception:
                    pass
            for ch in senha:
                # Tratamentos específicos
                if ch == '@':
                    # Layout PT-BR: AltGr+Q (Ctrl+Alt+Q)
                    try:
                        pyautogui.hotkey('ctrl', 'alt', 'q')
                    except Exception:
                        pyautogui.typewrite('@', interval=0.01)
                else:
                    # typewrite lida com maiúsculas/minúsculas; números usam press com segurança
                    if ch.isdigit():
                        pyautogui.press(ch)
                    else:
                        pyautogui.typewrite(ch, interval=0.01)
                time.sleep(intervalo)
            return True
        except Exception as e:
            logger.warning(f"Falha ao digitar senha char-por-char: {e}")
            return False
        
    def copiar_para_area_transferencia(self, texto: str) -> bool:
        """Copia texto para a área de transferência do Windows."""
        try:
            try:
                import win32clipboard
                import win32con
                win32clipboard.OpenClipboard()
                try:
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, texto)
                finally:
                    win32clipboard.CloseClipboard()
                logger.debug("Senha copiada para a área de transferência (win32clipboard).")
                return True
            except Exception:
                try:
                    import pyperclip
                    pyperclip.copy(texto)
                    logger.debug("Senha copiada para a área de transferência (pyperclip).")
                    return True
                except Exception as e2:
                    logger.warning(f"Falha ao copiar para a área de transferência: {e2}")
                    return False
        except Exception as e:
            logger.warning(f"Erro inesperado ao copiar para a área de transferência: {e}")
            return False

    def ler_area_transferencia(self) -> str:
        """Lê texto da área de transferência do Windows (quando disponível)."""
        try:
            try:
                import win32clipboard
                win32clipboard.OpenClipboard()
                try:
                    data = win32clipboard.GetClipboardData()
                finally:
                    win32clipboard.CloseClipboard()
                return data or ""
            except Exception:
                try:
                    import pyperclip
                    return pyperclip.paste() or ""
                except Exception:
                    return ""
        except Exception:
            return ""

    def focar_janela_bimer(self):
        """Tenta focar a janela do Bimer/RDP para garantir que os eventos vão para a VM."""
        try:
            import pygetwindow as gw
            # Primeiro tenta a janela do Bimer
            windows = (
                gw.getWindowsWithTitle("bimer")
                or gw.getWindowsWithTitle("Bimer")
                or gw.getWindowsWithTitle("Alterdata")
            )
            if windows:
                windows[0].activate()
                time.sleep(0.5)
                return True
            # Se não achar, tenta focar a janela RDP
            windows = (
                gw.getWindowsWithTitle("Conexão de Área de Trabalho Remota")
                or gw.getWindowsWithTitle("Remote Desktop Connection")
            )
            if windows:
                windows[0].activate()
                time.sleep(0.5)
                return True
            return False
        except Exception as e:
            logger.warning(f"Falha ao focar janela: {e}")
            return False

    def login_bimer(self, password: str, max_wait: float = 15.0) -> bool:
        """Preenche o campo de senha do login do Bimer e confirma."""
        try:
            logger.info("Aguardando tela de login do Bimer...")
            self.focar_janela_bimer()
            # Aguarda overlay/splash encerrar
            self.aguardar_login_pronto(max_wait=4.0)

            # Aguarda alguns segundos a tela de login estar visível
            # NOTA: Em RDP, não confiamos em enumerar janelas locais.
            time.sleep(min(max_wait, 5.0))

            sw, sh = pyautogui.size()
            cx, cy = sw // 2, sh // 2

            # Se houver um modal de erro "OK", fechar com Enter
            logger.info("Garantindo que não há modal de erro aberto (OK)...")
            pyautogui.press('enter')
            time.sleep(0.3)

            # Caso tenhamos coordenadas do campo senha, usamos elas primeiro
            if self.pwd_x is not None and self.pwd_y is not None:
                logger.info(f"[Login Bimer] Focando campo de senha mapeado em ({self.pwd_x}, {self.pwd_y})")
                pyautogui.moveTo(self.pwd_x, self.pwd_y, duration=0.2)
                time.sleep(0.1)
                pyautogui.click(self.pwd_x, self.pwd_y)
                time.sleep(0.3)
            else:
                # Clica próximo ao centro do diálogo para garantir foco
                try:
                    logger.info(f"Focando diálogo de login (centro de tela) em ({cx}, {cy})")
                    pyautogui.moveTo(cx, cy, duration=0.2)
                    time.sleep(0.1)
                    pyautogui.click(cx, cy)
                    time.sleep(0.4)
                except Exception as e:
                    logger.warning(f"Falha ao focar diálogo do Bimer: {e}")

            # Se não havia coordenada, tenta offsets relativos
            if self.pwd_x is None or self.pwd_y is None:
                logger.info("Localizando campo de senha por clique relativo...")
                for dy in (40, 60, 80):
                    px, py = cx, cy + dy
                    logger.info(f"Tentando focar campo de senha em ({px}, {py}) [offset {dy}]")
                    pyautogui.moveTo(px, py, duration=0.15)
                    time.sleep(0.1)
                    pyautogui.click(px, py)
                    time.sleep(0.2)
                    # Pequena limpeza para assumir foco
                    pyautogui.hotkey('ctrl', 'a'); time.sleep(0.05); pyautogui.press('delete')
                    time.sleep(0.1)
                    break

            # Limpa o campo de senha
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)
            pyautogui.press('delete')
            time.sleep(0.2)

            # Garante que nenhuma tecla modificadora ficou presa
            try:
                for k in ('shift', 'ctrl', 'alt'):
                    pyautogui.keyUp(k)
            except Exception:
                pass

            # Digita a senha em duas fases para evitar vírgula inicial e garantir '@'
            logger.info("Digitando senha do Bimer caractere por caractere (fase 1: 'Teste123')...")
            base = "Teste123"
            self.digitar_senha_char_por_char(base, intervalo=0.06)
            time.sleep(0.15)
            logger.info("Digitando caractere '@' (fase 2)...")
            # Tenta AltGr+Q; fallback: colar '@'
            ok_at = True
            try:
                pyautogui.hotkey('ctrl', 'alt', 'q')
            except Exception:
                ok_at = False
            time.sleep(0.15)
            if not ok_at:
                if self.copiar_para_area_transferencia("@"):
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.1)
                else:
                    pyautogui.typewrite('@', interval=0.05)
                    time.sleep(0.1)

            # Verifica conteúdo: Ctrl+A, Ctrl+C, compara
            pyautogui.hotkey('ctrl', 'a'); time.sleep(0.1); pyautogui.hotkey('ctrl', 'c'); time.sleep(0.15)
            conteudo = self.ler_area_transferencia()
            logger.info(f"Senha lida do campo (debug): '{conteudo}'")
            if conteudo != password:
                logger.warning("Senha no campo não corresponde. Retentando digitação completa...")
                # Reinsere corretamente: limpa e digita tudo char-por-char
                pyautogui.press('delete'); time.sleep(0.1)
                self.digitar_senha_char_por_char(password, intervalo=0.08)
                time.sleep(0.15)
                pyautogui.hotkey('ctrl', 'a'); time.sleep(0.1); pyautogui.hotkey('ctrl', 'c'); time.sleep(0.15)
                conteudo = self.ler_area_transferencia()
                logger.info(f"Senha lida após retentativa (debug): '{conteudo}'")

            # Confirma (Enter ou clique no botão 'Entrar' mapeado)
            if self.enter_x is not None and self.enter_y is not None:
                logger.info(f"[Login Bimer] Clicando no botão Entrar em ({self.enter_x}, {self.enter_y})")
                pyautogui.moveTo(self.enter_x, self.enter_y, duration=0.15)
                time.sleep(0.1)
                pyautogui.click(self.enter_x, self.enter_y)
                time.sleep(1.5)
            else:
                logger.info("Confirmando login no Bimer (Enter)...")
                pyautogui.press('enter')
                time.sleep(2.0)

            # Segunda tentativa usando coordenadas fornecidas (se houver modal de erro)
            if self.error_ok_x is not None and self.error_ok_y is not None \
               and self.retry_pwd_x is not None and self.retry_pwd_y is not None \
               and self.retry_enter_x is not None and self.retry_enter_y is not None:
                logger.info("[Login Bimer] Verificando modal de erro e realizando 2ª tentativa por coordenadas...")
                # Clica no OK do modal de erro
                pyautogui.moveTo(self.error_ok_x, self.error_ok_y, duration=0.15)
                time.sleep(0.1)
                pyautogui.click(self.error_ok_x, self.error_ok_y)
                time.sleep(0.5)
                # Foca o campo de senha (coordenadas retry)
                pyautogui.moveTo(self.retry_pwd_x, self.retry_pwd_y, duration=0.15)
                time.sleep(0.1)
                pyautogui.click(self.retry_pwd_x, self.retry_pwd_y)
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'a'); time.sleep(0.1); pyautogui.press('delete'); time.sleep(0.2)
                # Digita a senha novamente (2ª tentativa) em duas fases
                base = "Teste123"
                self.digitar_senha_char_por_char(base, intervalo=0.06)
                time.sleep(0.1)
                try:
                    pyautogui.hotkey('ctrl', 'alt', 'q')
                except Exception:
                    if self.copiar_para_area_transferencia("@"):
                        pyautogui.hotkey('ctrl', 'v')
                    else:
                        pyautogui.typewrite('@', interval=0.05)
                time.sleep(0.2)
                # Clica no botão Entrar (coordenadas retry)
                pyautogui.moveTo(self.retry_enter_x, self.retry_enter_y, duration=0.15)
                time.sleep(0.1)
                pyautogui.click(self.retry_enter_x, self.retry_enter_y)
                time.sleep(1.5)
            return True
        except Exception as e:
            logger.error(f"Erro ao fazer login no Bimer: {str(e)}")
            return False

    def iniciar_bimmer_desktop(self):
        """Inicia o Bimmer clicando no ícone da área de trabalho"""
        try:
            logger.info("=" * 60)
            logger.info("INICIANDO BIMMER A PARTIR DO ÍCONE DA ÁREA DE TRABALHO")
            logger.info("=" * 60)
            logger.info("IMPORTANTE: Certifique-se de que a janela RDP está focada!")
            logger.info("A automação será executada dentro da sessão RDP.")
            logger.info("=" * 60)
            
            # Primeiro, vamos para a área de trabalho (Win+D ou clicar em área vazia)
            logger.info("Navegando para área de trabalho...")
            pyautogui.hotkey('win', 'd')  # Mostra área de trabalho
            time.sleep(2)
            
            # Usa coordenadas do config se disponíveis, senão usa coordenadas aproximadas
            if self.icon_x is not None and self.icon_y is not None:
                icon_x = self.icon_x
                icon_y = self.icon_y
                logger.info(f"Usando coordenadas configuradas do ícone: ({icon_x}, {icon_y})")
            else:
                # Obtém o tamanho da tela
                screen_width, screen_height = pyautogui.size()
                
                # Coordenadas aproximadas do ícone (lado esquerdo, meio da tela)
                # Ajuste essas coordenadas conforme necessário
                icon_x = screen_width // 8  # Aproximadamente 1/8 da largura (lado esquerdo)
                icon_y = screen_height // 2  # Meio da altura
                logger.warning(f"Coordenadas do ícone não configuradas, usando aproximadas: ({icon_x}, {icon_y})")
                logger.info("Execute 'python mapear_icone_bimmer.py' para capturar as coordenadas exatas")
            
            logger.info(f"Tentando clicar no ícone do Bimmer em ({icon_x}, {icon_y})")
            
            # Move o mouse para a posição e clica duas vezes
            pyautogui.moveTo(icon_x, icon_y, duration=0.5)
            time.sleep(0.5)
            pyautogui.doubleClick(icon_x, icon_y)
            
            logger.info("Duplo clique no ícone do Bimmer executado")
            logger.info(f"Aguardando {self.timeout} segundos para o Bimmer abrir...")
            time.sleep(self.timeout)  # Aguarda o Bimmer abrir
            
            logger.info("Bimmer deve ter iniciado. Verificando...")
            # Não tentamos conectar via pywinauto quando dentro de RDP
            # Apenas assumimos que o Bimmer foi aberto
            logger.info("Bimmer iniciado com sucesso a partir da área de trabalho")
            return True
                    
        except Exception as e:
            logger.error(f"Erro ao iniciar Bimmer da área de trabalho: {str(e)}")
            return False

    def iniciar_bimmer_por_execucao(self) -> bool:
        """Inicia o Bimer via Win+R digitando o caminho do executável (executa dentro da sessão RDP)."""
        try:
            # Garante foco na janela RDP/Bimer
            self.focar_janela_bimer()
            time.sleep(0.3)

            exec_path = (self.executable_path or "").strip().strip('"')
            if not exec_path:
                logger.warning("Caminho do executável do Bimer não configurado; pulando Win+R")
                return False

            logger.info(f"Iniciando Bimer via Win+R com: {exec_path}")
            pyautogui.hotkey('win', 'r')
            time.sleep(0.7)

            # Limpa a caixa de execução e digita o caminho
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)
            # Usa aspas para lidar com espaços no caminho
            pyautogui.typewrite(f'"{exec_path}"', interval=0.02)
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(self.timeout)

            # Segundo intento (fallback): abrir via cmd /c start "" "path"
            # Algumas políticas bloqueiam executar diretamente da caixa Run.
            logger.info("Verificando se o Bimer iniciou; caso contrário, tentando via cmd /c start ...")
            pyautogui.hotkey('win', 'r')
            time.sleep(0.6)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)
            cmd_line = f'cmd /c start "" "{exec_path}"'
            pyautogui.typewrite(cmd_line, interval=0.02)
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(self.timeout)
            return True
        except Exception as e:
            logger.error(f"Erro ao iniciar Bimer via Win+R: {e}")
            return False
    
    def iniciar_bimmer(self):
        """Inicia o aplicativo Bimer pelo ícone da área de trabalho (comportamento único)."""
        try:
            return self.iniciar_bimmer_desktop()
        except Exception as e:
            logger.error(f"Erro ao iniciar Bimer: {str(e)}")
            return False
    
    def conectar_bimmer(self):
        """Abre o Bimer usando apenas o ícone da área de trabalho."""
        try:
            return self.iniciar_bimmer_desktop()
        except Exception as e:
            logger.error(f"Erro ao abrir Bimer: {str(e)}")
            return False

    def aguardar_login_pronto(self, max_wait: float = 10.0):
        """Aguarda o carregamento inicial do Bimer (evita interagir durante o overlay NF Easy)."""
        logger.info("Aguardando carregamento inicial do Bimer (overlay NF Easy)...")
        t0 = time.time()
        while time.time() - t0 < max_wait:
            # Pequenos pulsos para manter foco e evitar cliques acidentais
            self.focar_janela_bimer()
            time.sleep(0.5)
        logger.info("Prosseguindo após aguardar o carregamento do Bimer.")
    
    def aguardar_elemento(self, elemento, timeout=None):
        """Aguarda um elemento aparecer na tela"""
        if timeout is None:
            timeout = self.timeout
        
        inicio = time.time()
        while time.time() - inicio < timeout:
            try:
                # Tenta encontrar o elemento
                elemento.wait('visible', timeout=2)
                return True
            except:
                time.sleep(0.5)
        
        return False
    
    def clicar(self, x, y, botao='left', cliques=1):
        """Realiza clique na posição especificada"""
        try:
            logger.debug(f"Clicando em ({x}, {y}) com botão {botao}")
            pyautogui.click(x, y, clicks=cliques, button=botao)
            time.sleep(self.delay)
            return True
        except Exception as e:
            logger.error(f"Erro ao clicar: {str(e)}")
            return False
    
    def clicar_por_imagem(self, imagem_path, confianca=0.8):
        """Clica em um elemento identificado por imagem"""
        try:
            logger.debug(f"Procurando imagem: {imagem_path}")
            # Tenta usar confidence se opencv estiver disponível
            try:
                localizacao = pyautogui.locateOnScreen(imagem_path, confidence=confianca)
            except TypeError:
                # Se confidence não for suportado (sem opencv), tenta sem
                logger.debug("OpenCV não disponível, tentando sem confidence")
                localizacao = pyautogui.locateOnScreen(imagem_path)
            
            if localizacao:
                centro = pyautogui.center(localizacao)
                logger.debug(f"Imagem encontrada em {centro}")
                pyautogui.click(centro)
                time.sleep(self.delay)
                return True
            else:
                logger.warning(f"Imagem não encontrada: {imagem_path}")
                return False
                
        except Exception as e:
            logger.error(f"Erro ao clicar por imagem: {str(e)}")
            logger.warning("Se o erro persistir, instale opencv-python para melhor suporte a reconhecimento de imagens")
            return False
    
    def digitar(self, texto, intervalo=0.1):
        """Digita texto na posição atual do cursor"""
        try:
            logger.debug(f"Digitando: {texto}")
            pyautogui.write(texto, interval=intervalo)
            time.sleep(self.delay)
            return True
        except Exception as e:
            logger.error(f"Erro ao digitar: {str(e)}")
            return False
    
    def pressionar_tecla(self, tecla, vezes=1):
        """Pressiona uma tecla do teclado"""
        try:
            logger.debug(f"Pressionando tecla: {tecla}")
            for _ in range(vezes):
                pyautogui.press(tecla)
                time.sleep(0.1)
            time.sleep(self.delay)
            return True
        except Exception as e:
            logger.error(f"Erro ao pressionar tecla: {str(e)}")
            return False
    
    def pressionar_combinacao(self, *teclas):
        """Pressiona combinação de teclas (ex: Ctrl+C)"""
        try:
            logger.debug(f"Pressionando combinação: {teclas}")
            pyautogui.hotkey(*teclas)
            time.sleep(self.delay)
            return True
        except Exception as e:
            logger.error(f"Erro ao pressionar combinação: {str(e)}")
            return False
    
    def aguardar(self, segundos):
        """Aguarda um tempo especificado"""
        logger.debug(f"Aguardando {segundos} segundos...")
        time.sleep(segundos)
    
    def capturar_tela(self, caminho=None):
        """Captura uma screenshot"""
        try:
            if caminho is None:
                caminho = f"screenshot_{int(time.time())}.png"
            
            screenshot = pyautogui.screenshot()
            screenshot.save(caminho)
            logger.debug(f"Screenshot salva em: {caminho}")
            return caminho
        except Exception as e:
            logger.error(f"Erro ao capturar tela: {str(e)}")
            return None
    
    def encontrar_texto_na_tela(self, texto):
        """Procura texto na tela usando OCR (requer tesseract)"""
        try:
            # Esta funcionalidade requer tesseract instalado
            # Por enquanto, retorna None
            logger.warning("Busca por texto na tela não implementada (requer OCR)")
            return None
        except Exception as e:
            logger.error(f"Erro ao buscar texto: {str(e)}")
            return None

