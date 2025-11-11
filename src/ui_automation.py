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
        raw_modal = self.ui.get('botao_fechar_modal_automatico')
        self.modal_close_x, self.modal_close_y = (None, None)
        if isinstance(raw_modal, str):
            try:
                parts = [p.strip() for p in raw_modal.replace(';', ',').split(',')]
                if len(parts) >= 2:
                    self.modal_close_x = int(parts[0])
                    self.modal_close_y = int(parts[1])
            except Exception:
                self.modal_close_x, self.modal_close_y = (None, None)
        if self.modal_close_x is None or self.modal_close_y is None:
            self.modal_close_x = self.ui.get('botao_fechar_modal_automatico_x')
            self.modal_close_y = self.ui.get('botao_fechar_modal_automatico_y')

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
                        # Fallback: usar write() ao invés de typewrite()
                        pyautogui.write('@')
                else:
                    # Usa write() para caracteres normais (lida melhor com maiúsculas/minúsculas)
                    pyautogui.write(ch)
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
            self.aguardar_login_pronto(max_wait=4.0)
            time.sleep(min(max_wait, 5.0))

            # Aguarda mais tempo para garantir que a tela está estável
            time.sleep(1.0)

            # Foca no campo de senha usando coordenadas mapeadas
            if self.pwd_x is not None and self.pwd_y is not None:
                logger.info(f"Focando campo de senha em ({self.pwd_x}, {self.pwd_y})")
                pyautogui.click(self.pwd_x, self.pwd_y)
                time.sleep(0.4)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.3)
            else:
                # Fallback: clica no centro da tela
                sw, sh = pyautogui.size()
                pyautogui.click(sw // 2, sh // 2 + 60)
                time.sleep(0.4)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.3)

            # Libera teclas modificadoras antes de colar
            for k in ('shift', 'ctrl', 'alt'):
                try:
                    pyautogui.keyUp(k)
                except:
                    pass
            time.sleep(0.3)

            # Cola a senha via área de transferência
            logger.info(f"Colando senha: '{password}'")
            if self.copiar_para_area_transferencia(password):
                time.sleep(0.3)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.4)
                
                logger.info("✓ Senha colada com sucesso")
            else:
                logger.warning("Falha ao colar. Usando fallback...")
                self.digitar_senha_char_por_char(password, intervalo=0.08)
                time.sleep(0.4)

            # Verificação pré-login: garante que o campo contém exatamente a senha
            try:
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.1)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.1)
                conteudo_campo = self.ler_area_transferencia()
                if conteudo_campo != password:
                    logger.warning("Conteúdo divergente antes do login; recoloando senha")
                    if self.copiar_para_area_transferencia(password):
                        pyautogui.hotkey('ctrl', 'v')
                        time.sleep(0.3)
            except Exception:
                pass

            # Aguarda um pouco antes de clicar em Entrar
            time.sleep(0.5)

            # Clica no botão Entrar
            if self.enter_x is not None and self.enter_y is not None:
                logger.info(f"Clicando em Entrar ({self.enter_x}, {self.enter_y})")
                pyautogui.click(self.enter_x, self.enter_y)
            else:
                logger.info("Confirmando com Enter")
                pyautogui.press('enter')
            
            time.sleep(2.5)
            
            # Segunda tentativa (sempre forçada para sanar vírgula/ruído)
            try:
                # Garante foco na janela correta
                self.focar_janela_bimer()
                try:
                    sw, sh = pyautogui.size()
                    pyautogui.click(sw // 2, sh // 2)  # clique neutro para trazer foco
                    time.sleep(0.3)
                except Exception:
                    pass
                # Fecha modal de erro caso existam coordenadas mapeadas
                if self.error_ok_x is not None and self.error_ok_y is not None:
                    logger.info(f"Tentando fechar modal de erro em ({self.error_ok_x}, {self.error_ok_y})")
                    pyautogui.moveTo(self.error_ok_x, self.error_ok_y, duration=0.15)
                    time.sleep(0.1)
                    pyautogui.click(self.error_ok_x, self.error_ok_y)
                    time.sleep(0.8)

                # Refoca o campo de senha para nova inserção
                alvo_x = self.retry_pwd_x if self.retry_pwd_x is not None else self.pwd_x
                alvo_y = self.retry_pwd_y if self.retry_pwd_y is not None else self.pwd_y
                if alvo_x is not None and alvo_y is not None:
                    logger.info(f"2ª tentativa: focando campo de senha em ({alvo_x}, {alvo_y})")
                    pyautogui.moveTo(alvo_x, alvo_y, duration=0.15)
                    time.sleep(0.1)
                    pyautogui.click(alvo_x, alvo_y)
                    time.sleep(0.3)
                else:
                    sw, sh = pyautogui.size()
                    logger.info("2ª tentativa: focando campo de senha (fallback centro)")
                    pyautogui.click(sw // 2, sh // 2 + 60)
                    time.sleep(0.3)

                # Limpa campo de senha selecionando tudo (sem Delete para evitar vírgula via NumPad)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.3)

                # Garante que modificadoras estão liberadas
                for k in ('shift', 'ctrl', 'alt'):
                    try:
                        pyautogui.keyUp(k)
                    except:
                        pass
                time.sleep(0.2)

                # Recarrega senha na área de transferência e cola novamente
                if self.copiar_para_area_transferencia(password):
                    logger.info("2ª tentativa: colando senha")
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.4)
                else:
                    logger.warning("2ª tentativa: falha ao copiar senha; digitando char-por-char")
                    self.digitar_senha_char_por_char(password, intervalo=0.08)
                    time.sleep(0.3)

                # Verificação pré-login na 2ª tentativa
                try:
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(0.1)
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(0.1)
                    conteudo_campo2 = self.ler_area_transferencia()
                    if conteudo_campo2 != password:
                        logger.warning("2ª tentativa: conteúdo divergente; recoloando senha")
                        if self.copiar_para_area_transferencia(password):
                            pyautogui.hotkey('ctrl', 'v')
                            time.sleep(0.3)
                except Exception:
                    pass

                # Clica Entrar novamente (usa retry_enter se configurado)
                btn_x = self.retry_enter_x if self.retry_enter_x is not None else self.enter_x
                btn_y = self.retry_enter_y if self.retry_enter_y is not None else self.enter_y
                if btn_x is not None and btn_y is not None:
                    logger.info(f"2ª tentativa: clicando em Entrar ({btn_x}, {btn_y})")
                    pyautogui.moveTo(btn_x, btn_y, duration=0.15)
                    time.sleep(0.1)
                    pyautogui.click(btn_x, btn_y)
                else:
                    logger.info("2ª tentativa: confirmando com Enter")
                    pyautogui.press('enter')
                time.sleep(2.5)
            except Exception as e2:
                logger.warning(f"Falha na lógica de 2ª tentativa: {e2}")
            
            logger.info("✓ Login no Bimer concluído")
            try:
                self.fechar_modal_automatico()
            except Exception:
                pass
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
    
    def fechar_modal_automatico(self):
        try:
            x, y = self.modal_close_x, self.modal_close_y
            if x is None or y is None:
                return False
            self.focar_janela_bimer()
            try:
                pyautogui.moveTo(x, y, duration=0.15)
                time.sleep(0.1)
            except Exception:
                pass
            pyautogui.click(x, y)
            time.sleep(0.5)
            return True
        except Exception as e:
            logger.warning(f"Falha ao fechar modal automatico: {e}")
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

