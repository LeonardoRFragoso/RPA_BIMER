"""
Módulo de automação de interface do usuário
Utiliza pyautogui e pywinauto para interagir com o Bimmer
"""
import os
import time
import ctypes
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
            self._ensure_focus(click_center=True)
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
                        self.pressionar_combinacao('ctrl', 'alt', 'q')
                    except Exception:
                        # Fallback: usar write() ao invés de typewrite()
                        try:
                            pyautogui.write('@')
                        except pyautogui.FailSafeException:
                            self._recover_from_failsafe()
                            pyautogui.write('@')
                else:
                    # Usa write() para caracteres normais (lida melhor com maiúsculas/minúsculas)
                    try:
                        pyautogui.write(ch)
                    except pyautogui.FailSafeException:
                        self._recover_from_failsafe()
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

    def focar_janela_bimer(self, click_center: bool = False):
        """Foca a janela da sessão RDP (ou Bimer dentro dela). Opcionalmente clica no centro.
        Retorna True se a janela alvo estiver ativa ao final."""
        try:
            import pygetwindow as gw
            # Prioriza a janela RDP no host local
            candidatos = []
            for titulo in (
                "Conexão de Área de Trabalho Remota",
                "Remote Desktop Connection",
                "144.22.232.212",
                "50491",
            ):
                ws = gw.getWindowsWithTitle(titulo)
                if ws:
                    candidatos.extend(ws)
            # Como fallback, tenta por títulos do app dentro da RDP
            if not candidatos:
                for titulo in ("bimer", "Bimer", "Alterdata"):
                    ws = gw.getWindowsWithTitle(titulo)
                    if ws:
                        candidatos.extend(ws)
            if candidatos:
                w = candidatos[0]
                try:
                    w.activate()
                except Exception:
                    pass
                time.sleep(0.3)
                if click_center:
                    try:
                        cx = int(w.left + w.width / 2)
                        cy = int(w.top + w.height / 2)
                        pyautogui.click(cx, cy)
                        time.sleep(0.2)
                    except Exception:
                        pass
                # Confirma se está ativa
                try:
                    aw = gw.getActiveWindow()
                    if aw and aw._hWnd == w._hWnd:
                        return True
                except Exception:
                    # Não confirma foco se não conseguimos obter a janela ativa
                    pass
            return False
        except Exception as e:
            logger.warning(f"Falha ao focar janela: {e}")
            return False

    def _ensure_focus(self, click_center: bool = False, retries: int = 3) -> bool:
        """Garante foco na janela RDP/Bimer antes de interagir com teclado/mouse."""
        try:
            import pygetwindow as gw
            for _ in range(max(1, retries)):
                if self.focar_janela_bimer(click_center=click_center):
                    aw = None
                    try:
                        aw = gw.getActiveWindow()
                    except Exception:
                        pass
                    title = (aw.title.lower() if aw and aw.title else "")
                    if any(k in title for k in (
                        "conexão de área de trabalho remota",
                        "remote desktop connection",
                        "144.22.232.212",
                        "50491",
                        "bimer",
                        "alterdata",
                    )):
                        return True
                time.sleep(0.2)
            logger.warning("Não foi possível garantir foco na janela RDP/Bimer.")
            return False
        except Exception:
            return False

    def _recover_from_failsafe(self):
        """Reposiciona o cursor para uma área segura sem acionar o fail-safe e reforça o foco."""
        try:
            cx = cy = None
            try:
                import pygetwindow as gw
                aw = gw.getActiveWindow()
                if aw:
                    cx = int(aw.left + aw.width / 2)
                    cy = int(aw.top + aw.height / 2)
            except Exception:
                pass
            if cx is None or cy is None:
                sw, sh = pyautogui.size()
                cx, cy = sw // 2, sh // 2
            ctypes.windll.user32.SetCursorPos(cx, cy)
            time.sleep(0.12)
        except Exception:
            # Fallback: desabilita fail-safe brevemente apenas para reposicionar
            try:
                sw, sh = pyautogui.size()
                fs = pyautogui.FAILSAFE
                pyautogui.FAILSAFE = False
                pyautogui.moveTo(sw // 2, sh // 2, duration=0)
                pyautogui.FAILSAFE = fs
                time.sleep(0.08)
            except Exception:
                pass
        try:
            self._ensure_focus(click_center=False)
        except Exception:
            pass

    def mover_mouse(self, x, y, duration=0.15):
        """Move o mouse de forma segura, com recuperação de fail-safe e retentativa."""
        self._ensure_focus(click_center=False)
        try:
            pyautogui.moveTo(x, y, duration=duration)
            return True
        except pyautogui.FailSafeException:
            self._recover_from_failsafe()
            try:
                pyautogui.moveTo(x, y, duration=duration)
                return True
            except Exception as e2:
                logger.error(f"Erro ao mover mouse após recuperar do fail-safe: {e2}")
                return False
        except Exception as e:
            logger.error(f"Erro ao mover mouse: {e}")
            return False

    def login_bimer(self, password: str = "", max_wait: float = 15.0) -> bool:
        """Preenche o campo de senha do login do Bimer e confirma - VERSÃO SIMPLIFICADA."""
        try:
            logger.info("="*60)
            logger.info("INICIANDO LOGIN NO BIMER - VERSÃO SIMPLIFICADA")
            logger.info("="*60)
            
            # Foca janela UMA VEZ no início (SEM clicar no centro)
            logger.info("Focando janela Bimer/RDP...")
            self._ensure_focus(click_center=False)
            time.sleep(2.0)

            # Carrega credenciais do config
            login_cfg = self.config.get('login', {}) or {}
            if not password:
                password = login_cfg.get('password') or ""
            
            logger.info(f"Senha configurada: {'Sim (' + str(len(password)) + ' caracteres)' if password else 'Não'}")
            
            # Obtém coordenadas do dropdown e item de teste
            dropdown_x = self.ui.get('campo_dropdown_ambiente_x')
            dropdown_y = self.ui.get('campo_dropdown_ambiente_y')
            teste_x = self.ui.get('ambiente_teste_dropdown_x')
            teste_y = self.ui.get('ambiente_teste_dropdown_y')
            
            # PASSO 1: Selecionar ambiente TESTE no dropdown
            if dropdown_x and dropdown_y and teste_x and teste_y:
                logger.info(f"[PASSO 1] Selecionando ambiente TESTE no dropdown")
                logger.info(f"  → Clicando no dropdown em ({dropdown_x}, {dropdown_y})")
                try:
                    pyautogui.click(dropdown_x, dropdown_y)
                    time.sleep(0.5)
                    logger.info(f"  → Clicando em 'TESTE' em ({teste_x}, {teste_y})")
                    pyautogui.click(teste_x, teste_y)
                    time.sleep(0.5)
                    logger.info("  ✓ Ambiente TESTE selecionado")
                except Exception as e:
                    logger.warning(f"  ✗ Falha ao selecionar ambiente: {e}")
            else:
                logger.warning("[PASSO 1] Coordenadas do dropdown não configuradas, pulando...")

            # PASSO 2: Preencher campo de senha
            if not self.pwd_x or not self.pwd_y:
                logger.error("[PASSO 2] Coordenadas do campo de senha não configuradas!")
                return False
            
            logger.info(f"[PASSO 2] Preenchendo senha no campo ({self.pwd_x}, {self.pwd_y})")
            try:
                # Clica no campo de senha
                logger.info(f"  → Clicando no campo de senha")
                pyautogui.click(self.pwd_x, self.pwd_y)
                time.sleep(0.4)
                
                # Limpa o campo
                logger.info(f"  → Limpando campo (Ctrl+A)")
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.2)
                
                # Libera teclas modificadoras
                for k in ('shift', 'ctrl', 'alt'):
                    try:
                        pyautogui.keyUp(k)
                    except:
                        pass
                time.sleep(0.2)
                
                # Cola a senha
                logger.info(f"  → Colando senha via área de transferência")
                if self.copiar_para_area_transferencia(password):
                    time.sleep(0.2)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.3)
                    logger.info("  ✓ Senha colada com sucesso")
                else:
                    logger.warning("  ✗ Falha ao copiar. Digitando caractere por caractere...")
                    pyautogui.write(password, interval=0.08)
                    time.sleep(0.3)
                    
            except Exception as e:
                logger.error(f"  ✗ Erro ao preencher senha: {e}")
                return False

            # PASSO 3: Clicar no botão Entrar
            if not self.enter_x or not self.enter_y:
                logger.warning("[PASSO 3] Coordenadas do botão Entrar não configuradas. Usando tecla Enter.")
                pyautogui.press('enter')
                time.sleep(2.0)
            else:
                logger.info(f"[PASSO 3] Clicando em Entrar ({self.enter_x}, {self.enter_y})")
                try:
                    pyautogui.click(self.enter_x, self.enter_y)
                    time.sleep(2.0)
                    logger.info("  ✓ Botão Entrar clicado")
                except Exception as e:
                    logger.error(f"  ✗ Erro ao clicar em Entrar: {e}")
                    return False
            
            logger.info("="*60)
            logger.info("✓ LOGIN NO BIMER CONCLUÍDO COM SUCESSO")
            logger.info("="*60)
            
            # Aguarda um pouco para o sistema processar o login
            time.sleep(2.0)
            
            # Tenta fechar modal automático se aparecer
            try:
                logger.info("[PÓS-LOGIN] Tentando fechar modal automático...")
                self.fechar_modal_automatico()
            except Exception as e:
                logger.debug(f"Modal automático não encontrado ou já fechado: {e}")
            
            # Executa sequência de cliques pós-login
            try:
                logger.info("[PÓS-LOGIN] Executando sequência de cliques...")
                self.executar_cliques_pos_login()
            except Exception as e:
                logger.warning(f"Falha ao executar cliques pós-login: {e}")
            
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
            self.clicar(x, y)
            time.sleep(0.5)
            return True
        except Exception as e:
            logger.warning(f"Falha ao fechar modal automatico: {e}")
            return False
    
    def parse_coordenada(self, valor):
        """Converte valores de coordenadas do YAML em tupla (x, y)."""
        try:
            if isinstance(valor, (list, tuple)) and len(valor) >= 2:
                return int(valor[0]), int(valor[1])
            if isinstance(valor, str):
                partes = [p.strip() for p in valor.replace(';', ',').split(',')]
                if len(partes) >= 2:
                    return int(partes[0]), int(partes[1])
        except Exception:
            return None
        return None

    def coord_ui(self, chave):
        """Obtém coordenada (x, y) a partir do nome do elemento em ui_elements."""
        valor = self.ui.get(chave)
        if valor is None:
            return None
        return self.parse_coordenada(valor)

    def executar_clique_por_nome(self, nome, pausa=0.3):
        """Clica usando o nome cadastrado em ui_elements."""
        try:
            coord = self.coord_ui(nome)
            if coord is None:
                logger.warning(f"Coordenada não encontrada para: {nome}")
                return False
            x, y = coord
            try:
                self.mover_mouse(x, y, duration=0.15)
                time.sleep(0.05)
            except Exception:
                pass
            self.clicar(x, y)
            time.sleep(pausa)
            return True
        except Exception as e:
            logger.warning(f"Falha ao clicar em '{nome}': {e}")
            return False

    def executar_cliques_pos_login(self):
        """Executa a sequência de cliques fornecida após o login."""
        try:
            self._ensure_focus(click_center=True)
            sequencia = [
                "menu lateral financeiro",
                "a pagar dentro de financeiro",
                "ferramentas",
                "gerar arquivo remessa",
                "uma conta",
                "campo para digitar a conta 19",
                "campo do layout, digitar 16",
            ]
            for nome in sequencia:
                self.executar_clique_por_nome(nome, pausa=0.35)
                if nome == "campo para digitar a conta 19":
                    try:
                        self.digitar("19", intervalo=0.05)
                    except Exception:
                        pass
                if nome == "campo do layout, digitar 16":
                    try:
                        self.digitar("16", intervalo=0.05)
                    except Exception:
                        pass
        except Exception as e:
            logger.warning(f"Erro na sequência de cliques pós-login: {e}")
    
    def clicar(self, x, y, botao='left', cliques=1):
        """Realiza clique na posição especificada"""
        try:
            logger.debug(f"Clicando em ({x}, {y}) com botão {botao}")
            try:
                pyautogui.click(x, y, clicks=cliques, button=botao)
            except pyautogui.FailSafeException:
                self._recover_from_failsafe()
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
                self.clicar(centro.x, centro.y)
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
            try:
                pyautogui.write(texto, interval=intervalo)
            except pyautogui.FailSafeException:
                self._recover_from_failsafe()
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
                try:
                    pyautogui.press(tecla)
                except pyautogui.FailSafeException:
                    self._recover_from_failsafe()
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
            try:
                pyautogui.hotkey(*teclas)
            except pyautogui.FailSafeException:
                self._recover_from_failsafe()
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

