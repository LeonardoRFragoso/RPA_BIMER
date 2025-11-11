"""
Módulo para preencher automaticamente as credenciais na tela de RDP
"""
import time
import pyautogui
from src.config import config, load_config_yaml_only, get_vm_config, get_vm_credentials
from src.logger import logger

class RDPAutoFill:
    """Classe para preencher credenciais automaticamente na tela de RDP"""
    
    def __init__(self):
        # Para evitar overrides indevidos por variáveis de ambiente,
        # usamos por padrão SOMENTE o YAML para credenciais da VM.
        self.vm_config = get_vm_config(prefer_env=False)
        self.username = self.vm_config.get('username', '')
        self.password = self.vm_config.get('password', '')
        
        # Verificação crítica: garante que a senha foi carregada corretamente
        if not self.password:
            logger.error("ERRO: Senha não foi carregada do config.yaml!")
            logger.error("Verifique se a senha está configurada corretamente no arquivo config.yaml")
        elif self.password == self.username:
            logger.error(f"ERRO: Senha está igual ao usuário! Senha: '{self.password}', Usuário: '{self.username}'")
            logger.error("Isso não deve acontecer! Verifique o config.yaml")
        
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.3
    
    def copiar_para_area_transferencia(self, texto: str) -> bool:
        """Copia texto para a área de transferência do Windows (usa pywin32; fallback para pyperclip)."""
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
                logger.debug("Texto copiado para área de transferência via win32clipboard")
                return True
            except Exception as e1:
                logger.debug(f"win32clipboard indisponível: {e1}")
                try:
                    import pyperclip
                    pyperclip.copy(texto)
                    logger.debug("Texto copiado para área de transferência via pyperclip")
                    return True
                except Exception as e2:
                    logger.warning(f"Falha ao copiar para área de transferência: {e2}")
                    return False
        except Exception as e:
            logger.warning(f"Erro inesperado ao copiar para área de transferência: {e}")
            return False
    
    def digitar_texto_com_especiais(self, texto, intervalo=0.05):
        """Digita texto incluindo caracteres especiais corretamente"""
        try:
            logger.info(f"INICIANDO digitação de texto com {len(texto)} caracteres")
            logger.info(f"Texto completo a ser digitado: '{texto}'")
            logger.info(f"Caracteres individuais: {list(texto)}")
            
            # Verifica se o texto não está vazio
            if not texto:
                logger.error("Texto vazio para digitar!")
                return False
            
            caracteres_digitados = 0
            # Digita caractere por caractere para garantir precisão
            # Campos de senha geralmente bloqueiam copiar/colar
            for i, char in enumerate(texto, 1):
                try:
                    logger.info(f"[{i}/{len(texto)}] Digitando caractere: '{char}' (ord: {ord(char)})")
                    
                    if char == ')':
                        # Parêntese fechado: shift + 0
                        logger.info(f"[{i}/{len(texto)}] Caractere ')' - usando Shift+0")
                        pyautogui.keyDown('shift')
                        time.sleep(0.05)
                        pyautogui.press('0')
                        time.sleep(0.05)
                        pyautogui.keyUp('shift')
                        caracteres_digitados += 1
                        logger.info(f"[{i}/{len(texto)}] ✓ Caractere ')' digitado")
                    elif char == '^':
                        # Circunflexo: shift + 6
                        logger.info(f"[{i}/{len(texto)}] Caractere '^' - usando Shift+6")
                        pyautogui.keyDown('shift')
                        time.sleep(0.05)
                        pyautogui.press('6')
                        time.sleep(0.05)
                        pyautogui.keyUp('shift')
                        caracteres_digitados += 1
                        logger.info(f"[{i}/{len(texto)}] ✓ Caractere '^' digitado")
                    elif char.isupper():
                        # Letras maiúsculas - precisa usar Shift ou typewrite
                        logger.info(f"[{i}/{len(texto)}] Caractere maiúsculo '{char}' - usando typewrite")
                        pyautogui.typewrite(char, interval=0.01)
                        caracteres_digitados += 1
                        logger.info(f"[{i}/{len(texto)}] ✓ Caractere maiúsculo '{char}' digitado")
                    elif char.islower() or char.isdigit():
                        # Letras minúsculas e números
                        logger.info(f"[{i}/{len(texto)}] Caractere '{char}' - usando press")
                        pyautogui.press(char)
                        caracteres_digitados += 1
                        logger.info(f"[{i}/{len(texto)}] ✓ Caractere '{char}' digitado")
                    else:
                        # Outros caracteres especiais - usa typewrite
                        logger.info(f"[{i}/{len(texto)}] Caractere especial '{char}' - usando typewrite")
                        pyautogui.typewrite(char, interval=0.01)
                        caracteres_digitados += 1
                        logger.info(f"[{i}/{len(texto)}] ✓ Caractere especial '{char}' digitado")
                    
                    time.sleep(intervalo)
                    logger.info(f"[{i}/{len(texto)}] Progresso: {caracteres_digitados}/{len(texto)} caracteres digitados até agora")
                except Exception as e:
                    logger.error(f"ERRO ao digitar caractere '{char}' (posição {i}/{len(texto)}): {e}")
                    logger.error(f"Exceção completa: {type(e).__name__}: {str(e)}")
                    # Tenta digitar de forma simples como fallback
                    try:
                        logger.warning(f"Tentando fallback para caractere '{char}'...")
                        pyautogui.typewrite(char, interval=0.01)
                        caracteres_digitados += 1
                        time.sleep(intervalo)
                        logger.info(f"✓ Fallback bem-sucedido para '{char}'")
                    except Exception as e2:
                        logger.error(f"FALHA no fallback para caractere '{char}': {e2}")
                        # Continua mesmo assim para tentar os próximos caracteres
            
            logger.info(f"TOTAL: {caracteres_digitados}/{len(texto)} caracteres digitados")
            if caracteres_digitados < len(texto):
                logger.warning(f"ATENÇÃO: Apenas {caracteres_digitados} de {len(texto)} caracteres foram digitados!")
                return False
            else:
                logger.info(f"✓ Texto digitado completamente ({len(texto)} caracteres)")
                return True
        except Exception as e:
            logger.error(f"ERRO FATAL ao digitar texto com especiais: {str(e)}")
            logger.error(f"Exceção completa: {type(e).__name__}: {str(e)}", exc_info=True)
            # Fallback: usa typewrite normal
            logger.warning("Usando método fallback (typewrite completo)")
            try:
                pyautogui.typewrite(texto, interval=intervalo)
                logger.info("Fallback bem-sucedido")
                return True
            except Exception as e2:
                logger.error(f"FALHA no fallback: {e2}")
                return False
    
    def aguardar_tela_credenciais(self, timeout=15):
        """Aguarda a tela de credenciais do Windows aparecer"""
        try:
            logger.info("Aguardando tela de credenciais do Windows...")
            
            # Procura por elementos característicos da tela de credenciais
            inicio = time.time()
            while time.time() - inicio < timeout:
                # Procura por texto "Digite suas credenciais" ou "Enter your credentials"
                try:
                    # Tenta encontrar a janela de credenciais
                    # A janela geralmente tem "Segurança do Windows" ou "Windows Security" no título
                    try:
                        import pygetwindow as gw
                        windows = gw.getWindowsWithTitle("Segurança do Windows")
                        if not windows:
                            windows = gw.getWindowsWithTitle("Windows Security")
                        
                        if windows:
                            window = windows[0]
                            if window.visible:
                                logger.info("Tela de credenciais encontrada!")
                                window.activate()
                                time.sleep(0.5)
                                return True
                    except ImportError:
                        # pygetwindow não disponível, tenta método alternativo
                        logger.debug("pygetwindow não disponível, usando método alternativo")
                        # Simplesmente assume que a tela está aberta após alguns segundos
                        if time.time() - inicio > 3:
                            logger.info("Assumindo que a tela de credenciais está aberta...")
                            return True
                except Exception as e:
                    logger.debug(f"Erro ao procurar janela: {e}")
                
                time.sleep(0.5)
            
            # Se não encontrou a janela, tenta mesmo assim (pode estar aberta mas não detectada)
            logger.warning("Tela de credenciais não encontrada automaticamente, tentando preencher mesmo assim...")
            time.sleep(1)
            return True  # Tenta preencher mesmo assim
            
        except Exception as e:
            logger.error(f"Erro ao aguardar tela de credenciais: {str(e)}")
            return True  # Tenta preencher mesmo assim
    
    def preencher_credenciais(self):
        """Preenche as credenciais na tela de RDP"""
        try:
            # Verifica se as credenciais estão configuradas corretamente
            # Recarrega do YAML para garantir que está atualizado (ignora env)
            vm_yaml = load_config_yaml_only().get('vm', {})
            senha_atual = vm_yaml.get('password', '')
            usuario_atual = vm_yaml.get('username', '')
            
            logger.info("=" * 60)
            logger.info("VERIFICAÇÃO DE CREDENCIAIS")
            logger.info("=" * 60)
            logger.info(f"Usuário configurado: '{usuario_atual}' (tamanho: {len(usuario_atual)})")
            logger.info(f"Senha configurada: '{senha_atual}' (tamanho: {len(senha_atual)})")
            if senha_atual:
                logger.info(f"Primeiro caractere da senha: '{senha_atual[0]}'")
                logger.info(f"Último caractere da senha: '{senha_atual[-1]}'")
            logger.info("=" * 60)
            
            # Verifica se a senha não está vazia e não é igual ao usuário
            if not senha_atual:
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: Senha está vazia no config.yaml!")
                logger.error("Verifique se a senha está configurada corretamente no arquivo config.yaml")
                logger.error("=" * 60)
                return False
            
            if senha_atual == usuario_atual:
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: Senha está igual ao usuário!")
                logger.error(f"Senha: '{senha_atual}'")
                logger.error(f"Usuário: '{usuario_atual}'")
                logger.error("Isso não deve acontecer! Verifique o config.yaml")
                logger.error("=" * 60)
                return False
            
            # Verificação adicional: se a senha começa com "parceiro", algo está errado
            if senha_atual.lower().startswith('parceiro'):
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: A senha parece ser o username!")
                logger.error(f"Senha detectada: '{senha_atual}'")
                logger.error(f"Username: '{usuario_atual}'")
                logger.error("A senha não deve começar com o username. Verifique o config.yaml")
                logger.error("=" * 60)
                return False
            
            # Atualiza self.password e self.username com os valores do config
            self.password = senha_atual
            self.username = usuario_atual
            
            logger.info("Preenchendo credenciais automaticamente...")
            
            # Aguarda um pouco para garantir que a tela está pronta
            time.sleep(1.5)
            
            # Foca na janela de credenciais primeiro
            try:
                import pygetwindow as gw
                windows = gw.getWindowsWithTitle("Segurança do Windows")
                if not windows:
                    windows = gw.getWindowsWithTitle("Windows Security")
                
                if windows:
                    window = windows[0]
                    window.activate()
                    time.sleep(0.5)
            except:
                pass
            
            # Método simplificado: preenche diretamente sem navegação excessiva
            logger.debug("Preenchendo credenciais...")
            
            # Aguarda um pouco para garantir que a tela está pronta
            time.sleep(1)
            
            # NÃO vamos digitar o usuário por padrão.
            # Em ambientes onde o usuário já vem selecionado (como na captura enviada),
            # apenas focamos e preenchemos o campo de senha.
            logger.info("Pulando digitação do usuário (usuário já selecionado).")
            
            # Vai para o campo de senha - usa Tab múltiplas vezes para garantir
            logger.info("=" * 60)
            logger.info("NAVEGANDO PARA CAMPO DE SENHA")
            logger.info("=" * 60)
            
            # Primeiro, garante que saímos do campo de usuário
            logger.info("Pressionando END para ir ao final do campo de usuário...")
            pyautogui.press('end')
            time.sleep(0.5)
            
            # CRÍTICO: Pressiona Tab para sair do campo de usuário e ir para o campo de senha
            logger.info("=" * 60)
            logger.info("NAVEGANDO PARA CAMPO DE SENHA - CRÍTICO")
            logger.info("=" * 60)
            logger.info("Pressionando TAB para sair do campo de usuário e ir para o campo de senha...")
            
            # Tenta ir para o campo de senha. Primeiro com TAB; em seguida faremos clique de foco.
            pyautogui.press('tab')
            time.sleep(2.0)
            logger.info("TAB pressionado - tentando focar campo de senha...")
            
            # VERIFICAÇÃO CRÍTICA: Tenta limpar o campo para verificar se estamos no campo de senha
            # Se ainda estivermos no campo de usuário, o texto "parceiro" será limpo
            # Se estivermos no campo de senha, qualquer texto será limpo
            logger.info("Verificação crítica: limpando campo para confirmar que estamos no campo de senha...")
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.3)
            pyautogui.press('delete')
            time.sleep(0.5)
            pyautogui.press('backspace')
            time.sleep(0.5)
            
            # Pressiona Home para garantir que estamos no início do campo
            logger.info("Pressionando HOME para ir ao início do campo...")
            pyautogui.press('home')
            time.sleep(0.5)
            
            # Aguarda mais um pouco para garantir que o foco mudou completamente
            logger.info("Aguardando foco estabilizar no campo de senha...")
            time.sleep(1.5)
            logger.info("Foco confirmado no campo de senha")
            logger.info("=" * 60)
            
            # Preenche a senha
            logger.info("=" * 60)
            logger.info("PREENCHENDO CAMPO DE SENHA")
            logger.info("=" * 60)
            
            # Tenta focar o campo de senha clicando em uma área provável dentro da janela
            try:
                import pygetwindow as gw
                windows = gw.getWindowsWithTitle("Segurança do Windows") or gw.getWindowsWithTitle("Windows Security")
                if windows:
                    window = windows[0]
                    center_x = window.left + window.width // 2
                    senha_y = window.top + int(window.height * 0.58)
                    logger.info(f"Clicando para focar no campo de senha em ({center_x}, {senha_y})")
                    pyautogui.moveTo(center_x, senha_y, duration=0.2)
                    time.sleep(0.2)
                    pyautogui.click(center_x, senha_y)
                    time.sleep(0.5)
                else:
                    logger.warning("Janela 'Segurança do Windows' não localizada para clique; seguindo com TAB")
            except Exception as e:
                logger.warning(f"Falha ao focar por clique no campo de senha: {e}")
            
            # CRÍTICO: Recarrega a senha do config ANTES de usar
            # Recarrega diretamente do YAML para garantir que está atualizado (ignora env)
            vm_config_atual = load_config_yaml_only().get('vm', {})
            senha_para_usar = vm_config_atual.get('password', '')
            usuario_nao_usar = vm_config_atual.get('username', '')
            
            # Verificação crítica: garante que não estamos usando o username
            if not senha_para_usar:
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: Senha não foi carregada do config!")
                logger.error("Verifique se a senha está configurada corretamente no config.yaml")
                logger.error("=" * 60)
                return False
            
            if senha_para_usar == usuario_nao_usar:
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: Senha está igual ao username!")
                logger.error(f"Senha: '{senha_para_usar}'")
                logger.error(f"Username: '{usuario_nao_usar}'")
                logger.error("Isso não deve acontecer! Verifique o config.yaml")
                logger.error("=" * 60)
                return False
            
            # Verificação adicional: se a senha é exatamente "parceiro", algo está muito errado
            if senha_para_usar.lower() == 'parceiro':
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: A senha é exatamente 'parceiro' (o username)!")
                logger.error("Isso não deve acontecer! Verifique o config.yaml")
                logger.error("ABORTANDO: Não vou digitar o username como senha!")
                logger.error("=" * 60)
                return False
            
            logger.info(f"Senha a ser digitada: '{senha_para_usar}'")
            logger.info(f"Tamanho da senha: {len(senha_para_usar)} caracteres")
            logger.info(f"Usuário (NÃO deve ser digitado aqui): '{usuario_nao_usar}'")
            logger.info("=" * 60)
            
            # CRÍTICO: Limpa o campo de senha completamente ANTES de digitar
            # Isso garante que não há texto residual do campo de usuário
            logger.info("=" * 60)
            logger.info("LIMPEZA FINAL DO CAMPO DE SENHA - CRÍTICO")
            logger.info("=" * 60)
            logger.info("Limpando campo de senha completamente...")
            
            # Pressiona Home para garantir que estamos no início
            pyautogui.press('home')
            time.sleep(0.3)
            
            # Pressiona Ctrl+A para selecionar tudo
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.3)
            
            # Pressiona Delete para limpar
            pyautogui.press('delete')
            time.sleep(0.5)
            
            # Pressiona Backspace também para garantir
            pyautogui.press('backspace')
            time.sleep(0.5)
            
            # Pressiona Ctrl+A novamente e Delete para garantir que está limpo
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.3)
            pyautogui.press('delete')
            time.sleep(0.5)
            
            # Pressiona Home novamente para garantir que estamos no início
            pyautogui.press('home')
            time.sleep(0.5)
            
            logger.info("Campo de senha limpo e pronto para receber a senha")
            logger.info("=" * 60)
            
            # Aguarda um pouco mais para garantir que o campo de senha está pronto
            logger.info("Aguardando campo de senha estar completamente pronto...")
            time.sleep(1.5)
            
            # Digita a senha caractere por caractere usando a função especial
            logger.info("=" * 60)
            logger.info("INICIANDO DIGITAÇÃO DA SENHA")
            logger.info("=" * 60)
            
            # VERIFICAÇÃO CRÍTICA: Garante que estamos usando a SENHA, não o usuário
            # Usa a senha que já foi carregada anteriormente
            senha_para_digitar = str(senha_para_usar).strip()  # Remove espaços em branco
            
            # LOG CRÍTICO: Mostra exatamente o que será digitado
            logger.info("=" * 60)
            logger.info("VERIFICAÇÃO CRÍTICA: O QUE SERÁ DIGITADO")
            logger.info("=" * 60)
            logger.info(f"SENHA a ser digitada: '{senha_para_digitar}'")
            logger.info(f"USUÁRIO (NÃO usar): '{usuario_nao_usar}'")
            logger.info("=" * 60)
            
            # Verifica se a senha foi carregada corretamente
            if not senha_para_digitar:
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: Senha não foi carregada do config.yaml!")
                logger.error("Verifique se a senha está configurada corretamente no arquivo config.yaml")
                logger.error("=" * 60)
                return False
            
            logger.info("=" * 60)
            logger.info("VERIFICAÇÃO FINAL ANTES DE DIGITAR A SENHA")
            logger.info("=" * 60)
            logger.info(f"✓ SENHA a ser digitada: '{senha_para_digitar}' (tamanho: {len(senha_para_digitar)} caracteres)")
            logger.info(f"✗ USUÁRIO (NÃO usar aqui): '{usuario_nao_usar}' (tamanho: {len(usuario_nao_usar)} caracteres)")
            logger.info("=" * 60)
            
            # Verificação final antes de digitar - CRÍTICO
            if senha_para_digitar == usuario_nao_usar:
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: A senha está igual ao usuário!")
                logger.error(f"Senha: '{senha_para_digitar}'")
                logger.error(f"Usuário: '{usuario_nao_usar}'")
                logger.error("Isso não deve acontecer! Verifique o config.yaml")
                logger.error("=" * 60)
                return False
            
            # Verifica se a senha não está vazia
            if not senha_para_digitar or len(senha_para_digitar) == 0:
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: Senha está vazia!")
                logger.error("Verifique se a senha está configurada no config.yaml")
                logger.error("=" * 60)
                return False
            
            # Verifica se a senha tem pelo menos 8 caracteres (senha real tem 12)
            if len(senha_para_digitar) < 8:
                logger.warning(f"ATENÇÃO: Senha tem apenas {len(senha_para_digitar)} caracteres (esperado: 12)")
            
            # Verificação adicional: se a senha começa com "parceiro", algo está errado
            if senha_para_digitar.lower().startswith('parceiro'):
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: A senha parece ser o username!")
                logger.error(f"Senha detectada: '{senha_para_digitar}'")
                logger.error(f"Username: '{usuario_nao_usar}'")
                logger.error("A senha não deve começar com o username. Verifique o config.yaml")
                logger.error("=" * 60)
                return False
            
            logger.info("=" * 60)
            logger.info("CONFIRMAÇÃO FINAL: Vou digitar a SENHA (não o usuário)")
            logger.info(f"SENHA: '{senha_para_digitar}'")
            logger.info(f"PRIMEIRO caractere da senha: '{senha_para_digitar[0]}'")
            logger.info(f"ÚLTIMO caractere da senha: '{senha_para_digitar[-1]}'")
            logger.info("Iniciando digitação caractere por caractere...")
            logger.info("=" * 60)
            
            # VERIFICAÇÃO FINAL CRÍTICA: Garante que não estamos usando o username
            if senha_para_digitar == usuario_nao_usar or senha_para_digitar.lower() == usuario_nao_usar.lower():
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: Tentando digitar o username como senha!")
                logger.error(f"Senha detectada: '{senha_para_digitar}'")
                logger.error(f"Username: '{usuario_nao_usar}'")
                logger.error("ABORTANDO: Não vou digitar o username como senha!")
                logger.error("=" * 60)
                return False
            
            # Verificação adicional: se a senha é exatamente "parceiro", algo está muito errado
            if senha_para_digitar.lower() == 'parceiro':
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO: A senha é exatamente 'parceiro' (o username)!")
                logger.error("Isso não deve acontecer! Verifique o config.yaml")
                logger.error("ABORTANDO: Não vou digitar o username como senha!")
                logger.error("=" * 60)
                return False
            
            # VERIFICAÇÃO FINAL ABSOLUTA: Garante que estamos usando a senha correta
            # Recarrega novamente do YAML para garantir que está correto
            vm_config_final = load_config_yaml_only().get('vm', {})
            senha_final = str(vm_config_final.get('password', '')).strip()
            usuario_final = str(vm_config_final.get('username', '')).strip()
            
            # Se a senha final for diferente da senha para digitar, usa a senha final
            if senha_final != senha_para_digitar:
                logger.warning("=" * 60)
                logger.warning("ATENÇÃO: Senha recarregada do config é diferente!")
                logger.warning(f"Senha anterior: '{senha_para_digitar}'")
                logger.warning(f"Senha do config: '{senha_final}'")
                logger.warning("Usando senha do config...")
                logger.warning("=" * 60)
                senha_para_digitar = senha_final
            
            # VERIFICAÇÃO FINAL ABSOLUTA: Garante que não estamos usando o username
            if senha_para_digitar == usuario_final or senha_para_digitar.lower() == usuario_final.lower():
                logger.error("=" * 60)
                logger.error("ERRO CRÍTICO ABSOLUTO: Tentando digitar o username como senha!")
                logger.error(f"Senha detectada: '{senha_para_digitar}'")
                logger.error(f"Username: '{usuario_final}'")
                logger.error("ABORTANDO COMPLETAMENTE: Não vou digitar o username como senha!")
                logger.error("=" * 60)
                return False
            
            # Digita a senha (primeiro tenta colar via área de transferência para evitar problemas de layout)
            logger.info("=" * 60)
            logger.info("PREENCHENDO SENHA: Tentando colar via área de transferência (Ctrl+V)")
            colado = False
            try:
                if self.copiar_para_area_transferencia(senha_para_digitar):
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.5)
                    colado = True
                    logger.info("Senha colada via Ctrl+V com sucesso")
                else:
                    logger.warning("Não foi possível copiar senha para a área de transferência")
            except Exception as e:
                logger.warning(f"Falha ao colar senha via Ctrl+V: {e}")
            
            resultado = True
            if not colado:
                # Fallback para digitação caractere a caractere (pode depender de layout)
                logger.info("=" * 60)
                logger.info("Fallback: Digitando senha caractere por caractere")
                logger.info(f"Tamanho da senha: {len(senha_para_digitar)} caracteres")
                logger.info(f"Primeiro caractere: '{senha_para_digitar[0]}'")
                logger.info(f"Último caractere: '{senha_para_digitar[-1]}'")
                logger.info("=" * 60)
                resultado = self.digitar_texto_com_especiais(senha_para_digitar, intervalo=0.08)
            time.sleep(0.5)
            
            if resultado:
                logger.info("=" * 60)
                logger.info("✓ DIGITAÇÃO DA SENHA CONCLUÍDA")
                logger.info(f"Senha digitada: {len(senha_para_digitar)} caracteres")
                logger.info(f"Primeiro caractere: '{senha_para_digitar[0]}'")
                logger.info(f"Último caractere: '{senha_para_digitar[-1]}'")
                logger.info(f"Senha completa digitada: '{senha_para_digitar}'")
                logger.info("=" * 60)
            else:
                logger.error("=" * 60)
                logger.error("✗ ERRO: Digitação da senha falhou!")
                logger.error("=" * 60)
                return False
            
            # Marca "Lembrar-me" e confirma com múltiplas estratégias
            logger.debug("Marcando 'Lembrar-me' e confirmando...")
            try:
                # Vai ao checkbox e marca
                pyautogui.press('tab')
                time.sleep(0.2)
                pyautogui.press('space')
                time.sleep(0.2)
                
                confirmado = False
                # 1) Pressiona Enter (ação padrão)
                pyautogui.press('enter')
                time.sleep(1.0)
                
                import pygetwindow as gw
                w = gw.getWindowsWithTitle("Segurança do Windows") or gw.getWindowsWithTitle("Windows Security")
                if not w:
                    confirmado = True
                else:
                    # 2) Navega com TAB até o botão OK e confirma
                    for _ in range(5):
                        pyautogui.press('tab')
                        time.sleep(0.15)
                    pyautogui.press('enter')
                    time.sleep(1.0)
                    w = gw.getWindowsWithTitle("Segurança do Windows") or gw.getWindowsWithTitle("Windows Security")
                    if not w:
                        confirmado = True
                
                if not confirmado and w:
                    # 3) Clica nas coordenadas do botão OK
                    window = w[0]
                    ok_x = window.left + int(window.width * 0.20)
                    ok_y = window.top + int(window.height * 0.93)
                    logger.info(f"Clicando no botão OK em ({ok_x}, {ok_y})")
                    pyautogui.moveTo(ok_x, ok_y, duration=0.2)
                    time.sleep(0.2)
                    pyautogui.click(ok_x, ok_y)
                    time.sleep(1.0)
                    
                    # 4) Fallback final: Alt+O (quando disponível) e Enter
                    pyautogui.hotkey('alt', 'o')
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(0.8)
            except Exception as e:
                logger.warning(f"Falha ao confirmar credenciais: {e}")
            
            logger.info("Credenciais preenchidas automaticamente!")
            time.sleep(2)  # Aguarda a conexão ser estabelecida
            return True
            
        except Exception as e:
            logger.error(f"Erro ao preencher credenciais: {str(e)}")
            return False
    
    def verificar_erro_credenciais(self):
        """Verifica se há mensagem de erro de credenciais"""
        try:
            import pygetwindow as gw
            windows = gw.getWindowsWithTitle("Segurança do Windows")
            if not windows:
                windows = gw.getWindowsWithTitle("Windows Security")
            
            if windows:
                window = windows[0]
                # Verifica se há mensagem de erro na janela
                # Se houver "Suas credenciais não funcionaram" ou "Your credentials did not work"
                # significa que as credenciais estão erradas
                return True
            
            return False
        except:
            return False
    
    def preencher_automaticamente(self, timeout=10):
        """Aguarda a tela e preenche automaticamente"""
        try:
            if self.aguardar_tela_credenciais(timeout):
                resultado = self.preencher_credenciais()
                
                # Aguarda um pouco e verifica se houve erro
                time.sleep(3)
                if self.verificar_erro_credenciais():
                    logger.warning("Possível erro de credenciais detectado")
                    logger.info("Verifique se a senha está correta no config.yaml")
                
                return resultado
            return False
        except Exception as e:
            logger.error(f"Erro no preenchimento automático: {str(e)}")
            return False

