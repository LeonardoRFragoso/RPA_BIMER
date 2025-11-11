"""
Módulo para gerenciar conexão RDP com a VM Windows
"""
import os
import subprocess
import tempfile
from pathlib import Path
from src.config import get_vm_config
from src.logger import logger

class RDPConnection:
    """Classe para gerenciar conexão RDP"""
    
    def __init__(self):
        # Usa YAML apenas para evitar sobrescrita indevida por variáveis de ambiente
        self.vm_config = get_vm_config(prefer_env=False)
        self.host = self.vm_config.get('host', '')
        self.port = self.vm_config.get('port', 3389)
        self.username = self.vm_config.get('username', '')
        self.password = self.vm_config.get('password', '')
    
    def criar_arquivo_rdp(self, caminho=None):
        """Cria um arquivo .rdp temporário com as configurações"""
        if caminho is None:
            # Cria arquivo temporário
            temp_dir = tempfile.gettempdir()
            caminho = os.path.join(temp_dir, 'bimmer_connection.rdp')
        
        # Conteúdo do arquivo RDP
        conteudo_rdp = f"""screen mode id:i:2
use multimon:i:0
desktopwidth:i:1920
desktopheight:i:1080
session bpp:i:32
winposstr:s:0,1,0,0,1920,1080
compression:i:1
keyboardhook:i:2
audiocapturemode:i:0
videoplaybackmode:i:1
connection type:i:7
networkautodetect:i:1
bandwidthautodetect:i:1
enableworkspacereconnect:i:0
disable wallpaper:i:0
allow font smoothing:i:0
allow desktop composition:i:0
disable full window drag:i:1
disable menu anims:i:1
disable themes:i:0
disable cursor setting:i:0
bitmapcachepersistenable:i:1
full address:s:{self.host}:{self.port}
audiomode:i:0
redirectprinters:i:1
redirectcomports:i:0
redirectsmartcards:i:1
redirectclipboard:i:1
redirectposdevices:i:0
autoreconnection enabled:i:1
authentication level:i:2
prompt for credentials:i:0
negotiate security layer:i:1
enablecredsspsupport:i:1
credsspcredentialstype:i:1
remoteapplicationmode:i:0
alternate shell:s:
shell working directory:s:
gatewayhostname:s:
gatewayusagemethod:i:4
gatewaycredentialssource:i:4
gatewayprofileusagemethod:i:0
promptcredentialonce:i:0
gatewaybrokeringtype:i:0
use redirection server name:i:0
rdgiskdcproxy:i:0
kdcproxyname:s:
username:s:{self.username}
domain:s:
"""
        
        try:
            with open(caminho, 'w', encoding='utf-8') as f:
                f.write(conteudo_rdp)
            
            logger.info(f"Arquivo RDP criado em: {caminho}")
            return caminho
            
        except Exception as e:
            logger.error(f"Erro ao criar arquivo RDP: {str(e)}")
            return None
    
    def conectar_rdp(self, usar_arquivo=True):
        """Conecta à VM via RDP usando mstsc"""
        try:
            logger.info(f"Conectando à VM {self.host}:{self.port} como {self.username}...")
            
            if usar_arquivo:
                # Cria arquivo RDP e abre
                arquivo_rdp = self.criar_arquivo_rdp()
                if arquivo_rdp:
                    # Abre o arquivo RDP com mstsc
                    subprocess.Popen(['mstsc', arquivo_rdp])
                    logger.info("Conexão RDP iniciada via arquivo")
                    return True
            else:
                # Conecta diretamente via linha de comando
                # Nota: mstsc não aceita senha diretamente na linha de comando por segurança
                # Será necessário usar o arquivo RDP ou autenticação interativa
                cmd = ['mstsc', f'/v:{self.host}:{self.port}']
                subprocess.Popen(cmd)
                logger.info("Conexão RDP iniciada (será necessário inserir credenciais)")
                return True
                
        except Exception as e:
            logger.error(f"Erro ao conectar via RDP: {str(e)}")
            return False
    
    def conectar_com_senha(self):
        """
        Conecta via RDP com senha usando cmdkey e arquivo RDP
        Nota: Requer cmdkey para armazenar credenciais
        """
        import time
        
        try:
            logger.info(f"Configurando credenciais para {self.host}:{self.port}...")
            
            # Primeiro, remove credenciais antigas se existirem (tenta diferentes formatos)
            targets = [
                f"TERMSRV/{self.host}:{self.port}",
                f"TERMSRV/{self.host}",
                f"{self.host}:{self.port}",
                f"{self.host}"
            ]
            
            for target in targets:
                subprocess.run(['cmdkey', '/delete:{}'.format(target)], 
                              capture_output=True, text=True)
            
            # Cria arquivo RDP primeiro (com username correto)
            arquivo_rdp = self.criar_arquivo_rdp()
            
            # Armazena credenciais usando cmdkey (formato correto)
            # Usa o formato TERMSRV/host:port para RDP
            target = f"TERMSRV/{self.host}:{self.port}"
            
            # Tenta armazenar credenciais
            cmd_cmdkey = [
                'cmdkey',
                '/generic:{}'.format(target),
                '/user:{}'.format(self.username),
                '/pass:{}'.format(self.password)
            ]
            
            result = subprocess.run(cmd_cmdkey, capture_output=True, text=True, shell=True)
            
            if result.returncode == 0:
                logger.info("Credenciais armazenadas com sucesso")
                logger.debug(f"Output cmdkey: {result.stdout}")
            else:
                logger.warning(f"Possível erro ao armazenar credenciais: {result.stderr}")
                logger.debug(f"Output cmdkey: {result.stdout}")
            
            # Aguarda um pouco para garantir que as credenciais foram armazenadas
            time.sleep(2)
            
            # Verifica se as credenciais foram armazenadas
            cmd_list = ['cmdkey', '/list:{}'.format(target)]
            result_list = subprocess.run(cmd_list, capture_output=True, text=True, shell=True)
            
            if result_list.returncode == 0 and self.username in result_list.stdout:
                logger.info("Credenciais verificadas e confirmadas")
            else:
                logger.warning("Credenciais podem não ter sido armazenadas corretamente")
            
            # Conecta usando o arquivo RDP (que tem o username correto)
            if arquivo_rdp:
                logger.info(f"Conectando usando arquivo RDP: {arquivo_rdp}")
                # Usa o arquivo RDP que já tem o username configurado
                subprocess.Popen(['mstsc', arquivo_rdp], shell=True)
                logger.info("Conexão RDP iniciada com credenciais armazenadas")
                return True
            else:
                # Fallback: conecta diretamente
                logger.warning("Arquivo RDP não criado, usando conexão direta")
                cmd_mstsc = ['mstsc', f'/v:{self.host}:{self.port}']
                subprocess.Popen(cmd_mstsc, shell=True)
                logger.info("Conexão RDP iniciada (fallback)")
                return True
                
        except Exception as e:
            logger.error(f"Erro ao conectar com senha: {str(e)}")
            return False
    
    def remover_credenciais(self):
        """Remove credenciais armazenadas"""
        try:
            target = f"TERMSRV/{self.host}:{self.port}"
            cmd = ['cmdkey', '/delete:{}'.format(target)]
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                logger.info("Credenciais removidas com sucesso")
                return True
            else:
                logger.warning(f"Erro ao remover credenciais: {result.stderr}")
                return False
                
        except Exception as e:
            logger.error(f"Erro ao remover credenciais: {str(e)}")
            return False

