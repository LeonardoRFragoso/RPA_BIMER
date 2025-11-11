"""
Bot principal para automação de remessas bancárias no Bimmer
"""
import csv
import time
from pathlib import Path
from src.config import config, BASE_DIR
from src.logger import logger
from src.ui_automation import BimmerAutomation

class BimmerBot:
    """Bot principal para automação do Bimmer"""
    
    def __init__(self):
        self.automation = BimmerAutomation()
        self.config = config
        self.remessas_config = config.get('remessas', {})
        
    def carregar_remessas(self):
        """Carrega dados das remessas do arquivo CSV"""
        try:
            arquivo = self.remessas_config.get('arquivo_dados', 'remessas.csv')
            caminho_arquivo = BASE_DIR / arquivo
            
            if not caminho_arquivo.exists():
                logger.error(f"Arquivo de remessas não encontrado: {caminho_arquivo}")
                return []
            
            remessas = []
            with open(caminho_arquivo, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    remessas.append(row)
            
            logger.info(f"Carregadas {len(remessas)} remessas do arquivo")
            return remessas
            
        except Exception as e:
            logger.error(f"Erro ao carregar remessas: {str(e)}")
            return []
    
    def processar_remessa(self, remessa):
        """Processa uma remessa individual"""
        try:
            logger.info(f"Processando remessa: {remessa}")
            
            # Exemplo de fluxo de processamento
            # Este fluxo deve ser adaptado conforme a interface do Bimmer
            
            # 1. Navegar até o menu de remessas
            # self.automation.clicar(x, y)  # Clicar no menu
            
            # 2. Preencher dados da remessa
            # self.automation.digitar(remessa.get('beneficiario', ''))
            
            # 3. Preencher valor
            # self.automation.digitar(remessa.get('valor', ''))
            
            # 4. Confirmar e gerar remessa
            # self.automation.clicar(x, y)  # Clicar em gerar
            
            # TODO: Implementar fluxo específico do Bimmer
            logger.warning("Fluxo de processamento de remessa ainda não implementado")
            logger.info("Por favor, adapte este método conforme a interface do Bimmer")
            
            return True
            
        except Exception as e:
            logger.error(f"Erro ao processar remessa: {str(e)}")
            return False
    
    def executar(self):
        """Executa o bot principal"""
        try:
            logger.info("Iniciando bot do Bimmer...")

            # Por solicitação: não abrir nenhum software manualmente (sem cliques/Win+R).
            logger.info("Não abriremos nenhum software manualmente. Prosseguindo para o login se o Bimer já estiver aberto.")
            
            # Login no Bimer: usar credenciais do config (username/password/account)
            try:
                logger.info("Realizando login no Bimer (usando credenciais do config)...")
                self.automation.login_bimer(max_wait=15.0)
            except Exception as e:
                logger.warning(f"Falha ao executar login automático no Bimer: {e}")
            
            # Fluxo temporário: pular processamento de remessas por enquanto
            # O arquivo ficará na VM e o caminho será informado no fluxo futuro.
            bimmer_cfg = self.config.get('bimmer', {})
            if bimmer_cfg.get('skip_remessas', True):
                logger.info("Modo temporário: pulando processamento de remessas (sem carregar CSV).")
                return True

            # Carrega remessas
            remessas = self.carregar_remessas()
            
            if not remessas:
                logger.warning("Nenhuma remessa encontrada para processar")
                return False
            
            # Processa cada remessa
            sucesso = 0
            falhas = 0
            
            for i, remessa in enumerate(remessas, 1):
                logger.info(f"Processando remessa {i}/{len(remessas)}")
                
                if self.processar_remessa(remessa):
                    sucesso += 1
                    logger.info(f"Remessa {i} processada com sucesso")
                else:
                    falhas += 1
                    logger.error(f"Falha ao processar remessa {i}")
                
                # Aguarda entre remessas
                if i < len(remessas):
                    self.automation.aguardar(2)
            
            logger.info(f"Processamento concluído: {sucesso} sucesso(s), {falhas} falha(s)")
            return True
            
        except Exception as e:
            logger.error(f"Erro ao executar bot: {str(e)}")
            return False
        finally:
            logger.info("Bot finalizado")

