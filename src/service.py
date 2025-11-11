"""
Módulo para criação e gerenciamento do serviço Windows
"""
import os
import sys
import win32serviceutil
import win32service
import servicemanager
from pathlib import Path
from src.bot import BimmerBot
from src.logger import logger

class BimmerService(win32serviceutil.ServiceFramework):
    """Serviço Windows para executar o bot do Bimmer"""
    
    _svc_name_ = "RPA-Bimmer-Service"
    _svc_display_name_ = "RPA Bimmer Service"
    _svc_description_ = "Serviço para automação de remessas bancárias no Bimmer"
    
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32service.CreateEvent(None, 0, 0, None)
        self.bot = None
        
    def SvcStop(self):
        """Para o serviço"""
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32service.SetEvent(self.stop_event)
        logger.info("Serviço sendo parado...")
        
    def SvcDoRun(self):
        """Executa o serviço"""
        try:
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, '')
            )
            
            logger.info("Serviço iniciado")
            self.main()
            
        except Exception as e:
            servicemanager.LogErrorMsg(f"Erro no serviço: {str(e)}")
            logger.error(f"Erro no serviço: {str(e)}")
    
    def main(self):
        """Função principal do serviço"""
        try:
            # Cria instância do bot
            self.bot = BimmerBot()
            
            # Loop principal do serviço
            while True:
                # Verifica se o serviço deve parar
                if win32service.WaitForSingleObject(self.stop_event, 1000) == win32service.WAIT_OBJECT_0:
                    break
                
                # Executa o bot
                logger.info("Executando bot...")
                self.bot.executar()
                
                # Aguarda antes da próxima execução
                # Pode ser configurado para executar em intervalos específicos
                import time
                time.sleep(3600)  # Executa a cada hora (ajustar conforme necessário)
                
        except Exception as e:
            logger.error(f"Erro no loop principal: {str(e)}")
        finally:
            logger.info("Serviço finalizado")

def main():
    """Função principal para instalar/desinstalar/executar o serviço"""
    if len(sys.argv) == 1:
        # Executa o serviço normalmente
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(BimmerService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        # Instala, desinstala ou executa o serviço
        win32serviceutil.HandleCommandLine(BimmerService)

if __name__ == '__main__':
    main()

