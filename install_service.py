"""
Script para instalar o serviço Windows do RPA Bimmer
Execute como administrador: python install_service.py install
"""
import sys
import os
from pathlib import Path

# Adiciona o diretório raiz ao path
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

try:
    import win32serviceutil
    from src.service import BimmerService
    
    def instalar_servico():
        """Instala o serviço Windows"""
        try:
            print("Instalando serviço RPA Bimmer...")
            win32serviceutil.HandleCommandLine(BimmerService)
            print("Serviço instalado com sucesso!")
            print("\nComandos úteis:")
            print("  python install_service.py start    - Inicia o serviço")
            print("  python install_service.py stop     - Para o serviço")
            print("  python install_service.py remove   - Remove o serviço")
            print("  python install_service.py restart  - Reinicia o serviço")
        except Exception as e:
            print(f"Erro ao instalar serviço: {str(e)}")
            print("\nCertifique-se de executar como administrador!")
            sys.exit(1)
    
    if __name__ == '__main__':
        if len(sys.argv) < 2:
            print("Uso: python install_service.py [install|start|stop|remove|restart]")
            print("\nExemplo para instalar:")
            print("  python install_service.py install")
            sys.exit(1)
        
        instalar_servico()
        
except ImportError as e:
    print(f"Erro: {str(e)}")
    print("\nCertifique-se de ter instalado todas as dependências:")
    print("  pip install -r requirements.txt")
    sys.exit(1)

