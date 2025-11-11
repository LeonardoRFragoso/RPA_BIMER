"""
Script para conectar à VM Windows via RDP
"""
import sys
from pathlib import Path

# Adiciona o diretório raiz ao path
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

from src.rdp_connection import RDPConnection
from src.logger import logger

def main():
    """Função principal para conectar à VM"""
    try:
        print("=" * 50)
        print("RPA Bimmer - Conexão RDP com VM")
        print("=" * 50)
        print()
        
        # Cria instância de conexão RDP
        rdp = RDPConnection()
        
        print(f"Host: {rdp.host}:{rdp.port}")
        print(f"Usuário: {rdp.username}")
        print()
        
        # Pergunta método de conexão
        print("Escolha o método de conexão:")
        print("1. Usar arquivo RDP (recomendado)")
        print("2. Conectar com senha automática (cmdkey)")
        print("3. Conectar sem senha (inserir manualmente)")
        print()
        
        escolha = input("Digite sua escolha (1-3) [1]: ").strip() or "1"
        
        if escolha == "1":
            print("\nCriando arquivo RDP e conectando...")
            rdp.conectar_rdp(usar_arquivo=True)
        elif escolha == "2":
            print("\nArmazenando credenciais e conectando...")
            rdp.conectar_com_senha()
        elif escolha == "3":
            print("\nConectando (será necessário inserir senha)...")
            rdp.conectar_rdp(usar_arquivo=False)
        else:
            print("Opção inválida!")
            sys.exit(1)
        
        print("\nConexão iniciada! Aguarde a janela do RDP abrir.")
        print("\nNota: Se optou pela opção 2, as credenciais foram armazenadas.")
        print("      Para removê-las, execute: python conectar_vm.py --remover-credenciais")
        
    except KeyboardInterrupt:
        print("\nOperação cancelada pelo usuário")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Erro ao conectar: {str(e)}", exc_info=True)
        print(f"\nErro: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--remover-credenciais':
        # Remove credenciais armazenadas
        rdp = RDPConnection()
        rdp.remover_credenciais()
    else:
        main()

