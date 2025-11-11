"""
Script para conectar automaticamente à VM Windows via RDP
Sem necessidade de interação do usuário
"""
import sys
from pathlib import Path

# Adiciona o diretório raiz ao path
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

from src.rdp_connection import RDPConnection
from src.rdp_auto_fill import RDPAutoFill
from src.logger import logger
import time

def main():
    """Função principal para conectar à VM automaticamente"""
    try:
        print("=" * 50)
        print("RPA Bimmer - Conexão RDP Automática com VM")
        print("=" * 50)
        print()
        
        # Cria instância de conexão RDP
        rdp = RDPConnection()
        
        print(f"Host: {rdp.host}:{rdp.port}")
        print(f"Usuário: {rdp.username}")
        print()
        
        # Tenta diferentes métodos de conexão automaticamente
        print("Tentando conectar automaticamente...")
        print()
        
        # Método 1: Tentar com cmdkey primeiro (melhor método - senha automática)
        print("1. Tentando conexão com senha automática (cmdkey)...")
        if rdp.conectar_com_senha():
            print("✓ Conexão RDP iniciada com senha automática")
            print("\nAguarde a janela do RDP abrir.")
            print("\nTentando preencher credenciais automaticamente se necessário...")
            
            # Aguarda um pouco para a janela abrir
            time.sleep(2)
            
            # Tenta preencher credenciais automaticamente se a tela aparecer
            try:
                auto_fill = RDPAutoFill()
                if auto_fill.preencher_automaticamente(timeout=5):
                    print("✓ Credenciais preenchidas automaticamente!")
                else:
                    print("\n⚠ Se o Windows pedir credenciais manualmente:")
                    print(f"   - Usuário: {rdp.username}")
                    print(f"   - Senha: (a senha configurada no config.yaml)")
                    print("   - Marque 'Lembrar-me' para não pedir novamente")
            except Exception as e:
                logger.warning(f"Não foi possível preencher credenciais automaticamente: {e}")
                print("\n⚠ Se o Windows pedir credenciais manualmente:")
                print(f"   - Usuário: {rdp.username}")
                print(f"   - Senha: (a senha configurada no config.yaml)")
                print("   - Marque 'Lembrar-me' para não pedir novamente")
            
            return True
        
        # Método 2: Se falhar, tentar com arquivo RDP
        print("\n2. Tentando conexão via arquivo RDP...")
        if rdp.conectar_rdp(usar_arquivo=True):
            print("✓ Conexão RDP iniciada via arquivo")
            print("\nAguarde a janela do RDP abrir.")
            return True
        
        # Método 3: Última tentativa sem senha
        print("\n3. Tentando conexão sem senha (será necessário inserir manualmente)...")
        if rdp.conectar_rdp(usar_arquivo=False):
            print("✓ Conexão RDP iniciada (insira a senha quando solicitado)")
            print("\nAguarde a janela do RDP abrir.")
            return True
        
        print("\n✗ Não foi possível iniciar a conexão RDP")
        print("Verifique as configurações em config.yaml")
        return False
        
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
        print("Removendo credenciais armazenadas...")
        rdp = RDPConnection()
        if rdp.remover_credenciais():
            print("✓ Credenciais removidas com sucesso")
        else:
            print("✗ Erro ao remover credenciais")
    else:
        sucesso = main()
        sys.exit(0 if sucesso else 1)

