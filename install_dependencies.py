"""
Script para instalar dependências com tratamento de erros
"""
import subprocess
import sys
import os

def install_package(package):
    """Instala um pacote e trata erros"""
    try:
        print(f"Instalando {package}...")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", package],
            capture_output=True,
            text=True,
            check=True
        )
        print(f"✓ {package} instalado com sucesso")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ Erro ao instalar {package}: {e.stderr}")
        return False

def main():
    """Instala todas as dependências"""
    print("=" * 50)
    print("Instalando dependências do RPA Bimmer")
    print("=" * 50)
    print()
    
    # Lista de pacotes para instalar
    packages = [
        "pyautogui>=0.9.54",
        "pywinauto>=0.6.8",
        "pywin32",  # Sem versão específica para compatibilidade
        "Pillow>=10.1.0",
        "opencv-python>=4.8.1.78",
        "pynput>=1.7.6",
        "python-dotenv>=1.0.0",
        "schedule>=1.2.0",
        "pyyaml>=6.0.1"
    ]
    
    # Atualiza pip primeiro
    print("Atualizando pip...")
    try:
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "--upgrade", "pip"],
            check=True
        )
        print("✓ pip atualizado\n")
    except:
        print("⚠ Aviso: Não foi possível atualizar pip\n")
    
    # Instala cada pacote
    sucesso = 0
    falhas = 0
    
    for package in packages:
        if install_package(package):
            sucesso += 1
        else:
            falhas += 1
        print()
    
    print("=" * 50)
    print(f"Instalação concluída: {sucesso} sucesso(s), {falhas} falha(s)")
    print("=" * 50)
    
    # Se pywin32 foi instalado, tenta executar o post-install
    if sucesso > 0:
        print("\nExecutando pós-instalação do pywin32...")
        try:
            # Tenta encontrar o script de pós-instalação
            import site
            site_packages = site.getsitepackages()[0]
            post_install_script = os.path.join(site_packages, "pywin32_system32", "pywin32_postinstall.py")
            
            if os.path.exists(post_install_script):
                subprocess.run([sys.executable, post_install_script, "-install"], check=True)
                print("✓ Pós-instalação do pywin32 concluída")
            else:
                # Tenta método alternativo
                try:
                    import win32api
                    print("✓ pywin32 já está configurado corretamente")
                except:
                    print("⚠ Aviso: pywin32 instalado, mas pós-instalação pode ser necessária")
                    print("   Execute manualmente: python -m pywin32_postinstall -install")
        except Exception as e:
            print(f"⚠ Aviso: Não foi possível executar pós-instalação do pywin32: {e}")
            print("   Execute manualmente: python -m pywin32_postinstall -install")
    
    if falhas > 0:
        print("\n⚠ Alguns pacotes falharam na instalação.")
        print("   Tente instalar manualmente os pacotes que falharam.")
        sys.exit(1)
    else:
        print("\n✓ Todas as dependências foram instaladas com sucesso!")
        sys.exit(0)

if __name__ == '__main__':
    main()

