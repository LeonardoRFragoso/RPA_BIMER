"""
Script auxiliar para capturar coordenadas do mouse e identificar elementos da interface
Útil para mapear elementos do Bimmer para automação
"""
import pyautogui
import time
import sys

print("=" * 60)
print("Capturador de Coordenadas - RPA Bimmer")
print("=" * 60)
print()
print("INSTRUÇÕES:")
print("1. Posicione o mouse sobre o elemento que deseja mapear")
print("2. Pressione ENTER para capturar a coordenada")
print("3. Digite um nome para o elemento")
print("4. Pressione 'q' e ENTER para sair")
print()
print("Pressione ENTER para começar...")
input()

pyautogui.FAILSAFE = True
elementos = {}

try:
    while True:
        print("\n" + "-" * 60)
        print("Posicione o mouse sobre o elemento e pressione ENTER")
        print("Ou digite 'q' e pressione ENTER para sair")
        print("-" * 60)
        
        entrada = input("> ").strip().lower()
        
        if entrada == 'q':
            break
        
        # Captura a posição atual do mouse
        x, y = pyautogui.position()
        
        print(f"\nCoordenada capturada: ({x}, {y})")
        nome = input("Digite um nome para este elemento: ").strip()
        
        if nome:
            elementos[nome] = {'x': x, 'y': y}
            print(f"✓ Elemento '{nome}' salvo em ({x}, {y})")
        
        time.sleep(0.5)
    
    # Exibe todos os elementos capturados
    print("\n" + "=" * 60)
    print("ELEMENTOS CAPTURADOS:")
    print("=" * 60)
    
    if elementos:
        for nome, coord in elementos.items():
            print(f"{nome}: ({coord['x']}, {coord['y']})")
        
        print("\n" + "=" * 60)
        print("CÓDIGO PARA USAR NO BOT:")
        print("=" * 60)
        print()
        print("# Adicione ao config.yaml:")
        print("ui_elements:")
        for nome, coord in elementos.items():
            print(f"  {nome}: \"{coord['x']},{coord['y']}\"")
        print()
        print("# Ou use diretamente no código:")
        for nome, coord in elementos.items():
            print(f"self.automation.clicar({coord['x']}, {coord['y']})  # {nome}")
    else:
        print("Nenhum elemento foi capturado.")
    
    print("\n" + "=" * 60)
    
except KeyboardInterrupt:
    print("\n\nOperação cancelada pelo usuário.")
except Exception as e:
    print(f"\nErro: {str(e)}")

