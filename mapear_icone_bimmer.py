"""
Script para mapear a posição exata do ícone do Bimmer na área de trabalho
Execute este script na VM para capturar as coordenadas do ícone
"""
import pyautogui
import time

print("=" * 60)
print("Mapeador de Ícone do Bimmer - Área de Trabalho")
print("=" * 60)
print()
print("INSTRUÇÕES:")
print("1. Certifique-se de que a área de trabalho está visível")
print("2. Posicione o mouse sobre o ícone do Bimmer")
print("3. Pressione ENTER para capturar a coordenada")
print("4. As coordenadas serão salvas para uso no bot")
print()
print("Pressione ENTER para começar...")
input()

# Mostra área de trabalho
pyautogui.hotkey('win', 'd')
time.sleep(1)

print("\n" + "-" * 60)
print("Posicione o mouse sobre o ícone do Bimmer")
print("Pressione ENTER quando estiver pronto...")
print("-" * 60)
input()

# Captura a posição atual do mouse
x, y = pyautogui.position()

print("\n" + "=" * 60)
print(f"Coordenada capturada: ({x}, {y})")
print("=" * 60)
print()
print("CÓDIGO PARA USAR NO BOT:")
print("=" * 60)
print()
print("# Adicione ao config.yaml:")
print("bimmer:")
print(f"  icon_x: {x}")
print(f"  icon_y: {y}")
print()
print("# Ou use diretamente no código:")
print(f"pyautogui.doubleClick({x}, {y})  # Ícone do Bimmer")
print()
print("=" * 60)
print("Coordenadas salvas! Use essas coordenadas no bot.")
print("=" * 60)

