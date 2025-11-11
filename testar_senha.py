"""
Script para testar se a senha está correta
Mostra a senha configurada para verificação
"""
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

from src.config import config

print("=" * 60)
print("Verificação de Credenciais - RPA Bimmer")
print("=" * 60)
print()
print("Credenciais configuradas:")
print(f"  Host: {config['vm']['host']}:{config['vm']['port']}")
print(f"  Usuário: {config['vm']['username']}")
print(f"  Senha: {config['vm']['password']}")
print()
print(f"Tamanho da senha: {len(config['vm']['password'])} caracteres")
print()
print("Caracteres da senha (um por linha):")
for i, char in enumerate(config['vm']['password'], 1):
    if char == ' ':
        print(f"  {i}: [ESPAÇO]")
    elif char == '\t':
        print(f"  {i}: [TAB]")
    elif char == '\n':
        print(f"  {i}: [NOVA LINHA]")
    else:
        print(f"  {i}: '{char}' (ord: {ord(char)})")
print()
print("=" * 60)
print("Verifique se a senha está correta!")
print("=" * 60)

