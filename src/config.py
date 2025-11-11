"""
Módulo de configuração do RPA Bimmer
Carrega configurações do arquivo YAML e variáveis de ambiente
"""
import os
import yaml
from pathlib import Path
from dotenv import load_dotenv

# Carrega variáveis de ambiente
load_dotenv()

# Caminho base do projeto
BASE_DIR = Path(__file__).parent.parent

# Carrega configurações do YAML
CONFIG_PATH = BASE_DIR / "config.yaml"

def load_config():
    """Carrega configurações do arquivo YAML"""
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
    
    # Sobrescreve com variáveis de ambiente se existirem
    if os.getenv('VM_HOST'):
        config['vm']['host'] = os.getenv('VM_HOST')
    if os.getenv('VM_USERNAME'):
        config['vm']['username'] = os.getenv('VM_USERNAME')
    if os.getenv('VM_PASSWORD'):
        config['vm']['password'] = os.getenv('VM_PASSWORD')
    if os.getenv('BIMMER_EXECUTABLE_PATH'):
        config['bimmer']['executable_path'] = os.getenv('BIMMER_EXECUTABLE_PATH')
    if os.getenv('LOG_LEVEL'):
        config['logging']['nivel'] = os.getenv('LOG_LEVEL')
    
    return config

# Utilitários adicionais: leitura "somente YAML" (ignora envs)
def load_config_yaml_only():
    """Carrega SOMENTE o YAML, sem sobrescrever com variáveis de ambiente."""
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)

def get_vm_config(prefer_env: bool = True):
    """
    Retorna o bloco vm do config.
    - prefer_env=True: usa config com override por env (padrão atual)
    - prefer_env=False: ignora envs e usa apenas YAML
    """
    cfg = load_config() if prefer_env else load_config_yaml_only()
    return cfg.get('vm', {}) or {}

def get_vm_credentials(prefer_env: bool = True):
    """
    Retorna (username, password) conforme a preferência de fonte.
    """
    vm_cfg = get_vm_config(prefer_env=prefer_env)
    return vm_cfg.get('username', ''), vm_cfg.get('password', '')

# Carrega configurações globalmente
config = load_config()

