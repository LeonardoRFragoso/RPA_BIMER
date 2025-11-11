"""
Módulo de logging para o RPA Bimmer
"""
import logging
import os
from pathlib import Path
from datetime import datetime
from logging.handlers import RotatingFileHandler
from src.config import config, BASE_DIR

def setup_logger():
    """Configura o logger do sistema"""
    log_config = config.get('logging', {})
    log_level = getattr(logging, log_config.get('nivel', 'INFO'))
    
    # Caminho completo do arquivo de log
    log_file_path = log_config.get('arquivo_log', 'logs/rpa_bimmer.log')
    log_file = BASE_DIR / log_file_path
    
    # Cria diretório de logs se não existir
    log_dir = log_file.parent
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # Configura o logger
    logger = logging.getLogger('RPA_Bimmer')
    logger.setLevel(log_level)
    
    # Evita duplicação de handlers
    if logger.handlers:
        return logger
    
    # Handler para arquivo com rotação
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(log_level)
    
    # Handler para console
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    
    # Formato das mensagens
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# Logger global
logger = setup_logger()

