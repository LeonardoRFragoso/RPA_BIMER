# Guia de Próximos Passos - RPA Bimmer

## ✅ Status Atual

- [x] Projeto criado e estruturado
- [x] Dependências instaladas
- [x] Conexão RDP com VM funcionando
- [x] Bimmer aberto na VM
- [ ] Caminho do executável verificado
- [ ] Módulo de remessas bancárias identificado
- [ ] Elementos da interface mapeados
- [ ] Fluxo de automação implementado
- [ ] Testes realizados
- [ ] Serviço Windows instalado

## 📋 Próximos Passos

### 1. Verificar Caminho do Executável do Bimmer

Na VM, verifique o caminho exato do executável do Bimmer:

1. Abra o Gerenciador de Tarefas (Ctrl+Shift+Esc)
2. Procure pelo processo "Bimmer.exe" ou similar
3. Clique com botão direito > "Abrir local do arquivo"
4. Copie o caminho completo

Atualize o `config.yaml` se necessário:
```yaml
bimmer:
  executable_path: "C:\\Program Files\\Alterdata\\Bimmer\\Bimmer.exe"  # Verifique este caminho
```

### 2. Identificar o Módulo de Remessas Bancárias

No Bimmer, navegue até o módulo de remessas bancárias:

1. Procure no menu lateral ou na barra de ferramentas
2. Pode estar em:
   - Menu "Bancos" ou "Financeiro"
   - Menu "Remessas" ou "Transferências"
   - Aba "Ferramentas" ou "Relatórios"
3. Anote o caminho completo para chegar até lá

### 3. Mapear Elementos da Interface

Use o script `capturar_coordenadas.py` para mapear os elementos:

```bash
python capturar_coordenadas.py
```

**Elementos que precisam ser mapeados:**
- Menu/ícone para acessar remessas bancárias
- Campo de beneficiário
- Campo de valor
- Campo de banco/agência/conta
- Botão "Gerar" ou "Confirmar"
- Botão "Salvar" ou "Finalizar"
- Qualquer botão de confirmação ou fechamento

### 4. Criar Arquivo de Remessas de Teste

Crie um arquivo `remessas.csv` com dados de teste:

```csv
beneficiario,valor,banco,agencia,conta,observacao
João Silva,1000.00,001,1234,567890,Remessa teste 1
Maria Santos,2500.50,237,5678,123456,Remessa teste 2
```

### 5. Implementar o Fluxo de Automação

Edite o método `processar_remessa()` em `src/bot.py`:

```python
def processar_remessa(self, remessa):
    """Processa uma remessa individual"""
    try:
        logger.info(f"Processando remessa: {remessa}")
        
        # 1. Navegar até o menu de remessas
        # self.automation.clicar(x, y)  # Substitua pelas coordenadas
        
        # 2. Clicar em "Nova Remessa" ou similar
        # self.automation.clicar(x, y)
        
        # 3. Preencher beneficiário
        # self.automation.clicar(x, y)  # Campo beneficiário
        # self.automation.digitar(remessa.get('beneficiario', ''))
        
        # 4. Preencher valor
        # self.automation.pressionar_tecla('tab')
        # self.automation.digitar(remessa.get('valor', ''))
        
        # 5. Preencher banco/agência/conta
        # ... (continue conforme a interface)
        
        # 6. Confirmar e gerar
        # self.automation.clicar(x, y)  # Botão gerar
        
        return True
    except Exception as e:
        logger.error(f"Erro ao processar remessa: {str(e)}")
        return False
```

### 6. Testar o Bot

Execute o bot em modo de teste:

```bash
python main.py
```

**Dicas para testes:**
- Use dados de teste primeiro
- Execute uma remessa por vez inicialmente
- Verifique os logs em `logs/rpa_bimmer.log`
- Capture screenshots se houver erros

### 7. Instalar como Serviço Windows (Produção)

Quando o bot estiver funcionando corretamente:

1. Copie todo o projeto para a VM
2. Instale as dependências na VM
3. Execute como administrador:
   ```bash
   python install_service.py install
   python install_service.py start
   ```

## 🔧 Ferramentas Úteis

### Para Identificar Elementos:

1. **Inspect.exe** (Windows SDK)
   - Identifica elementos UI Automation
   - Útil para encontrar IDs e propriedades

2. **PyAutoGUI - Mouse Info**
   ```python
   import pyautogui
   pyautogui.mouseInfo()  # Mostra coordenadas em tempo real
   ```

3. **Screenshots**
   - Capture telas do Bimmer
   - Use `pyautogui.locateOnScreen()` para encontrar elementos por imagem

### Para Debug:

1. **Logs detalhados**
   - Configure `LOG_LEVEL: DEBUG` no `config.yaml`
   - Verifique `logs/rpa_bimmer.log`

2. **Screenshots automáticas**
   - O bot pode capturar screenshots em caso de erro
   - Use `self.automation.capturar_tela()` no código

## 📝 Checklist de Implementação

- [ ] Caminho do executável verificado
- [ ] Módulo de remessas identificado
- [ ] Menu de remessas mapeado
- [ ] Campos de formulário mapeados
- [ ] Botões de ação mapeados
- [ ] Fluxo completo testado manualmente
- [ ] Fluxo implementado no código
- [ ] Teste com uma remessa
- [ ] Teste com múltiplas remessas
- [ ] Tratamento de erros implementado
- [ ] Logs verificados
- [ ] Serviço instalado e testado

## ⚠️ Observações Importantes

1. **Teste sempre em ambiente de desenvolvimento primeiro**
2. **Use dados de teste antes de processar remessas reais**
3. **Mantenha backups dos arquivos de remessas**
4. **Monitore os logs regularmente**
5. **Valide os resultados no Bimmer após cada execução**

## 🆘 Em Caso de Problemas

1. Verifique os logs: `logs/rpa_bimmer.log`
2. Capture screenshots dos erros
3. Teste o fluxo manualmente no Bimmer
4. Verifique se as coordenadas ainda estão corretas (pode mudar com atualizações do Bimmer)

