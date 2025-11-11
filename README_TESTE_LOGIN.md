# 🧪 Script de Teste - Login Bimer

## 📋 Descrição

Script standalone para testar o fluxo de login no Bimer **diretamente dentro da VM**, sem precisar de conexão RDP. Ideal para desenvolvimento e debug rápido.

---

## 🚀 Como Usar

### 1️⃣ **Copie o script para a VM**

Transfira o arquivo `testar_login_bimer.py` para a VM onde o Bimer está instalado.

### 2️⃣ **Instale as dependências na VM**

```bash
pip install pyautogui pywin32
```

ou

```bash
pip install pyautogui pyperclip
```

### 3️⃣ **Abra o Bimer na VM**

- Certifique-se de que o Bimer está aberto
- A tela de login deve estar visível
- Posicione a janela do Bimer de forma que fique visível

### 4️⃣ **Execute o script**

```bash
python testar_login_bimer.py
```

### 5️⃣ **Não mova o mouse**

O script irá:
1. Selecionar ambiente TESTE no dropdown
2. Preencher a senha
3. Clicar em Entrar
4. Executar ações pós-login (se configuradas)

---

## ⚙️ Configuração

### **Coordenadas Atuais**

```python
# Dropdown de ambiente
DROPDOWN_AMBIENTE_X = 866
DROPDOWN_AMBIENTE_Y = 579
AMBIENTE_TESTE_X = 974
AMBIENTE_TESTE_Y = 677

# Campo de senha
CAMPO_SENHA_X = 904
CAMPO_SENHA_Y = 520

# Botão Entrar
BOTAO_ENTRAR_X = 511
BOTAO_ENTRAR_Y = 682
```

### **Como Atualizar Coordenadas**

Se as coordenadas mudarem, use o script `capturar_coordenadas.py`:

```bash
python capturar_coordenadas.py
```

Depois, atualize as constantes no início do arquivo `testar_login_bimer.py`.

---

## 🔧 Adicionar Cliques Pós-Login

Para adicionar ações após o login, edite a lista `CLIQUES_POS_LOGIN`:

```python
CLIQUES_POS_LOGIN = [
    ("menu_financeiro", 100, 200, None),
    ("a_pagar", 150, 250, None),
    ("campo_conta", 300, 400, "19"),  # Clica e digita "19"
    ("campo_layout", 300, 450, "16"), # Clica e digita "16"
]
```

**Formato:**
```python
("nome_descritivo", coordenada_x, coordenada_y, "texto_para_digitar_ou_None")
```

---

## 📊 Logs Detalhados

O script mostra logs detalhados de cada passo:

```
============================================================
[PASSO 1/4] SELECIONANDO AMBIENTE TESTE NO DROPDOWN
============================================================
→ Clicando no dropdown em (866, 579)
⏳ Aguardando dropdown abrir...
→ Clicando em 'TESTE' em (974, 677)
⏳ Ambiente selecionado
✓ Ambiente TESTE selecionado com sucesso

============================================================
[PASSO 2/4] PREENCHENDO SENHA
============================================================
→ Clicando no campo de senha em (904, 520)
⏳ Campo focado
→ Limpando campo (Ctrl+A)
→ Colando senha (9 caracteres)
✓ Texto copiado para área de transferência
✓ Senha colada com sucesso
```

---

## 🐛 Troubleshooting

### **Problema: "Coordenadas não configuradas"**
- Capture as coordenadas usando `capturar_coordenadas.py`
- Atualize as constantes no início do script

### **Problema: "Senha não foi colada"**
- Verifique se `pywin32` ou `pyperclip` está instalado
- O script tentará digitar caractere por caractere como fallback

### **Problema: "Clicou no lugar errado"**
- A resolução da tela pode ter mudado
- Recapture as coordenadas
- Certifique-se de que a janela do Bimer está na mesma posição

### **Problema: "FailSafeException"**
- Você moveu o mouse para o canto superior esquerdo
- Isso é uma proteção do PyAutoGUI
- Execute novamente e não mova o mouse

---

## 🎯 Vantagens deste Script

✅ **Rápido** - Testa apenas o login, sem conexão RDP  
✅ **Isolado** - Não depende de outros módulos  
✅ **Logs claros** - Mostra exatamente o que está acontecendo  
✅ **Fácil debug** - Rode diretamente na VM  
✅ **Extensível** - Adicione novos passos facilmente  

---

## 📝 Próximos Passos

1. ✅ Teste o login básico
2. 📸 Capture coordenadas de elementos pós-login
3. ➕ Adicione cliques em `CLIQUES_POS_LOGIN`
4. 🔄 Teste a sequência completa
5. 📋 Documente o fluxo final

---

## 🔗 Arquivos Relacionados

- `capturar_coordenadas.py` - Captura coordenadas de elementos
- `config.yaml` - Configuração principal do bot
- `src/ui_automation.py` - Módulo de automação completo
- `main.py` - Script principal com RDP

---

**Desenvolvido para acelerar o desenvolvimento do RPA-Bimmer** 🚀
