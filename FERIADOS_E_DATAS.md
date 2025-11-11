# 📅 Lógica de Feriados e Datas - RPA Bimer

## 🎯 Objetivo

O sistema agora possui lógica inteligente para lidar com **feriados** e **finais de semana**, garantindo que as buscas de remessas considerem todos os dias úteis não processados.

---

## 🔍 Como Funciona

### **Cenário 1: Execução em Dia Útil (Segunda a Sexta, não feriado)**

#### Se o dia anterior também foi útil:
- **Data Início**: Hoje
- **Data Fim**: Hoje
- **Exemplo**: Terça-feira 12/11
  - Busca: `12/11 até 12/11`

#### Se o dia anterior foi fim de semana/feriado:
- **Data Início**: Último dia útil
- **Data Fim**: Hoje
- **Exemplo**: Segunda-feira 17/11 (após fim de semana)
  - Busca: `15/11 até 17/11` (inclui Sexta, Sábado e Domingo)

---

### **Cenário 2: Execução em Fim de Semana/Feriado**

- **Data Início**: Último dia útil
- **Data Fim**: Hoje
- **Exemplo**: Sábado 15/11
  - Busca: `14/11 até 15/11`

---

## 📋 Feriados Configurados

### **Feriados Nacionais Fixos:**
- 01/01 - Ano Novo
- 21/04 - Tiradentes
- 01/05 - Dia do Trabalho
- 07/09 - Independência do Brasil
- 12/10 - Nossa Senhora Aparecida
- 02/11 - Finados
- 15/11 - Proclamação da República
- 20/11 - Consciência Negra
- 25/12 - Natal

### **Feriados Móveis 2025:**
- 03/03 - Carnaval
- 04/03 - Carnaval
- 18/04 - Sexta-feira Santa
- 30/05 - Corpus Christi

⚠️ **IMPORTANTE**: Atualizar feriados móveis anualmente!

---

## 🛠️ Funções Implementadas

### `eh_feriado(data)`
Verifica se uma data é feriado nacional.

### `eh_dia_util(data)`
Verifica se uma data é dia útil (não é fim de semana nem feriado).

### `obter_ultimo_dia_util()`
Retorna o último dia útil antes de hoje.

### `obter_periodo_busca()`
Calcula o período de busca inteligente (data_inicio, data_fim).

### `obter_data_inicio_busca()`
Retorna a data de início no formato `dd/mm/aaaa`.

### `obter_data_fim_busca()`
Retorna a data de fim no formato `dd/mm/aaaa`.

---

## 📊 Exemplos Práticos

### **Exemplo 1: Sexta-feira 14/11/2025**
- Execução: Sexta-feira (dia útil)
- Ontem: Quinta-feira (dia útil)
- **Resultado**: Busca apenas `14/11 até 14/11`

### **Exemplo 2: Sábado 15/11/2025**
- Execução: Sábado (fim de semana)
- Último dia útil: Sexta-feira 14/11
- **Resultado**: Busca `14/11 até 15/11`

### **Exemplo 3: Segunda-feira 17/11/2025**
- Execução: Segunda-feira (dia útil)
- Ontem: Domingo (fim de semana)
- Último dia útil: Sexta-feira 14/11
- **Resultado**: Busca `15/11 até 17/11` ⭐

> ⭐ Este é o caso que você mencionou! A busca inclui Sexta (15/11), Sábado e Domingo.

### **Exemplo 4: Quarta-feira 20/11/2025**
- Execução: Quarta-feira (dia útil)
- Ontem: Terça-feira (feriado - Consciência Negra)
- Último dia útil: Segunda-feira 18/11
- **Resultado**: Busca `19/11 até 20/11`

---

## 🔄 Atualização Anual de Feriados

Para atualizar os feriados móveis de 2026, edite o arquivo `testar_login_bimer.py`:

```python
# Feriados móveis 2026 (atualizar anualmente)
FERIADOS_MOVEIS_2026 = [
    "16/02",  # Carnaval
    "17/02",  # Carnaval
    "03/04",  # Sexta-feira Santa
    "04/06",  # Corpus Christi
]
```

E atualize a função `eh_feriado`:

```python
# Verifica feriados móveis do ano atual
if data.year == 2026 and data_completa in FERIADOS_MOVEIS_2026:
    return True
```

---

## 📝 Logs do Sistema

O sistema agora exibe informações detalhadas sobre o período de busca:

```
📅 INFORMAÇÕES DE DATA:
   • Hoje: 17/11/2025 (Dia útil)
   • Período de busca: 15/11 até 17/11
   ⚠️  Buscando múltiplos dias (incluindo dias não úteis anteriores)
```

E no resumo final:

```
📊 RESUMO DA EXECUÇÃO:
   ✓ Login realizado
   📅 Data de execução: 17/11/2025 (Dia útil)
   🔍 Período de busca: 15/11 até 17/11
   
   🏢 EMPRESA 1:
      ✓ Período: 15/11 até 17/11
```

---

## ✅ Benefícios

1. **Automático**: Não precisa configurar datas manualmente
2. **Inteligente**: Considera feriados e finais de semana
3. **Completo**: Não perde nenhum dia útil
4. **Transparente**: Logs claros sobre o período de busca
5. **Flexível**: Fácil adicionar feriados municipais/estaduais

---

## 🚀 Próximos Passos

Para adicionar **feriados municipais** (ex: São Paulo):

```python
# Feriados municipais de São Paulo
FERIADOS_MUNICIPAIS_SP = [
    "25/01",  # Aniversário de São Paulo
    "09/07",  # Revolução Constitucionalista
]
```

E atualizar a função `eh_feriado` para incluí-los.
