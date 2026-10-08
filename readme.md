# Central de Vistorias - Distribuição Automática Inteligente

Sistema em Python para alocação e distribuição otimizada de sinistros/vistorias entre técnicos de campo e estagiários, respeitando limites operacionais, janelas de horário e regras de negócio.

---

## 📌 Funcionalidades

- **Distribuição Inteligente Balancada**: Algoritmo com sistema de pesos para priorizar horários críticos, preservar janelas curtas e balancear a carga entre a equipe.
- **Leitura Multi-Formato**: Suporte a arquivos de entrada `.xlsx`, `.xlsm`, `.csv` e `.txt`.
- **Tratamento de Exceções**:
  - Bloqueio automático de sinistros duplicados.
  - Identificação e separação de horários e sinistros inválidos.
  - Sinalização de demandas não distribuídas por falta de capacidade.
- **Relatório Completo em Excel**:
  - **Painel**: Visão geral dos indicadores de desempenho (KPIs).
  - **Agenda**: Lista sequencial dos atendimentos por técnico e horário.
  - **Grade**: Tabela visual estilo calendário por técnico.
  - **Resumo**: Ocupação e saldo de capacidade por técnico.
  - **Capacidade**: Análise da demanda vs. capacidade por slot de horário.
  - **Auditoria**: Log completo da pontuação/score e decisores da distribuição.
  - **Não Distribuídas / Duplicados / Linhas Inválidas**: Abas dedicadas ao saneamento de inconsistências.

---

## 🛠️ Requisitos e Instalação

### Pré-requisitos
- Python 3.8 ou superior instalado.

### Instalando as Dependências

Instale as bibliotecas necessárias executando o comando abaixo na raiz do projeto:

```bash
pip install -r requirements.txt
```

---

## 🚀 Como Usar

Execute o script fornecendo o arquivo de entrada e o nome do arquivo de saída desejado:

```bash
python distribuir.py Entrada.xlsx saida.xlsx
```

> **Nota:** Se os nomes dos arquivos não forem informados, o sistema tentará utilizar por padrão os nomes `Entrada.xlsx` para leitura e `saida.xlsx` para a saída gerada.

---

## 📥 Estrutura do Arquivo de Entrada

O arquivo de entrada (seja `.xlsx`, `.csv` ou `.txt`) deve conter as colunas organizadas na seguinte ordem:

| Coluna | Campo | Descrição |
|---|---|---|
| **1** | `SINISTRO` | Identificador único da vistoria / sinistro. |
| **2** | `HORÁRIO` | Horário agendado (Formatos aceitos: `HH:MM`, `HHhMM`, `HH.MM`). |
| **3** | `ASSISTÊNCIA_COD` | Código da assistência ou identificador auxiliar. |

---

## ⚙️ Regras de Negócio e Configurações

1. **Limite por Técnico**: Cada técnico pode receber no máximo **2 vistorias** por slot de horário.
2. **Restrição por Tipo**:
   - Técnicos com tipo `Estagiário` possuem restrições de horário operacional conforme cadastro.
   - O horário `16:10` é reservado exclusivamente para técnicos da categoria `Integral`.
3. **Parâmetros Alteráveis (`CONFIG` no código)**:
   - `bloquear_duplicados`: Evita que o mesmo sinistro entre duas vezes.
   - `respeitar_active`: Considera apenas técnicos ativos.
   - `preservar_janelas_restritas`: Prioriza guardar a disponibilidade de quem tem turno mais curto.
