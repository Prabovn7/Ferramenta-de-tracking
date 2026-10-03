<div align="center">

# Tracking de Cases

**Distribuição equilibrada, acompanhamento operacional e visão gerencial em Excel.**

Ferramenta local desenvolvida em Python para organizar uma base de cases, distribuir o trabalho entre responsáveis e gerar uma tracking com dashboard gerencial e relatórios.

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-XLSM-217346?logo=microsoft-excel&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D4?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

</div>

---

![Painel do gestor](docs/imagens/painel_gestor.png)

> Demonstração utilizando dados inteiramente fictícios.

## Sobre o projeto

O **Tracking de Cases** foi desenvolvido para facilitar a distribuição e o acompanhamento diário de demandas operacionais.

A ferramenta recebe uma base em Excel, valida os campos necessários, organiza os cases pela data de abertura e realiza uma distribuição equilibrada entre os responsáveis cadastrados.

Ao final, é gerada uma nova planilha contendo:

- tracking operacional;
- responsáveis por case;
- controle de status;
- motivos e observações;
- acompanhamento de escalonamentos;
- dashboard gerencial;
- indicadores por responsável;
- alertas de preenchimento;
- geração de resumo em PDF.

A aplicação funciona localmente e não depende de integrações externas para realizar a distribuição.

---

## Fluxo

```text
Base .xlsx
    ↓
Validação dos dados
    ↓
Cases em espera
    ↓
Ordenação por abertura
    ↓
Distribuição equilibrada
    ↓
Tracking .xlsm
    ↓
Dashboard do gestor
    ↓
Resumo em PDF
```

---

## Principais funcionalidades

### Distribuição automática

Os cases são ordenados pela data e hora de abertura, do mais antigo para o mais recente.

A distribuição ocorre de forma alternada entre os responsáveis informados.

A diferença inicial entre as quantidades atribuídas é de, no máximo, um case.

### Tracking operacional

A planilha gerada utiliza as seguintes colunas:

```text
CASE
DATA DE ABERTURA
DATA DE MODIFICAÇÃO
STATUS
MOTIVO
OBSERVAÇÃO
RESPONSÁVEL
```

Os campos de acompanhamento podem ser atualizados pela equipe durante o dia.

### Dashboard gerencial

A aba **Gestor** oferece uma visão consolidada da tracking.

É possível visualizar toda a equipe ou selecionar um responsável específico.

O painel apresenta:

| Indicador | Descrição |
|---|---|
| Distribuídos | Total atribuído |
| Concluídos | Cases finalizados |
| Ainda abertos | Cases não concluídos |
| Pendentes | Ainda não iniciados |
| Em andamento | Atualmente em tratamento |
| Revisar depois | Cases marcados para retomada |
| Escalados | Cases aguardando apoio ou retorno |
| Conclusão | Percentual concluído |
| Mais antigo aberto | Data do case aberto mais antigo |

O dashboard também identifica inconsistências como status inválidos, responsáveis desconhecidos e cases escalados sem observação.

---

## Interface

<table>
<tr>
<td align="center"><strong>Dashboard</strong></td>
<td align="center"><strong>Tracking</strong></td>
</tr>
<tr>
<td><img src="docs/imagens/painel_gestor.png" alt="Dashboard"></td>
<td><img src="docs/imagens/tracking.png" alt="Tracking"></td>
</tr>
</table>

---

## Status

| Status | Significado |
|---|---|
| **Pendente** | Ainda não iniciado |
| **Em andamento** | Em tratamento |
| **Concluído** | Finalizado |
| **Revisar depois** | Deve ser retomado posteriormente |
| **Escalado** | Aguarda apoio ou retorno |

Para cases escalados, o destino e o contexto podem ser informados em **OBSERVAÇÃO**.

Exemplo:

```text
Escalado para: Equipe Financeira — aguardando retorno.
```

O escalonamento é registrado apenas na tracking. A aplicação não executa ações em sistemas externos.

---

## Tecnologias

| Tecnologia | Finalidade |
|---|---|
| Python 3.10+ | Processamento e automação |
| openpyxl | Manipulação das planilhas |
| Excel / XLSM | Tracking e dashboard |
| VBA | Exportação do painel para PDF |
| FPDF | Geração alternativa de relatórios |

As dependências utilizadas pela versão operacional acompanham o programa em `Programa/bibliotecas`.

A execução não instala pacotes, não utiliza rede e não solicita privilégios administrativos.

---

## Início rápido

### 1. Baixe o projeto

Clone ou baixe o repositório e mantenha sua estrutura de pastas.

### 2. Prepare a base

A entrada deve ser um arquivo `.xlsx`.

Os seguintes campos são reconhecidos:

| Informação | Cabeçalhos |
|---|---|
| Case | `CASE` ou `Número do caso` |
| Abertura | `DATA DE ABERTURA` |
| Modificação | `DATA DE MODIFICAÇÃO` ou `Modificado em` |

Outras colunas podem existir normalmente e são ignoradas durante a distribuição.

### 3. Configure cases em espera

Caso algum case não deva entrar na distribuição atual, adicione seu número em:

```text
Programa/Cases_em_espera.txt
```

Um número por linha.

### 4. Execute

Abra:

```text
Programa/Iniciar.bat
```

Informe o caminho da base e os nomes dos responsáveis.

### 5. Abra a tracking

O resultado será criado em:

```text
Programa/Saidas/<execucao>/
```

com o nome:

```text
Tracking_DD-MM.xlsm
```

---

## Exemplo de distribuição

Com três responsáveis:

```text
Case 01 → Ana
Case 02 → Bruno
Case 03 → Carla
Case 04 → Ana
Case 05 → Bruno
Case 06 → Carla
```

A alternância continua mesmo quando a data de abertura muda.

Em casos com a mesma data e horário, a ordem original da base é preservada.

A distribuição equilibra quantidade, não complexidade.

---

## Teste com dados fictícios

O projeto pode incluir:

```text
exemplos/Base_Ficticia_Cases.xlsx
```

com registros criados exclusivamente para demonstração.

Exemplo:

```powershell
python Programa/distribuir_cases.py "exemplos/Base_Ficticia_Cases.xlsx" --analistas "Ana" "Bruno" "Carla" "Diego" "Elisa" --em-espera "exemplos/exclusoes_demo.txt" --sem-pausa
```

Consulte também:

[exemplos/COMO_TESTAR.md](exemplos/COMO_TESTAR.md)

---

## PDF

O dashboard pode ser exportado para PDF diretamente pelo Excel.

No aplicativo desktop:

```text
Gestor → GERAR RESUMO PDF
```

Os arquivos são gravados em:

```text
Relatorios_PDF/
```

Existe também uma alternativa em Python:

```text
Programa/Gerar_PDF.bat
```

---

## Uso compartilhado

O arquivo final pode ser armazenado em uma solução compatível com edição colaborativa.

Para evitar divergências, todos os usuários devem trabalhar sobre o mesmo arquivo.

A aplicação não realiza upload, autenticação ou configuração de permissões automaticamente.

---

## Estrutura

```text
Ferramenta-de-tracking/
│
├── README.md
├── LICENSE
│
├── docs/
│   ├── Guia_Tracking_Cases.docx
│   ├── Apresentacao_Tracking_Cases.pptx
│   └── imagens/
│       ├── painel_gestor.png
│       └── tracking.png
│
├── exemplos/
│   ├── Base_Ficticia_Cases.xlsx
│   ├── exclusoes_demo.txt
│   └── COMO_TESTAR.md
│
└── Programa/
    ├── Iniciar.bat
    ├── Atualizar_Painel.bat
    ├── Gerar_PDF.bat
    ├── distribuir_cases.py
    ├── painel_gestor.py
    ├── dashboard_gestor.py
    ├── atualizar_painel.py
    ├── gerar_relatorio.py
    ├── incluir_botao_pdf.py
    ├── Cases_em_espera.txt
    ├── bibliotecas/
    └── recursos/
```

---

## Limitações atuais

A versão atual não possui autenticação, controle de acesso por responsável, registro automático de início e pausa, cálculo automático do tempo de tratamento, sincronização com sistemas externos ou execução agendada.

O dashboard é alimentado pelas atualizações realizadas na própria tracking.

---

## Segurança dos dados

O repositório público deve conter somente código, documentação, recursos da aplicação e dados fictícios.

Bases reais, trackings geradas, relatórios e listas operacionais não devem ser versionados.

---

## Licença

Este projeto utiliza a licença **MIT**.

Consulte [LICENSE](LICENSE).

---

<div align="center">

Desenvolvido por **Pablo Santos**

**Python · Excel · Automation**

</div>