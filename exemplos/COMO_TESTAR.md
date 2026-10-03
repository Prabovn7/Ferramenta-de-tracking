# Teste com a base fictícia

Todos os 42 registros de `Base_Ficticia_Cases.xlsx` foram criados do zero. A base contém apenas identificadores numéricos inventados, datas de exemplo, equipes e pessoas explicitamente fictícias. Não contém CPF, e-mail, telefone, endereço, nome de cliente ou número de case da operação.

## Teste pelo programa

1. Abra `Programa/Iniciar.bat`.
2. Cole o caminho completo de `exemplos/Base_Ficticia_Cases.xlsx`.
3. Informe cinco pessoas: Ana, Bruno, Carla, Diego e Elisa.
4. Sem lista de exclusões, o resultado será de 42 cases: 9, 9, 8, 8 e 8 por pessoa.

Para testar também as exclusões, copie **somente os números fictícios** de `exclusoes_demo.txt` para `Programa/Cases_em_espera.txt` antes de executar. Depois do teste, deixe a lista do programa vazia ou configure a lista adequada ao próximo uso.

## Teste pela linha de comando

Na raiz do projeto, execute:

```powershell
python Programa/distribuir_cases.py "exemplos/Base_Ficticia_Cases.xlsx" --analistas "Ana" "Bruno" "Carla" "Diego" "Elisa" --em-espera "exemplos/exclusoes_demo.txt" --sem-pausa
```

Resultado esperado com a lista de demonstração:

| Conferência | Resultado |
| --- | --- |
| Recebidos | 42 |
| Em espera | 2, os cases fictícios 00090041 e 00090042 |
| Número da lista ausente | 00099999 |
| Distribuídos | 40 |
| Quantidade por pessoa | 8 para cada um dos cinco nomes |
| Status inicial | Pendente em todos os registros |
| Ordem | Abertura crescente, alternando os cinco nomes |

Na Tracking, mude um case para Concluído e outro para Escalado. Preencha o destino do escalonamento em OBSERVAÇÃO. O gestor deverá mostrar 1 concluído, 39 ainda abertos e 1 escalado. Esse escalado já faz parte dos 39 abertos.

Escolha Ana no filtro do gestor para ver seus 8 cases. Volte a Toda a equipe e teste o botão de PDF no Excel instalado, se as macros forem permitidas.

As imagens do guia e da apresentação mostram uma simulação posterior: 20 concluídos, 8 em andamento, 5 pendentes, 4 para revisão e 3 escalados. Esses status ilustram o uso durante o dia e não são o estado inicial gerado pelo distribuidor.

## Conferência do SLA

Use uma cópia de teste da tracking. Em um case aberto, informe abertura em **25/09/2026 08:00:00**. Sem feriados, **PRAZO ATÉ** deve mostrar **01/10/2026 23:59:59**.

Cadastre **28/09/2026** como feriado fictício em `SLA!K10`. O prazo deve mudar para **02/10/2026 23:59:59**. Retire o feriado após essa conferência.

Os indicadores usam a data atual do computador. Para uma consulta em 01/10/2026, o prazo de 01/10 aparece como **Vence hoje**. A partir de 02/10, fica **Vencido**, enquanto aberto. Pressione **F9** e confira o instante do cálculo no painel.

Verifique que prazo após hoje + vencem hoje + vencidos + alerta de datas = ainda abertos no filtro escolhido. Concluídos saem dessa soma. Escalados e cases para revisão continuam nela.

Na cópia de teste, uma abertura futura ou vazia deve aparecer no alerta. Corrija-a depois. A tabela de detalhes da aba SLA mostra o lote completo e possui filtros próprios.

## Atualização sem redistribuir

1. Salve os status, motivos e observações na cópia de teste.
2. Cadastre um feriado fictício válido em `SLA!K10`.
3. Execute `Programa/Atualizar_Painel.bat` com o `.xlsm` salvo.
4. Compare a nova cópia em `Painel_SLA_<execucao>` com o arquivo de entrada.

Cases, responsáveis, status, motivos, observações, Resumo e feriados devem permanecer preservados. A entrada deve continuar intacta. A rotina também aceita uma tracking compatível `.xlsx`.

## Duas opções de PDF

O botão **GERAR RESUMO PDF** no Excel desktop usa a visão atual, respeita o filtro do Gestor e inclui o SLA em outra página. Confira a saída em `Relatorios_PDF`.

**Gerar_PDF.bat** usa todos os registros salvos e gera um resumo operacional, sem o filtro do Gestor e sem o bloco de SLA.

As imagens de SLA usam uma simulação com data fixa. Não representam os vencimentos calculados na data em que você executar este roteiro. O alias `--exclusoes` continua aceito para compatibilidade.
