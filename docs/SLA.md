# Acompanhamento do SLA

O prazo é de **4 dias úteis**, até 23:59:59 do quarto dia útil após a data de abertura. O dia da abertura não é contado. Sábados, domingos e as datas cadastradas como feriados são desconsiderados.

Exemplo sem feriados: um case aberto na sexta-feira, 25/09/2026, vence na quinta-feira, 01/10/2026, ao fim do dia. A partir de 02/10/2026 ele fica vencido, se ainda estiver aberto.

## No painel Gestor

1. Escolha `Toda a equipe` ou um responsável no campo Visualizar.
2. Clique em **VER SLA DE 4 DIAS ÚTEIS** ou role até a seção de SLA.
3. Confira os três grupos, que não se sobrepõem:
   - **Prazo após hoje:** cases cujo prazo termina em uma data futura.
   - **Vencem hoje:** cases no último dia do prazo.
   - **SLA vencido:** cases cujo prazo já terminou.
4. Confira também o alerta de aberturas inválidas, vazias ou futuras. Esses registros não entram nos três grupos e precisam de correção na Tracking.

Os três grupos, somados ao alerta de datas, correspondem aos cases ainda abertos no filtro selecionado. Escalados e cases para revisão continuam no SLA. Concluídos ficam fora da contagem de SLA em aberto.

## Prazo por case e feriados

O atalho **VER PRAZO E FAIXA SLA DE CADA CASE** abre a aba SLA. Ela mostra a abertura, o responsável, o status, o prazo final, os dias úteis decorridos e a faixa D0 a D4 ou Vencido. D0 significa que nenhum dia útil entrou na contagem. DIAS ÚTEIS conta os dias úteis do calendário desde o dia seguinte à abertura, incluindo o dia atual quando útil; não mede horas de tratamento.

A lista de detalhes mostra o lote completo. Use os filtros próprios da tabela para escolher um responsável ou uma situação. Não altere as fórmulas dessa lista: os dados operacionais continuam sendo preenchidos na Tracking, que mantém suas sete colunas.

Cadastre os feriados aplicáveis à operação em **SLA!K10:K109**, um por linha, como datas do Excel. A lista é entregue vazia. Enquanto estiver vazia, apenas os fins de semana são descontados. Feriados são preservados ao atualizar novamente uma tracking que já possui esta versão do SLA.

## Atualização e PDF

O Excel recalcula as fórmulas ao abrir, ao editar e ao pressionar F9, com o cálculo automático ativo. O relógio não atualiza continuamente se o arquivo permanecer parado. A data e a hora utilizadas aparecem no painel.

O botão **GERAR RESUMO PDF** mantém o funcionamento anterior e inclui a seção de SLA na área de impressão. O bloco de SLA começa em uma nova página para evitar cortes nos indicadores. As contagens do PDF respeitam o filtro do gestor. O botão exige Excel instalado e macros permitidas. O gerador alternativo `Gerar_PDF.bat` mantém seu relatório operacional anterior, sem o novo bloco de SLA.

Esta versão acompanha o prazo dos cases em aberto. Para medir o cumprimento do SLA dos concluídos, é necessário acrescentar uma data de conclusão confiável; DATA DE MODIFICAÇÃO mantém o valor da base de entrada e não substitui uma data de conclusão.

## Atualizar uma tracking existente

Abra `Programa/Atualizar_Painel.bat` e informe o caminho do `.xlsx` ou `.xlsm` já utilizado. Será criada outra cópia em uma pasta `Painel_SLA_<execução>`. A rotina preserva os valores e status da Tracking, o Resumo e os feriados cadastrados. Não redistribui nem reinicia os atendimentos.

Novos lotes gerados por `Programa/Iniciar.bat` já incluem o SLA. Cada novo lote começa com a lista de feriados vazia; cadastre o calendário aplicável antes de utilizar os indicadores.
