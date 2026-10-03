# Teste com a base fictícia

Todos os 42 registros de `Base_Ficticia_Creatio.xlsx` foram criados do zero. A base contém apenas identificadores numéricos inventados, datas de exemplo, equipes e pessoas explicitamente fictícias. Não contém CPF, e-mail, telefone, endereço, nome de cliente ou número de case da operação.

## Teste pelo programa

1. Abra `Programa/Iniciar.bat`.
2. Cole o caminho completo de `exemplos/Base_Ficticia_Creatio.xlsx`.
3. Informe cinco pessoas: Ana, Bruno, Carla, Diego e Elisa.
4. Sem lista de exclusões, o resultado será de 42 cases: 9, 9, 8, 8 e 8 por pessoa.

Para testar também as exclusões, copie **somente os números fictícios** de `exclusoes_demo.txt` para `Programa/Cases_nao_distribuir.txt` antes de executar. Depois do teste, deixe a lista do programa vazia ou configure a lista adequada ao próximo uso.

## Teste pela linha de comando

Na raiz do projeto, execute:

```powershell
python Programa/distribuir_cases.py "exemplos/Base_Ficticia_Creatio.xlsx" --analistas "Ana" "Bruno" "Carla" "Diego" "Elisa" --exclusoes "exemplos/exclusoes_demo.txt" --sem-pausa
```

Resultado esperado com a lista de demonstração:

| Conferência | Resultado |
| --- | --- |
| Recebidos | 42 |
| Excluídos | 2, os cases fictícios 00090041 e 00090042 |
| Número da lista ausente | 00099999 |
| Distribuídos | 40 |
| Quantidade por pessoa | 8 para cada um dos cinco nomes |
| Status inicial | Pendente em todos os registros |
| Ordem | Abertura crescente, alternando os cinco nomes |

Na Tracking, mude um case para Concluído e outro para Escalado. Preencha o destino do escalonamento em OBSERVAÇÃO. O gestor deverá mostrar 1 concluído, 39 ainda abertos e 1 escalado. Esse escalado já faz parte dos 39 abertos.

Escolha Ana no filtro do gestor para ver seus 8 cases. Volte a Toda a equipe e teste o botão de PDF no Excel instalado, se as macros forem permitidas.

As imagens do guia e da apresentação mostram uma simulação posterior: 20 concluídos, 8 em andamento, 5 pendentes, 4 para revisão e 3 escalados. Esses status ilustram o uso durante o dia e não são o estado inicial gerado pelo distribuidor.
