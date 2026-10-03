# Teste com dados fictícios

A base `Base_Ficticia_Cases.xlsx` deve conter somente registros criados para demonstração.

Nenhum dado utilizado no teste deve representar clientes, pessoas ou operações reais.

## Teste pelo programa

1. Abra `Programa/Iniciar.bat`.
2. Informe o caminho de `exemplos/Base_Ficticia_Cases.xlsx`.
3. Informe cinco responsáveis:
   - Ana
   - Bruno
   - Carla
   - Diego
   - Elisa

Sem cases em espera, os 42 registros devem ser distribuídos da seguinte forma:

- Ana: 9
- Bruno: 9
- Carla: 8
- Diego: 8
- Elisa: 8

## Teste com cases em espera

Utilize a lista:

`exemplos/exclusoes_demo.txt`

pela linha de comando:

```powershell
python Programa/distribuir_cases.py "exemplos/Base_Ficticia_Cases.xlsx" --analistas "Ana" "Bruno" "Carla" "Diego" "Elisa" --em-espera "exemplos/exclusoes_demo.txt" --sem-pausa