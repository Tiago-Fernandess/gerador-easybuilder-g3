# Gerador EasyBuilder Pro — G3 Automation

## Uso

O pacote inclui o instalador oficial do Python 3.13.17. Nenhuma biblioteca externa é necessária. Edge ou Chrome devem estar presentes no Windows 10/11.

1. Abra esta pasta.
2. Execute apenas `Gerador_EasyBuilder.bat` — se Python/Tkinter não estiver instalado, ele será instalado automaticamente para o usuário e o app será aberto em seguida.
2. Arraste a planilha de variáveis (`.xlsx`) para a área central.
3. Confira os itens identificados (clique nos cartões para filtrar por categoria) e escolha a pasta de saída.
4. Clique em **Gerar arquivos**.

Arquivos gerados: `Definition_Variables.db`, `Transfer_Variables.xlsx`, `0_Executa Receita BOOLs.ebm` e `1_Salva Receita BOOLs.ebm` (as macros só existem quando há BOOLs em `IHM_RET.Bts`).

## Estrutura

- `app/` — janela do app e lógica de geração.
- `ui/` — interface visual.
- `modelos/` — modelos internos validados.
- `instalador/` — instalador do Python 3.13.17.
- `recursos/` — ícone e criação do atalho.
- `exemplos/` — arquivos de referência já gerados.

Execução sem interface: `python app.py --input planilha.xlsx --output pasta_de_saida`
