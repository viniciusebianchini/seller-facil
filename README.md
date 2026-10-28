# Etiquetas Shopee

Aplicativo local para organizar etiquetas de envio da Shopee, identificar produtos e gerar uma lista de separação. Você envia as planilhas e os ZIPs pelo navegador e baixa os resultados prontos para conferência e impressão.

Projeto independente, sem vínculo oficial com a Shopee. Os pedidos são processados no seu computador, sem envio a serviços externos.

## O que faz

- Recebe listas de empacotamento XLSX e ZIPs originais com etiquetas ZPL.
- Combina vários lotes, incluindo Entrega Padrão e Retire Perto de Você.
- Associa produtos às etiquetas pelo rastreio, sem depender dos números dos nomes dos arquivos.
- Agrupa pedidos por SKU e acrescenta nomes, variações e quantidades no espaço disponível.
- Gera páginas de separação quando os produtos não cabem no rodapé.
- Acrescenta um alerta grande quando o pedido soma mais de uma unidade, inclusive de SKUs diferentes.
- Mostra prévias e gera PDF, ZPL, ZIP e lista de separação em Word com data de emissão.

A sequência permanece **envio → alerta de quantidade, se necessário → páginas de separação, se necessárias**. Outro pedido nunca é colocado entre essas páginas. Kits contam conforme a quantidade de venda informada na planilha.

## Começar no Windows

1. Na seção **Releases** deste repositório, baixe `Etiquetas-Shopee-v1.1.0-Windows.zip`.
2. Extraia o ZIP em uma pasta permanente.
3. Instale Python 3.11 ou superior, com o inicializador `py` disponível, se ainda não tiver Python.
4. Abra `Iniciar.bat`. Na primeira execução, ele prepara as bibliotecas necessárias. Essa instalação precisa de internet.
5. Mantenha a janela aberta e use o aplicativo no navegador em `http://127.0.0.1:8765`.

O ZIP contém o aplicativo em Python, não um executável independente. Em computadores com o runtime compatível do Codex instalado, o iniciador aproveita esse Python. Não é necessário ter conta no Codex para usar o aplicativo com seu próprio Python.

## Preparar um lote

1. Exporte a lista de empacotamento XLSX e o ZIP original das etiquetas ZPL na Shopee.
2. Arraste ambos para a tela. É possível selecionar vários pares de arquivos.
3. Clique em **Preparar etiquetas**.
4. Confira as prévias, os totais e os alertas.
5. Baixe o PDF para impressão ou o ZPL para seu programa de impressão térmica. Baixe também a lista em Word para separar os produtos.

O PDF tem páginas de **100 × 150 mm**. Imprima em **tamanho real / 100%**, sem ajustar à página. Confira a paginação do Word antes de imprimir. A aplicação não envia trabalhos diretamente à impressora.

## Formatos e limites

- Layout suportado: ZPL original com imagem Z64 de **816 × 1218 pontos**.
- Planilha com as colunas `tracking_number`, `order_sn` e `product_info` no formato exportado pela Shopee.
- Até 40 arquivos e 80 MB por envio; limite de 180 MB descompactados por arquivo.
- ZIPs podem conter vários arquivos de etiquetas TXT/ZPL. Arquivos de texto auxiliares sem comandos de etiqueta são ignorados.
- Não envie ZPL já processado por este aplicativo.
- Rastreios duplicados, SKUs ausentes, etiquetas sem correspondência e layouts não suportados interrompem o lote. Nenhum resultado parcial é entregue.

Mudanças futuras nos arquivos exportados pela Shopee podem exigir adaptações. Este projeto não processa etiquetas de outros marketplaces nem oferece suporte geral a todos os comandos ZPL.

## Privacidade

O servidor aceita conexões apenas do próprio computador. Resultados e prévias ficam em `dados/`, dentro da pasta do aplicativo; não são apagados automaticamente. As entradas consolidadas usadas durante o processamento são removidas ao concluir.

Para apagar um lote, feche o aplicativo e remova a pasta correspondente dentro de `dados/`. Não publique essa pasta, arquivos de pedidos, etiquetas, endereços ou rastreios em repositórios e relatos de problemas. O `.gitignore` ajuda a excluir esses arquivos do Git, mas não protege anexos enviados manualmente.

**Esta versão foi feita para uso local.** Não exponha o servidor diretamente à internet: ele não implementa contas de usuários nem isolamento de dados entre usuários.

## Execução manual

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe server.py
```

Outra porta, caso a padrão esteja ocupada:

```powershell
.venv\Scripts\python.exe server.py --port 8766
```

Requer Arial ou DejaVu Sans. O iniciador é voltado ao Windows; outras plataformas não foram validadas.

## Problemas comuns

| Situação | Como resolver |
| --- | --- |
| Os números dos nomes são diferentes | Isso é permitido. Confira se o conteúdo do XLSX e do ZIP pertence ao mesmo lote. |
| Faltam etiquetas ou não há correspondência | Exporte novamente os dois arquivos do mesmo conjunto de pedidos. |
| Rastreio duplicado | Remova a planilha ou ZIP repetido e confira se os lotes se sobrepõem. |
| SKU ausente | Corrija o cadastro/dados na origem e exporte novamente. |
| Porta ocupada | Feche a instância anterior ou escolha outra porta. |
| Atualizei, mas vejo comportamento antigo | Encerre o servidor antigo, inicie a nova versão e recarregue o navegador. |

## Contribuir

Abra uma issue com passos para reproduzir, versão do aplicativo e mensagem de erro. Use somente exemplos fictícios ou devidamente anonimizados. Para alterações no processador, confira associação por rastreio, soma de quantidades, ordem das páginas, preservação dos códigos de envio e dimensões do PDF.

## Licença

MIT. Consulte [LICENSE](LICENSE). As bibliotecas utilizadas possuem suas próprias licenças; não estão incluídas no ZIP e são instaladas separadamente.
