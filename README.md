# Alerta de Estoque em Tempo Real

Sistema de monitoramento de estoque feito em Google Apps Script, integrado ao **Bling** (ERP), que acompanha em tempo real mais de 1.300 peças e prevê quando cada uma vai acabar.

## O problema

A empresa vende kits montados a partir de peças menores (presilhas e componentes). Quando uma dessas peças acaba, centenas de produtos param de vender. A checagem de estoque era manual e reativa: a falta só era percebida depois de já ter afetado as vendas.

## O que o sistema faz

1. **Recebe movimentações de estoque em tempo real** via webhook do Bling (`doPost`). Filtra apenas as peças rastreadas e ignora eventos de estoque virtual, que refletem kits e não movimentos reais da peça.
2. **Grava os eventos brutos num buffer** ao longo do dia, com trava de concorrência (LockService) para eventos que chegam ao mesmo tempo.
3. **Consolida tudo toda noite:** agrupa por peça, separa saídas e entradas, trata balanços de estoque como valor absoluto (e não como movimento), grava uma linha por peça por dia e apaga registros com mais de 90 dias.
4. **Calcula a previsão de ruptura:** média diária de saída em três janelas (7, 30 e 90 dias). Os dias restantes são calculados pela **maior média de consumo**, a estimativa mais conservadora, para não gerar falsa sensação de segurança.
5. **Classifica cada peça:** 🔴 Zerado, 🔴 Crítico (< 15 dias), 🟡 Atenção (< 45 dias), 🟢 Saudável ou ⚪ Sem movimento.
6. **Gera uma visão consolidada** das três categorias numa aba só, pronta para a equipe filtrar.

## Desafios técnicos

- **Sem histórico retroativo na API:** o Bling não fornece o histórico de movimentações. A solução combina estimativas de relatórios exportados manualmente com os dados reais capturados pelo webhook. Cada janela de média passa a usar dados reais assim que eles existem nela.
- **Códigos de produto sem padrão:** mais de 1.300 códigos cadastrados com formatos diferentes (`CP12-EP5`, `CP45E-P9`, `CP69 - EP0`...). Expressões regulares normalizam os códigos para ordenar as peças corretamente.
- **IDs "mortos":** quando um produto é excluído e recriado no Bling, ele ganha um ID novo, e a planilha continuava presa ao antigo sem atualizar o saldo. Criei rotinas de diagnóstico e limpeza que verificam cada ID na API, limpam os inválidos para forçar nova busca e removem duplicatas, sinalizando para decisão manual os casos ambíguos.
- **Limites da plataforma:** consultas de saldo em lotes de 50 peças, pausas entre chamadas e processamento com progresso salvo para respeitar o limite de execução do Apps Script.
- **Descoberta automática de peças novas:** basta adicionar o código na planilha; o sistema busca o ID no Bling na próxima atualização.

## Estrutura dos arquivos

| Arquivo | Função |
|---|---|
| `codigo.gs` | Autenticação OAuth2, webhook, consolidação noturna, cálculo das médias, previsão e rotinas de diagnóstico |
| `MapaPecas.gs` | Mapa das peças rastreadas (versão de exemplo; os dados reais não são publicados) |
| `appsscript.json` | Manifesto: biblioteca OAuth2 e configuração do web app |

## Tecnologias

Google Apps Script (JavaScript), API REST do Bling v3, OAuth2, Webhooks, Google Sheets, expressões regulares.

## Como configurar

1. Em *Configurações do projeto > Propriedades do script*, cadastre `BLING_ESTOQUE_CLIENT_ID` e `BLING_ESTOQUE_CLIENT_SECRET`.
2. Rode `iniciarAutorizacao` e abra o link gerado para autorizar o acesso ao Bling.
3. Implante o projeto como web app e cadastre a URL como webhook no Bling (eventos de estoque).
4. Crie um gatilho diário por tempo para `consolidarBufferDiario`.

## Próximas melhorias

- Validar a assinatura dos webhooks recebidos, para aceitar apenas requisições vindas do Bling.
- Alertas automáticos por e-mail quando uma peça entrar em estado crítico.

---

*Regras de negócio, arquitetura e validação definidas por mim; código desenvolvido com assistência de IA.*
