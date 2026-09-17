Automação desenvolvida em Python, com o uso de Inteligência Artificial (IA) como apoio no desenvolvimento e aprimoramento da solução, para consultar automaticamente os motivos de cancelamento de propostas na esteira de negociação.

A aplicação permite informar propostas manualmente ou por meio de uma planilha Excel/CSV, acessa o sistema web, consulta os detalhes de cada proposta e identifica o motivo registrado para o cancelamento. Quando o motivo não está disponível diretamente no histórico, a automação também consulta as atividades executadas para identificar o usuário responsável pelo cancelamento.

Ao final do processamento, os resultados são organizados em uma planilha Excel, contendo o número da proposta, o motivo ou usuário identificado e o log completo da consulta. A solução também possui tratamento de erros, salvamento incremental dos resultados e reinicialização automática do navegador em caso de falhas de sessão, trazendo mais segurança e continuidade ao processo.
