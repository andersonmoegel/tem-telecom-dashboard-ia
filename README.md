# TEM Telecom FinOps — Cockpit Executivo

Dashboard em Streamlit para análise financeira de telecom, baseado automaticamente na planilha mais recente da pasta `uploads` ou em arquivo enviado pelo usuário pela tela.

## O que foi melhorado nesta versão

- **Visual executivo:** paleta mais corporativa com azul, ciano, verde, amarelo e vermelho; menos excesso de gradientes e melhor leitura.
- **KPIs com contexto:** cards mostram valor, comparação recente, meta sugerida e alerta de revisão.
- **Tendência executiva:** faixa com mini gráfico dos últimos períodos para leitura rápida da evolução.
- **Insights automáticos:** cards de concentração, maior impacto financeiro e auditoria.
- **Gráficos mais analíticos:** roscas/pizzas foram substituídas por barras horizontais com participação percentual onde faz mais sentido.
- **Cascata FinOps:** novo visual `Contratado x realizado`, mostrando contratado, ajustes, diferenças e realizado.
- **IA mais presente:** painel lateral com contexto atual e botão para abrir o assistente executivo.
- **Base dinâmica:** o sistema agora lê automaticamente o arquivo Excel mais recente em `uploads`, permitindo adicionar novas bases mantendo a estrutura.

## Como executar

```bash
pip install -r requirements.txt
streamlit run app.py
```

Ou use:

```bash
run_app.bat
```

## Como atualizar a base

Coloque uma nova planilha `.xlsx` ou `.xls` na pasta:

```text
uploads/
```

O app vai priorizar automaticamente o arquivo Excel mais recente da pasta. Também é possível atualizar pela própria sidebar usando **Atualizar base**.

## Estrutura esperada da planilha

O app procura principalmente pela aba `tbFinanceiro` e complementa os dados com abas como:

- `tbFornecedor`
- `tbServicos`
- `tbFilial`
- `tbCentroCusto`

A estrutura pode receber novas operadoras, serviços, contratos e lançamentos. Campos ausentes são tratados com valores padrão sempre que possível, mantendo o dashboard carregado.

## IA

A IA usa Gemini quando existir uma chave no `.env`:

```env
GEMINI_API_KEY=sua_chave
```

Também aceita:

```env
GOOGLE_API_KEY=sua_chave
```

Sem chave, o app usa resumo local automático.

## Atualização V58 — Ver mais e filtro por clique

Esta versão corrige o detalhamento dos gráficos e melhora o comportamento de análise:

- Cada **Ver mais** mostra os dados relacionados ao visual selecionado.
- O detalhamento inclui rankings, quebras complementares, contratos/faturas e lançamentos auditáveis.
- Os botões manuais de filtro abaixo dos gráficos foram removidos.
- O painel passa a filtrar por clique em barras, pontos ou fatias dos gráficos, usando `streamlit-plotly-events` quando disponível.

Para instalar as dependências atualizadas:

```bash
pip install -r requirements.txt
```


## V59 — Analytics Fix e detalhamento avançado

Esta versão corrige o gráfico **Top fornecedores**, reforça o filtro global dos visuais e melhora os detalhamentos **Ver mais**. Também adiciona visuais analíticos extras: heatmap fornecedor x mês, exposição por contrato, anomalias de serviços e treemap fornecedor → serviço.


## V63 — Diagnóstico executivo e maturidade analítica

Esta versão adiciona uma camada de análise executiva sobre os dados da pasta `uploads`:

- Diagnóstico automático do período filtrado.
- Alertas inteligentes para concentração, variação, diferenças, anomalias e governança.
- Painel de oportunidades com estimativa inicial de economia.
- Qualidade dos dados para explicar inconsistências e orientar saneamento da base.
- Exportação de relatório executivo em Markdown.
- Detalhamentos com descrição e caminho de investigação: Fornecedor → Contrato → Serviço → Fatura → Lançamento.

A base continua dinâmica: novos arquivos Excel podem ser adicionados em `uploads`, mantendo a mesma estrutura principal de abas e colunas.


## Base demonstrativa pública

A planilha `TelecomDB_exemplo.xlsx` foi ajustada para demonstrar o potencial do dashboard sem expor dados corporativos reais:

- Localidades fictícias/profissionais, como `Unidade de São Paulo`, `Hub Corporativo Curitiba` e `Centro Operacional Campinas`.
- Centros de custo fictícios com códigos e descrições, como `CC-1101 — TI - Infraestrutura e Redes` e `CC-1402 — Financeiro - Contas a Pagar`.
- Lançamentos distribuídos em múltiplos meses para evidenciar tendência, comparação por período, alertas, variações, contestação e oportunidades de economia.
- Contratos, faturas, códigos de cliente, empresas e chaves de acesso totalmente demonstrativos, sem referência corporativa real.

## Segurança dos dados

Este repositório foi higienizado para publicação pública. A planilha `TelecomDB_exemplo.xlsx` contém dados fictícios/demonstrativos. Arquivos reais enviados pelo usuário devem ficar apenas na pasta `uploads/`, que está protegida no `.gitignore`. Não publique `.env`, bases reais, caches ou arquivos gerados em `dados_atualizados/`.

## Aviso sobre dados demonstrativos

Todos os dados, empresas, filiais, centros de custo, contratos, endereços, CNPJs e responsáveis utilizados neste projeto são fictícios e foram criados exclusivamente para demonstração pública das funcionalidades do dashboard. Qualquer semelhança com empresas, pessoas, contratos ou operações reais é mera coincidência.
