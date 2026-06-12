# Validação da versão final

Esta versão foi revisada para conter:

- `app.py` principal compilando sem erro de sintaxe.
- `requirements.txt` com dependências necessárias.
- `TelecomDB_exemplo.xlsx` como base inicial.
- `.env` e `.env.example` para uso de Gemini API sem expor chave na interface.
- `.gitignore` protegendo `.env`, dados atualizados, cache e secrets.
- Scripts `run_app.bat` e `run_app.ps1` para facilitar execução no Windows.
- Script `verificar_instalacao.py` para validar dependências e arquivos básicos.

## Como validar localmente

```bash
python verificar_instalacao.py
python -m streamlit run app.py
```

## Observação

A interface Streamlit precisa das dependências instaladas. Caso falte alguma, execute:

```bash
python -m pip install -r requirements.txt
```
