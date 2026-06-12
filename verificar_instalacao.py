import importlib
import py_compile
from pathlib import Path

base = Path(__file__).resolve().parent
print('Verificando projeto Telecom Analytics IA...')

for file_name in ['app.py', 'requirements.txt', 'TelecomDB_exemplo.xlsx', '.env.example', '.gitignore']:
    path = base / file_name
    print(f'{file_name}:', 'OK' if path.exists() else 'FALTANDO')

print('\nVerificando sintaxe do app.py...')
py_compile.compile(str(base / 'app.py'), doraise=True)
print('app.py: sintaxe OK')

print('\nVerificando dependências instaladas...')
modules = {
    'streamlit': 'streamlit',
    'pandas': 'pandas',
    'openpyxl': 'openpyxl',
    'plotly': 'plotly',
    'reportlab': 'reportlab',
    'python-dotenv': 'dotenv',
    'google-genai': 'google.genai',
}
for package, module in modules.items():
    try:
        importlib.import_module(module)
        print(f'{package}: OK')
    except Exception as exc:
        print(f'{package}: NÃO INSTALADO ({exc})')

print('\nVerificação concluída. Para executar: python -m streamlit run app.py')
