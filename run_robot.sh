#!/bin/bash

# Navega para o diretório onde este script (.sh) está localizado.
# Isso garante que todos os caminhos relativos funcionem corretamente.
cd "$(dirname "$0")"

echo "Ativando ambiente virtual..."
# Ativa o ambiente virtual (note o caminho para 'activate' no Linux).
source venv/bin/activate

echo "Executando o script Python..."
# É uma boa prática usar 'python3' explicitamente no Linux.
python3 main.py

echo "Desativando ambiente virtual..."
# Desativa o ambiente.
deactivate

echo "Processo concluído."