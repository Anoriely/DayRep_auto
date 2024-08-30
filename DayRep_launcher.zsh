#!/bin/zsh

if ! command -v R &> /dev/null
then
    echo "R не установлен."

    if ! command -v brew &> /dev/null
    then
        echo "Homebrew не установлен. Начинаю установку..."
        
        /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

        echo "Homebrew успешно установлен."
    else
        echo "Homebrew уже установлен."
    fi

    echo "Начинаю установку R через Homebrew..."
    export PATH=/opt/homebrew/bin:$PATH
    brew install --cask R

    echo "R успешно установлен."
else
    echo "R уже установлен."
fi

SCRIPT_DIR="$(cd "$(dirname "${(%):-%N}")" && pwd)"
Rscript "$SCRIPT_DIR/DayRep_auto.R"
