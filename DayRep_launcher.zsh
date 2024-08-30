#!/bin/zsh

if ! command -v R &> /dev/null
then
    echo "R не установлен. Начинаю установку..."

    curl -O https://cloud.r-project.org/bin/macosx/base/R-latest.pkg

    sudo installer -pkg R-latest.pkg -target /

    echo "R успешно установлен."
else
    echo "R уже установлен."
fi

# Запуск Rscript
SCRIPT_DIR="$(cd "$(dirname "${(%):-%N}")" && pwd)"
Rscript "$SCRIPT_DIR/DayRep_auto.R"
