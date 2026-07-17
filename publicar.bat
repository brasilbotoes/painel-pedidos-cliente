@echo off
REM ============================================================
REM Publicar Dashboard - Brasil Botoes
REM Roda o script Python (atualiza dados) e envia para o GitHub.
REM Configure este arquivo no Agendador de Tarefas do Windows
REM para rodar sozinho, por exemplo a cada 2 horas.
REM ============================================================

REM Ajuste este caminho para a pasta do seu repositorio git
cd /d "C:\Users\SEU_USUARIO\Documents\dashboard-brasil-botoes"

echo [%date% %time%] Atualizando dados...
python atualizar_dashboard.py

if errorlevel 1 (
    echo Falha ao atualizar os dados. Publicacao cancelada.
    exit /b 1
)

echo [%date% %time%] Publicando no GitHub...
git add data.json
git commit -m "Atualizacao automatica %date% %time%"
git push origin main

echo [%date% %time%] Concluido.
