# Cria (ou recria) a tarefa diaria das 08:00 no Agendador de Tarefas do Windows.
# Uso (PowerShell):  powershell -ExecutionPolicy Bypass -File .\agendar_aderencia.ps1
param([string]$Pythonw = "")

$ErrorActionPreference = "Stop"
$pasta  = "C:\Users\jaildo.junior\Desktop\ANALISE_PCP_RFK"
$script = Join-Path $pasta "atualizar_aderencia.py"
$nome   = "ANALISE_PCP_RFK - Atualizar Aderencia 08h"

if (-not (Test-Path $script)) { Write-Host "Nao achei $script"; exit 1 }

if (-not $Pythonw) {
    $cmd = Get-Command pythonw.exe -ErrorAction SilentlyContinue
    if ($cmd) { $Pythonw = $cmd.Source }
}
if (-not $Pythonw -or -not (Test-Path $Pythonw)) {
    Write-Host "pythonw.exe nao encontrado. Rode:  .\agendar_aderencia.ps1 -Pythonw 'C:\caminho\para\pythonw.exe'"
    exit 1
}

$acao    = New-ScheduledTaskAction -Execute $Pythonw -Argument "`"$script`"" -WorkingDirectory $pasta
$gatilho = New-ScheduledTaskTrigger -Daily -At "08:00"
$config  = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
           -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $nome -Action $acao -Trigger $gatilho -Settings $config `
    -Description "Copia a aba ADERENCIA para base_aderencia.xlsx e publica no GitHub" -Force | Out-Null

Write-Host "Tarefa criada: $nome (todo dia as 08:00)"
Write-Host "Usando: $Pythonw"
Write-Host "Testar agora:  Start-ScheduledTask -TaskName '$nome'"
Write-Host "Ver o resultado:  $pasta\logs\aderencia.log"
