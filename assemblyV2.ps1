Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Configura a codificação do console para UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- 1. SELEÇÃO DA PASTA RAIZ (EX: UNIDADE OU ANO) ---
$dialogue = New-Object System.Windows.Forms.FolderBrowserDialog
$dialogue.Description = "Selecione a pasta raiz (ex: Unidade ou Ano) para buscar as pastas de Relatórios"
$dialogue.ShowNewFolderButton = $false

if ($dialogue.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit }

$dossierRacine = $dialogue.SelectedPath

# Busca todas as pastas variantes de "RELATÓRIOS" (com/sem acento, singular/plural, maiúscula/minúscula)
$pastasRelatorios = Get-ChildItem -LiteralPath $dossierRacine -Directory -Recurse |
    Where-Object { $_.Name -match "^RELAT[OÓ]RIO[S]?$" }

if ($null -eq $pastasRelatorios -or $pastasRelatorios.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "Nenhuma pasta de relatórios foi encontrada em:`n$dossierRacine",
        "Informação",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    exit
}

# --- 2. CRIAÇÃO DA JANELA DE SELEÇÃO (GUI) ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Seleção de Relatórios por Equipamento"
$form.Size = New-Object System.Drawing.Size(650, 480)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Rótulo de instrução
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(15, 15)
$label.Size = New-Object System.Drawing.Size(600, 25)
$label.Text = "Marque os relatórios que deseja processar (pastas válidas vêm pré-selecionadas):"
$form.Controls.Add($label)

# Lista com caixas de seleção
$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
$checkedListBox.Location = New-Object System.Drawing.Point(15, 45)
$checkedListBox.Size = New-Object System.Drawing.Size(605, 310)
$checkedListBox.CheckOnClick = $true
$form.Controls.Add($checkedListBox)

# Tabela Hash para associar o caminho relativo ao caminho completo
$mapaCaminhos = @{}

# Análise prévia de cada pasta encontrada
foreach ($pastaObj in $pastasRelatorios) {
    $caminhoCompleto = $pastaObj.FullName
    
    # Caminho relativo para exibição amigável na interface
    $caminhoRelativo = $caminhoCompleto.Replace($dossierRacine, "").TrimStart("\", "/")

    # Busca dos 3 arquivos obrigatórios dentro da pasta
    $excel = Get-ChildItem -LiteralPath $caminhoCompleto -File | Where-Object { $_.Extension -in ".xls", ".xlsx" } | Select-Object -First 1
    $anexo = Get-ChildItem -LiteralPath $caminhoCompleto -File | Where-Object { $_.Extension -eq ".pdf" -and $_.Name -like "*ANEXO*" } | Select-Object -First 1
    $ppt   = Get-ChildItem -LiteralPath $caminhoCompleto -File | Where-Object { $_.Extension -in ".ppt", ".pptx" } | Select-Object -First 1

    $estValide = ($null -ne $excel -and $null -ne $anexo -and $null -ne $ppt)
    
    # Adiciona na interface e no mapa
    $checkedListBox.Items.Add($caminhoRelativo, $estValide) | Out-Null
    $mapaCaminhos[$caminhoRelativo] = $caminhoCompleto
}

# Botão Marcar Todos
$btnToutCocher = New-Object System.Windows.Forms.Button
$btnToutCocher.Location = New-Object System.Drawing.Point(15, 370)
$btnToutCocher.Size = New-Object System.Drawing.Size(130, 30)
$btnToutCocher.Text = "Marcar Todos"
$btnToutCocher.Add_Click({
    for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
        $checkedListBox.SetItemChecked($i, $true)
    }
})
$form.Controls.Add($btnToutCocher)

# Botão Desmarcar Todos
$btnToutDecocher = New-Object System.Windows.Forms.Button
$btnToutDecocher.Location = New-Object System.Drawing.Point(155, 370)
$btnToutDecocher.Size = New-Object System.Drawing.Size(130, 30)
$btnToutDecocher.Text = "Desmarcar Todos"
$btnToutDecocher.Add_Click({
    for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
        $checkedListBox.SetItemChecked($i, $false)
    }
})
$form.Controls.Add($btnToutDecocher)

# Botão Iniciar Mesclagem
$btnLancer = New-Object System.Windows.Forms.Button
$btnLancer.Location = New-Object System.Drawing.Point(450, 370)
$btnLancer.Size = New-Object System.Drawing.Size(170, 30)
$btnLancer.Text = "Iniciar Mesclagem"
$btnLancer.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $btnLancer
$form.Controls.Add($btnLancer)

# Exibição da janela
if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit }

# Recupera as pastas selecionadas pelo usuário
$dossiersChoisis = @()
foreach ($item in $checkedListBox.CheckedItems) {
    $dossiersChoisis += $mapaCaminhos[$item]
}

if ($dossiersChoisis.Count -eq 0) { exit }

# --- 3. EXECUÇÃO INDEPENDENTE PARA CADA PASTA SELECIONADA ---
$wshell = New-Object -ComObject WScript.Shell

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@ -ErrorAction SilentlyContinue

$compteurSucces = 0

foreach ($dossier in $dossiersChoisis) {

    # Busca os 3 arquivos exclusivamente na pasta atual
    $excel = Get-ChildItem -LiteralPath $dossier -File | Where-Object { $_.Extension -in ".xls", ".xlsx" } | Select-Object -First 1
    $anexo = Get-ChildItem -LiteralPath $dossier -File | Where-Object { $_.Extension -eq ".pdf" -and $_.Name -like "*ANEXO*" } | Select-Object -First 1
    $ppt   = Get-ChildItem -LiteralPath $dossier -File | Where-Object { $_.Extension -in ".ppt", ".pptx" } | Select-Object -First 1

    # Validação de segurança
    if (-not $excel -or -not $anexo -or -not $ppt) {
        continue
    }

    $fichiers = @($excel.FullName; $anexo.FullName; $ppt.FullName)

    # Execução do PDFelement
    $process = Start-Process `
        "C:\Program Files\Wondershare\PDFelement10\PDFelement.exe" `
        -ArgumentList ("/combine " + (($fichiers | ForEach-Object { '"' + $_ + '"' }) -join " ")) `
        -PassThru

    # Foco e Validação (ENTER)
    $timeout = 10
    while ($timeout -gt 0) {
        $pdfelement = Get-Process PDFelement -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
        if ($null -ne $pdfelement) {
            [Win32]::SetForegroundWindow($pdfelement.MainWindowHandle) | Out-Null
            Start-Sleep -Milliseconds 800
            $wshell.AppActivate($pdfelement.Id) | Out-Null
            $wshell.SendKeys("~")
            break
        }
        Start-Sleep -Seconds 1
        $timeout--
    }

    # Aguarda o arquivo _combine.pdf
    $fichierCombine = $null
    $attente = 30
    while ($attente -gt 0) {
        $fichierCombine = Get-ChildItem -LiteralPath $dossier -File | Where-Object { $_.Name -like "*_combine.pdf" } | Select-Object -First 1
        if ($null -ne $fichierCombine) { break }
        Start-Sleep -Seconds 1
        $attente--
    }

    # Fechamento do processo e renomeação para o nome do arquivo Excel
    if ($null -ne $fichierCombine) {
        Stop-Process -Name PDFelement -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3

        $nomExcel = [System.IO.Path]::GetFileNameWithoutExtension($excel.Name)
        $nouveauNom = "$nomExcel.pdf"
        
        Rename-Item -LiteralPath $fichierCombine.FullName -NewName $nouveauNom -Force
        $compteurSucces++
    }
}

# --- 4. RESUMO FINAL ---
[System.Windows.Forms.MessageBox]::Show(
    "Processamento em lote concluído!`n`nNúmero de pastas mescladas com sucesso: $compteurSucces de $($dossiersChoisis.Count)",
    "Sucesso",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)