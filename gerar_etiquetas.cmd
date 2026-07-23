<# :
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression (Get-Content '%~f0' -Raw -Encoding UTF8)"
exit
#>

# =========================================================
# Script: gerar_etiquetas.ps1
# =========================================================

param(
    [string]$PastaOneDrive = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\OneDrive").UserFolder,
    [string]$PastaPlanilhas = "$PastaOneDrive",
    [string]$DataArquivo = (Get-Date -Format "dd.MM.yyyy"),
    [string]$NomeArquivo = "$DataArquivo.xlsx",
    [string]$NomeAba = $null,
    [string]$FiltroResponsavel = "Kelcey",
    [string]$CaminhoHTML = "$PastaOneDrive\$DataArquivo KELCEY.html"
)

Import-Module ImportExcel -ErrorAction Stop

$CaminhoExcel = Join-Path $PastaPlanilhas $NomeArquivo

if (-not (Test-Path $CaminhoExcel)) {
    throw "Planilha do dia nao encontrada em: $CaminhoExcel"
}

$colunas = @(
    'Protocolo',
    'Nome',
    'Numero',
    'ColD',
    'ColE',
    'Responsavel',
    'ColG',
    'Cidade',
    'Bairro',
    'Endereco',
    'ColK'
)

$parametros = @{
    Path       = $CaminhoExcel
    HeaderName = $colunas
    StartRow   = 2
}
if ($NomeAba) { $parametros['WorksheetName'] = $NomeAba }

$dados = Import-Excel @parametros

if (-not $dados -or $dados.Count -eq 0) {
    throw "Nenhum dado encontrado na planilha do dia."
}

$dadosFiltrados = $dados | Where-Object { $_.Responsavel -like "*$FiltroResponsavel*" }

if (-not $dadosFiltrados -or $dadosFiltrados.Count -eq 0) {
    Write-Warning "Nenhuma linha encontrada com '$FiltroResponsavel' na coluna F."
}

$etiquetasHtml = ""
foreach ($linha in $dadosFiltrados) {
    $etiquetasHtml += @"
        <div class="etiqueta">
            <div class="campo"><span class="rotulo">NOME:</span> $($linha.Nome) <span class="rotulo">PROTOCOLO:</span> $($linha.Protocolo)</div>
            <div class="campo"><span class="rotulo">ENDERECO:</span> $($linha.Endereco)</div>
            <div class="campo"><span class="rotulo">BAIRRO:</span> $($linha.Bairro)</div>
            <div class="campo"><span class="rotulo">CIDADE:</span> $($linha.Cidade)</div>
            <div class="campo"><span class="rotulo">Número:</span> $($linha.Numero)</div>
        </div>
"@
}

$htmlCompleto = @"
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<title>Etiquetas - $NomeArquivo</title>
<style>
    @page {
        size: A4;
        margin: 10mm;
    }
    * { box-sizing: border-box; }
    body {
        margin: 0;
        font-family: Arial Bold, Helvetica, sans-serif;
    }
    .folha {
        display: flex;
        flex-wrap: wrap;
        gap: 8mm;
    }
    .etiqueta {
        width: calc(100% - 2mm);
        border: 1px dashed #999999;
        border-radius: 2mm;
        padding: 4mm;
        min-height: 32mm;
        overflow: hidden;

        break-inside: avoid;
        page-break-inside: avoid;
        -webkit-column-break-inside: avoid;

        display: flex;
        flex-direction: column;
        justify-content: center;
        gap: 1.5mm;
    }
    .campo {
        font-size: 14pt;
        color: #222222;
    }
    .rotulo {
        font-weight: bold;
        font-size: 14pt;
    }
    @media print {
        .etiqueta { border: 1px dashed #cccccc; }
    }
</style>
</head>
<body contenteditable="true">
    <div class="folha">
$etiquetasHtml
    </div>
</body>
</html>
"@

Set-Content -Path $CaminhoHTML -Value $htmlCompleto -Encoding UTF8

Write-Output "Etiquetas geradas com sucesso em: $CaminhoHTML"
Write-Output "Total de etiquetas geradas: $($dadosFiltrados.Count)"

Start-Process $CaminhoHTML
