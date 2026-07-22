# =========================================================
# Script: gerar_etiquetas.ps1
# Objetivo: ler a planilha do dia (nome = data de hoje, dd.mm.yyyy),
#           filtrar as linhas onde a coluna J contem "Kelcey"
#           e gerar um HTML pronto para impressao em formato
#           de etiquetas (grade, varias por pagina).
#
# Requisito (uma unica vez, no PowerShell normal do usuario):
#   Install-Module -Name ImportExcel -Scope CurrentUser -Force
# =========================================================

param(
    # Pasta onde ficam as planilhas diarias (OneDrive sincronizado)
    [string]$PastaPlanilhas = "C:\Users\ENTREGAS01\OneDrive",

    # Data usada tanto para achar a planilha do dia quanto para nomear o HTML final
    [string]$DataArquivo = (Get-Date -Format "dd.MM.yyyy"),

    # Nome do arquivo do dia. Por padrao, monta sozinho com a data de hoje.
    [string]$NomeArquivo = "$DataArquivo.xlsx",

    # Aba a ler. Deixe $null para pegar a primeira aba.
    [string]$NomeAba = $null,

    # Texto a ser buscado na coluna J (filtro)
    [string]$FiltroResponsavel = "Kelcey",

    # Onde salvar o HTML final: Area de Trabalho, com data + "KELCEY"
    [string]$CaminhoHTML = "C:\Users\ENTREGAS01\Desktop\$DataArquivo KELCEY.html"
)

Import-Module ImportExcel -ErrorAction Stop

$CaminhoExcel = Join-Path $PastaPlanilhas $NomeArquivo

if (-not (Test-Path $CaminhoExcel)) {
    throw "Planilha do dia nao encontrada em: $CaminhoExcel"
}

# --- 1. Ler os dados da planilha pelas letras das colunas -------------------
# HeaderName define nomes proprios para as colunas A, B, C, D, E, F, G, H, I, J, K
# (nao importa o titulo que estiver escrito na planilha, a leitura e' por posicao)
$colunas = @(
    'Protocolo',   # A
    'Nome',        # B
    'ColC',        # C (nao usada)
    'Numero',      # D
    'ColE',        # E (nao usada)
    'ColF',        # F (nao usada)
    'Cidade',      # G
    'Bairro',      # H
    'ColI',        # I (nao usada)
    'Responsavel', # J  -> filtro "Kelcey"
    'Endereco'     # K
)

$parametros = @{
    Path       = $CaminhoExcel
    HeaderName = $colunas
    StartRow   = 2   # pula a linha 1, que e' o cabecalho de verdade da planilha
}
if ($NomeAba) { $parametros['WorksheetName'] = $NomeAba }

$dados = Import-Excel @parametros

if (-not $dados -or $dados.Count -eq 0) {
    throw "Nenhum dado encontrado na planilha do dia."
}

# --- 2. Filtrar somente as linhas da pessoa desejada (coluna J) -------------
$dadosFiltrados = $dados | Where-Object { $_.Responsavel -like "*$FiltroResponsavel*" }

if (-not $dadosFiltrados -or $dadosFiltrados.Count -eq 0) {
    Write-Warning "Nenhuma linha encontrada com '$FiltroResponsavel' na coluna J."
}

# --- 3. Montar uma etiqueta em HTML para cada linha filtrada ----------------
$etiquetasHtml = ""
foreach ($linha in $dadosFiltrados) {
    $etiquetasHtml += @"
        <div class="etiqueta">
            <div class="campo"><span class="rotulo">PROTOCOLO:</span> $($linha.Protocolo)</div>
            <div class="campo"><span class="rotulo">NOME:</span> $($linha.Nome)</div>
            <div class="campo"><span class="rotulo">ENDERECO:</span> $($linha.Endereco) <span class="rotulo">N:</span> $($linha.Numero)</div>
            <div class="campo"><span class="rotulo">BAIRRO:</span> $($linha.Bairro)</div>
            <div class="campo"><span class="rotulo">CIDADE:</span> $($linha.Cidade)</div>
        </div>
"@
}

# --- 4. Montar o documento HTML completo -------------------------------------
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
        font-family: Arial, Helvetica, sans-serif;
    }
    /* Flexbox com quebra de linha: cada etiqueta e' protegida contra
       ser cortada ao meio entre duas paginas na impressao. */
    .folha {
        display: flex;
        flex-wrap: wrap;
        gap: 4mm;
    }
    .etiqueta {
        width: calc(50% - 2mm);   /* 2 etiquetas por linha - mude para 33.33% se quiser 3 */
        border: 1px dashed #999999;
        border-radius: 2mm;
        padding: 4mm;
        min-height: 32mm;
        overflow: hidden;

        /* Impede que a etiqueta seja dividida entre duas paginas */
        break-inside: avoid;
        page-break-inside: avoid;
        -webkit-column-break-inside: avoid;

        display: flex;
        flex-direction: column;
        justify-content: center;
        gap: 1.5mm;
    }
    .campo {
        font-size: 11pt;
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
<body>
    <div class="folha">
$etiquetasHtml
    </div>
</body>
</html>
"@

# --- 5. Salvar o arquivo -----------------------------------------------------
Set-Content -Path $CaminhoHTML -Value $htmlCompleto -Encoding UTF8

Write-Output "Etiquetas geradas com sucesso em: $CaminhoHTML"
Write-Output "Total de etiquetas geradas: $($dadosFiltrados.Count)"
