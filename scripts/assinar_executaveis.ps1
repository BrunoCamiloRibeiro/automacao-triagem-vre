# ==============================================================================
# SMAURB / INSPETORIA FISCAL - ASSINATURA DIGITAL AUTHENTICODE (WINDOWS)
# Gera certificado de Code Signing e assina os binarios para evitar bloqueio de Antivirus
# ==============================================================================

[CmdletBinding()]
param (
    [string]$TargetDir = "$PSScriptRoot\fbuild"
)

Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host " SMAURB - ASSINATURA DIGITAL DE BINARIOS WINDOWS (AUTHENTICODE)" -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan

# 1. Verifica ou cria o certificado de Code Signing na maquina
$certSubject = "CN=SMAURB Inspetoria Fiscal - Automacao VRE"
$cert = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue | Where-Object { $_.Subject -eq $certSubject } | Select-Object -First 1

if (-not $cert) {
    Write-Host "[1/3] Gerando Certificado de Assinatura de Codigo local..." -ForegroundColor Yellow
    try {
        $cert = New-SelfSignedCertificate `
            -Type CodeSigningCert `
            -Subject $certSubject `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -FriendlyName "Certificado SMAURB Triagem VRE" `
            -NotAfter (Get-Date).AddYears(5)
        
        # Adiciona nas autoridades confiaveis do usuario para eliminar aviso de editor desconhecido
        $storeRoot = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
        $storeRoot.Open("ReadWrite")
        $storeRoot.Add($cert)
        $storeRoot.Close()

        $storePub = New-Object System.Security.Cryptography.X509Certificates.X509Store("TrustedPublisher", "CurrentUser")
        $storePub.Open("ReadWrite")
        $storePub.Add($cert)
        $storePub.Close()
        
        Write-Host "    [OK] Certificado gerado e adicionado aos Editores Confiaveis." -ForegroundColor Green
    } catch {
        Write-Host "    [AVISO] Nao foi possivel instalar nas raizes confiaveis: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "[1/3] Certificado de Code Signing localizado: $($cert.Thumbprint)" -ForegroundColor Green
}

# 2. Localiza os executaveis
$arquivos = @(
    "$TargetDir\triagem_vre.exe",
    "$TargetDir\config_vre.exe",
    "$PSScriptRoot\triagem_vre.exe",
    "$PSScriptRoot\config_vre.exe"
)

Write-Host "[2/3] Aplicando Assinatura Digital Authenticode nos executaveis..." -ForegroundColor Yellow
$assinados = 0
foreach ($arq in $arquivos) {
    if (Test-Path $arq) {
        try {
            $sig = Set-AuthenticodeSignature -FilePath $arq -Certificate $cert -HashAlgorithm SHA256
            Write-Host "    [OK] Assinado com sucesso: $(Split-Path $arq -Leaf) (Status: $($sig.Status))" -ForegroundColor Green
            $assinados++
        } catch {
            Write-Host "    [ERRO] Falha ao assinar $arq : $_" -ForegroundColor Red
        }
    }
}

Write-Host "[3/3] Processo concluido! $assinados executavel(is) assinado(s) digitalmente." -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan
