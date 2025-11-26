# Script para aplicar permissões eDiscovery sem confirmação interativa
param(
    [Parameter(Mandatory=$true)]
    [string]$UserEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$RoleGroup = "eDiscoveryManager"
)

Write-Host "`n╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║       Aplicando Permissões eDiscovery Manager                    ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

Write-Host "`nUsuário: $UserEmail" -ForegroundColor Yellow
Write-Host "Grupo: $RoleGroup" -ForegroundColor Yellow

# Verifica e importa módulo
Write-Host "`nVerificando módulo ExchangeOnlineManagement..." -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Instalando módulo..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}
Import-Module ExchangeOnlineManagement
Write-Host "✓ Módulo carregado" -ForegroundColor Green

# Função para trazer janelas para frente
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class WindowHelper {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool BringWindowToTop(IntPtr hWnd);
        
        public const int SW_RESTORE = 9;
    }
"@

# Conecta ao Compliance Center
Write-Host "`nConectando ao Microsoft Purview (Security & Compliance)..." -ForegroundColor Cyan
Write-Host "Uma janela de autenticação será aberta EM PRIMEIRO PLANO..." -ForegroundColor Yellow
Write-Host "Por favor, faça login com suas credenciais de administrador" -ForegroundColor Gray

# Inicia monitoramento de janelas de autenticação em background
$job = Start-Job -ArgumentList @("Sign in", "Entrar", "Microsoft", "Authentication", "Autenticação") -ScriptBlock {
    param($titles)
    
    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        using System.Text;
        public class WindowMonitor {
            [DllImport("user32.dll")]
            public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
            
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
            
            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            
            [DllImport("user32.dll")]
            public static extern bool BringWindowToTop(IntPtr hWnd);
            
            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);
            
            [DllImport("user32.dll")]
            public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
            
            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
            
            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
            
            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern int GetWindowTextLength(IntPtr hWnd);
            
            public const int SW_RESTORE = 9;
            public const int SW_SHOW = 5;
        }
"@
    
    for ($i = 0; $i -lt 60; $i++) {
        [WindowMonitor]::EnumWindows({
            param($hWnd, $lParam)
            $length = [WindowMonitor]::GetWindowTextLength($hWnd)
            if ($length -gt 0) {
                $sb = New-Object System.Text.StringBuilder ($length + 1)
                [WindowMonitor]::GetWindowText($hWnd, $sb, $sb.Capacity) | Out-Null
                $windowTitle = $sb.ToString()
                
                foreach ($searchTitle in $titles) {
                    if ($windowTitle -like "*$searchTitle*") {
                        [WindowMonitor]::ShowWindow($hWnd, [WindowMonitor]::SW_RESTORE) | Out-Null
                        [WindowMonitor]::BringWindowToTop($hWnd) | Out-Null
                        [WindowMonitor]::SetForegroundWindow($hWnd) | Out-Null
                        Start-Sleep -Milliseconds 100
                    }
                }
            }
            return $true
        }, [IntPtr]::Zero) | Out-Null
        
        Start-Sleep -Milliseconds 500
    }
}

try {
    Connect-IPPSSession -ErrorAction Stop
    Write-Host "✓ Conectado com sucesso!" -ForegroundColor Green
}
catch {
    Write-Host "❌ Erro ao conectar: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    # Para o monitoramento de janelas
    if ($job) {
        Stop-Job -Job $job -ErrorAction SilentlyContinue
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
}

# Verifica usuário
Write-Host "`nVerificando usuário..." -ForegroundColor Cyan
try {
    $user = Get-User -Identity $UserEmail -ErrorAction Stop
    Write-Host "✓ Usuário encontrado: $($user.DisplayName)" -ForegroundColor Green
}
catch {
    Write-Host "❌ Usuário não encontrado: $UserEmail" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

# Verifica se já tem permissão
Write-Host "`nVerificando permissões atuais..." -ForegroundColor Cyan
$roleGroupMembers = Get-RoleGroupMember -Identity $RoleGroup -ErrorAction SilentlyContinue

if ($roleGroupMembers.PrimarySmtpAddress -contains $UserEmail) {
    Write-Host "✓ Usuário já possui estas permissões!" -ForegroundColor Yellow
    Write-Host "   Já é membro do grupo '$RoleGroup'" -ForegroundColor Gray
}
else {
    # Adiciona ao grupo
    Write-Host "`nAdicionando permissões..." -ForegroundColor Cyan
    try {
        Add-RoleGroupMember -Identity $RoleGroup -Member $UserEmail -ErrorAction Stop
        Write-Host "✓ Permissões atribuídas com sucesso!" -ForegroundColor Green
    }
    catch {
        # Se o erro for que o usuário já é membro, não é um erro crítico
        if ($_.Exception.Message -like "*already a member*" -or $_.Exception.Message -like "*já é membro*") {
            Write-Host "✓ Usuário já possui estas permissões!" -ForegroundColor Yellow
            Write-Host "   Já é membro do grupo '$RoleGroup'" -ForegroundColor Gray
        }
        else {
            Write-Host "❌ Erro ao atribuir permissões: $($_.Exception.Message)" -ForegroundColor Red
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            exit 1
        }
    }
}

# Exibe resumo
Write-Host "`n╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║              PERMISSÕES CONFIGURADAS COM SUCESSO                  ║" -ForegroundColor Green
Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Green

Write-Host "`nUsuário: $($user.DisplayName) ($UserEmail)" -ForegroundColor Cyan
Write-Host "Grupo: $RoleGroup" -ForegroundColor Cyan

Write-Host "`nPróximos Passos:" -ForegroundColor Yellow
Write-Host "  1. Aguardar ~15 minutos para propagação de permissões" -ForegroundColor Gray
Write-Host "  2. Fazer logout e login novamente no Microsoft 365" -ForegroundColor Gray
Write-Host "  3. Acessar: https://purview.microsoft.com/contentsearch" -ForegroundColor Gray

# Desconecta
Write-Host "`nDesconectando..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "✓ Concluído!" -ForegroundColor Green
Write-Host ""
