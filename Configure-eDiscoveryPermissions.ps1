<#
.SYNOPSIS
    Configura permissões de eDiscovery Manager para usuários executarem Content Search
    
.DESCRIPTION
    Este script atribui as permissões necessárias para que um usuário possa:
    - Criar e gerenciar Content Searches no Microsoft Purview
    - Exportar resultados de pesquisas de eDiscovery
    - Acessar o portal de Compliance/Purview
    
    Requer permissões de Administrador Global ou Compliance Administrator para executar.
    
.PARAMETER UserEmail
    Email do usuário que receberá as permissões de eDiscovery Manager
    
.PARAMETER RoleGroup
    Grupo de função a ser atribuído. Opções:
    - eDiscoveryManager: Permite criar e gerenciar suas próprias pesquisas (padrão)
    - eDiscoveryAdministrator: Permite gerenciar todas as pesquisas da organização
    
.EXAMPLE
    .\Configure-eDiscoveryPermissions.ps1
    Executa o script em modo interativo, solicitando o email do usuário
    
.EXAMPLE
    .\Configure-eDiscoveryPermissions.ps1 -UserEmail "admin@contoso.com"
    Atribui permissões de eDiscovery Manager ao usuário especificado
    
.EXAMPLE
    .\Configure-eDiscoveryPermissions.ps1 -UserEmail "admin@contoso.com" -RoleGroup "eDiscoveryAdministrator"
    Atribui permissões de eDiscovery Administrator ao usuário especificado
    
.NOTES
    Autor: Script de Configuração de Permissões eDiscovery
    Versão: 1.0
    Data: 2025-11-22
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, HelpMessage="Email do usuário que receberá as permissões")]
    [ValidatePattern('^[\w\.-]+@[\w\.-]+\.\w+$', ErrorMessage="Formato de email inválido")]
    [string]$UserEmail,
    
    [Parameter(Mandatory=$false, HelpMessage="Grupo de função: eDiscoveryManager ou eDiscoveryAdministrator")]
    [ValidateSet("eDiscoveryManager", "eDiscoveryAdministrator")]
    [string]$RoleGroup = "eDiscoveryManager"
)

Write-Host "`n╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║       Configuração de Permissões eDiscovery Manager              ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

# Função para solicitar email do usuário
function Get-UserEmailInput {
    Write-Host "`n📧 INFORMAR USUÁRIO" -ForegroundColor Yellow
    Write-Host "Por favor, informe o email do usuário que receberá as permissões de eDiscovery." -ForegroundColor Gray
    Write-Host ""
    
    do {
        $email = Read-Host "Email do usuário"
        
        # Valida formato do email
        if ($email -match '^[\w\.-]+@[\w\.-]+\.\w+$') {
            Write-Host "✓ Email válido: $email" -ForegroundColor Green
            return $email
        }
        else {
            Write-Host "❌ Formato de email inválido. Por favor, tente novamente." -ForegroundColor Red
            Write-Host "   Exemplo: usuario@contoso.com" -ForegroundColor Gray
            Write-Host ""
        }
    } while ($true)
}

# Função para solicitar confirmação
function Get-UserConfirmation {
    param(
        [string]$Email,
        [string]$Role
    )
    
    Write-Host "`n⚠️  CONFIRMAÇÃO" -ForegroundColor Yellow
    Write-Host "╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║  As seguintes permissões serão atribuídas:                       ║" -ForegroundColor Yellow
    Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   👤 Usuário: $Email" -ForegroundColor Cyan
    Write-Host "   🔐 Grupo de Permissão: $Role" -ForegroundColor Cyan
    Write-Host ""
    
    if ($Role -eq "eDiscoveryManager") {
        Write-Host "   Permissões concedidas:" -ForegroundColor Gray
        Write-Host "   • Criar e gerenciar suas próprias pesquisas de conteúdo" -ForegroundColor Gray
        Write-Host "   • Exportar resultados de pesquisas criadas pelo usuário" -ForegroundColor Gray
        Write-Host "   • Acessar o portal do Microsoft Purview" -ForegroundColor Gray
    }
    else {
        Write-Host "   Permissões concedidas:" -ForegroundColor Gray
        Write-Host "   • Gerenciar TODAS as pesquisas de conteúdo da organização" -ForegroundColor Gray
        Write-Host "   • Exportar qualquer resultado de pesquisa" -ForegroundColor Gray
        Write-Host "   • Acessar e administrar eDiscovery cases" -ForegroundColor Gray
    }
    
    Write-Host ""
    $response = Read-Host "Deseja continuar? (S/N)"
    
    return ($response -eq 'S' -or $response -eq 's' -or $response -eq 'Sim' -or $response -eq 'sim')
}

# Função para instalar módulo ExchangeOnlineManagement se necessário
function Install-ExchangeOnlineModule {
    Write-Host "`nVerificando módulo ExchangeOnlineManagement..." -ForegroundColor Cyan
    
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "Módulo não encontrado. Instalando..." -ForegroundColor Yellow
        try {
            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            Write-Host "✓ Módulo instalado com sucesso" -ForegroundColor Green
        }
        catch {
            Write-Error "Erro ao instalar módulo: $_"
            return $false
        }
    }
    else {
        Write-Host "✓ Módulo já instalado" -ForegroundColor Green
    }
    
    Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
    return $true
}

# Função para conectar ao Security & Compliance Center
function Connect-ToComplianceCenter {
    Write-Host "`nConectando ao Microsoft Purview (Security & Compliance)..." -ForegroundColor Cyan
    Write-Host "Uma janela de autenticação será aberta..." -ForegroundColor Yellow
    
    try {
        Connect-IPPSSession -ErrorAction Stop
        Write-Host "✓ Conectado com sucesso ao Compliance Center" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Erro ao conectar: $_"
        Write-Host "`nCertifique-se de que você tem permissões de Administrador Global ou Compliance Administrator" -ForegroundColor Yellow
        return $false
    }
}

# Função para verificar se usuário existe
function Test-UserExists {
    param([string]$Email)
    
    Write-Host "`n🔍 Verificando usuário $Email..." -ForegroundColor Cyan
    
    try {
        $user = Get-User -Identity $Email -ErrorAction Stop
        Write-Host "✓ Usuário encontrado!" -ForegroundColor Green
        Write-Host "   Nome: $($user.DisplayName)" -ForegroundColor Gray
        Write-Host "   UPN: $($user.UserPrincipalName)" -ForegroundColor Gray
        return $true
    }
    catch {
        Write-Host "❌ Usuário não encontrado no tenant: $Email" -ForegroundColor Red
        Write-Host "   Verifique se o email está correto e se o usuário existe no Microsoft 365" -ForegroundColor Yellow
        return $false
    }
}

# Função para adicionar usuário ao grupo de função
function Add-UserToRoleGroup {
    param(
        [string]$Email,
        [string]$RoleGroupName
    )
    
    Write-Host "`n🔐 Atribuindo permissões..." -ForegroundColor Cyan
    Write-Host "   Adicionando ao grupo: $RoleGroupName" -ForegroundColor Gray
    
    try {
        # Verifica se o usuário já está no grupo
        $roleGroupMembers = Get-RoleGroupMember -Identity $RoleGroupName -ErrorAction SilentlyContinue
        
        if ($roleGroupMembers.PrimarySmtpAddress -contains $Email) {
            Write-Host "⚠️  Usuário já possui estas permissões!" -ForegroundColor Yellow
            Write-Host "   O usuário já é membro do grupo '$RoleGroupName'" -ForegroundColor Gray
            return $true
        }
        
        # Adiciona o usuário ao grupo
        Add-RoleGroupMember -Identity $RoleGroupName -Member $Email -ErrorAction Stop
        Write-Host "✓ Permissões atribuídas com sucesso!" -ForegroundColor Green
        Write-Host "   Usuário adicionado ao grupo '$RoleGroupName'" -ForegroundColor Gray
        return $true
    }
    catch {
        Write-Host "❌ Erro ao atribuir permissões" -ForegroundColor Red
        Write-Error $_.Exception.Message
        return $false
    }
}

# Função para exibir permissões atuais do usuário
function Show-UserPermissions {
    param([string]$Email)
    
    Write-Host "`n╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                 PERMISSÕES CONFIGURADAS                           ║" -ForegroundColor Green
    Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    
    try {
        $user = Get-User -Identity $Email
        Write-Host "`nUsuário: $($user.DisplayName) ($Email)" -ForegroundColor Cyan
        Write-Host "`nGrupos de Função:" -ForegroundColor Yellow
        
        # Lista grupos de função relacionados a eDiscovery
        $eDiscoveryGroups = @(
            "eDiscovery Manager",
            "eDiscovery Administrator",
            "Compliance Administrator",
            "Organization Management"
        )
        
        $userGroups = @()
        foreach ($group in $eDiscoveryGroups) {
            try {
                $members = Get-RoleGroupMember -Identity $group -ErrorAction SilentlyContinue
                if ($members.PrimarySmtpAddress -contains $Email) {
                    $userGroups += $group
                    Write-Host "  ✓ $group" -ForegroundColor Green
                }
            }
            catch {
                # Grupo pode não existir
            }
        }
        
        if ($userGroups.Count -eq 0) {
            Write-Host "  ⚠️  Nenhum grupo de função eDiscovery atribuído" -ForegroundColor Yellow
        }
        
        Write-Host "`nPróximos Passos:" -ForegroundColor Cyan
        Write-Host "  1. Usuário deve aguardar ~15 minutos para propagação de permissões" -ForegroundColor Gray
        Write-Host "  2. Fazer logout e login novamente no Microsoft 365" -ForegroundColor Gray
        Write-Host "  3. Acessar: https://purview.microsoft.com/contentsearch" -ForegroundColor Gray
        Write-Host "  4. Executar: .\Export-ArchiveMailbox-EXO.ps1 -Mailbox <email> -OlderThanDays 730" -ForegroundColor Gray
        Write-Host ""
    }
    catch {
        Write-Error "Erro ao exibir permissões: $_"
    }
}

# ============================================================================
# SCRIPT PRINCIPAL
# ============================================================================

# Se o email não foi fornecido como parâmetro, solicita interativamente
if ([string]::IsNullOrWhiteSpace($UserEmail)) {
    $UserEmail = Get-UserEmailInput
}

Write-Host "`n📋 CONFIGURAÇÃO" -ForegroundColor Yellow
Write-Host "   Usuário: $UserEmail" -ForegroundColor Gray
Write-Host "   Grupo: $RoleGroup" -ForegroundColor Gray
Write-Host ""

# Solicita confirmação antes de prosseguir
if (-not (Get-UserConfirmation -Email $UserEmail -Role $RoleGroup)) {
    Write-Host "`n⚠️  Operação cancelada pelo usuário" -ForegroundColor Yellow
    exit 0
}

# 1. Instala módulo se necessário
if (-not (Install-ExchangeOnlineModule)) {
    Write-Host "`n❌ Não foi possível instalar o módulo necessário" -ForegroundColor Red
    exit 1
}

# 2. Conecta ao Compliance Center
if (-not (Connect-ToComplianceCenter)) {
    Write-Host "`n❌ Não foi possível conectar ao Compliance Center" -ForegroundColor Red
    exit 1
}

# 3. Verifica se usuário existe
if (-not (Test-UserExists -Email $UserEmail)) {
    Write-Host "`n❌ Usuário não encontrado no tenant" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

# 4. Adiciona usuário ao grupo de função
if (-not (Add-UserToRoleGroup -Email $UserEmail -RoleGroupName $RoleGroup)) {
    Write-Host "`n❌ Não foi possível adicionar usuário ao grupo" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

# 5. Exibe resumo das permissões
Show-UserPermissions -Email $UserEmail

# 6. Desconecta
Write-Host "`nDesconectando..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "✓ Desconectado" -ForegroundColor Green

Write-Host "`n✅ CONFIGURAÇÃO CONCLUÍDA COM SUCESSO!" -ForegroundColor Green
Write-Host "   O usuário $UserEmail agora pode executar Content Searches" -ForegroundColor Gray
Write-Host ""
