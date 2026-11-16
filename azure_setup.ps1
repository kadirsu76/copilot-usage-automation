<#
.SYNOPSIS
    Azure Automation Managed Identity'ye Microsoft Graph API İzinlerini Atama Scripti.

.DESCRIPTION
    Bu script, Azure Automation Account'ın System-Assigned Managed Identity'sine
    Microsoft Graph API için gerekli şu Application izinlerini (App Roles) atar:
      1. Reports.Read.All (Copilot kullanım raporlarını okumak için)
      2. Mail.Send        (Raporu e-posta olarak göndermek için)

.EXAMPLE
    .\azure_setup.ps1 -AutomationAccountName "aa-copilot-automation" -ResourceGroupName "rg-automation"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Azure Automation Account Adı")]
    [string]$AutomationAccountName,

    [Parameter(Mandatory = $true, HelpMessage = "Automation Account'ın bulunduğu Resource Group")]
    [string]$ResourceGroupName
)

$ErrorActionPreference = "Stop"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Azure Automation Managed Identity -> Graph API İzin Atama " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Modül Kontrolü
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Applications)) {
    Write-Warning "Microsoft.Graph.Applications modülü yüklü değil. Yükleniyor..."
    Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Repository PSGallery -Force
}
if (-not (Get-Module -ListAvailable -Name Az.Automation)) {
    Write-Warning "Az.Automation modülü yüklü değil. Yükleniyor..."
    Install-Module Az.Automation -Scope CurrentUser -Repository PSGallery -Force
}

# 2. Azure ve Graph Bağlantısı
Write-Host "`n[1/4] Azure ve Microsoft Graph'a bağlanılıyor..." -ForegroundColor Yellow
Connect-AzAccount | Out-Null
Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All" | Out-Null
Write-Host "Bağlantı başarılı." -ForegroundColor Green

# 3. Automation Account ve Managed Identity Bilgisini Alma
Write-Host "`n[2/4] Automation Account '$AutomationAccountName' aranıyor..." -ForegroundColor Yellow
$automationAccount = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName

if (-not $automationAccount) {
    throw "Automation Account '$AutomationAccountName' Resource Group '$ResourceGroupName' içinde bulunamadı!"
}

$principalId = $automationAccount.Identity.PrincipalId
if (-not $principalId) {
    throw "Automation Account üzerinde System-Assigned Managed Identity aktif görünmüyor! Lütfen Azure Portal'dan Automation Account -> Identity -> System assigned: ON yapın."
}

Write-Host "Managed Identity Principal ID (Object ID): $principalId" -ForegroundColor Green

# 4. Microsoft Graph Service Principal'ını Bulma
Write-Host "`n[3/4] Microsoft Graph kurumsal uygulaması aranıyor..." -ForegroundColor Yellow
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000002'"
if (-not $graphSp) {
    throw "Tenant içinde Microsoft Graph Enterprise Application bulunamadı!"
}

# 5. İzinleri Tanımlama ve Atama
$rolesToAssign = @("Reports.Read.All", "Mail.Send")

Write-Host "`n[4/4] İzinler atanıyor..." -ForegroundColor Yellow

# Mevcut atanmış izinleri al
$existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $principalId

foreach ($roleName in $rolesToAssign) {
    $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq $roleName -and $_.AllowedMemberTypes -contains "Application" }
    if (-not $appRole) {
        Write-Warning "Rol '$roleName' Graph API üzerinde bulunamadı!"
        continue
    }

    $alreadyAssigned = $existingAssignments | Where-Object { $_.AppRoleId -eq $appRole.Id }
    if ($alreadyAssigned) {
        Write-Host "  -> '$roleName' zaten atanmış durumda." -ForegroundColor Cyan
    } else {
        Write-Host "  -> '$roleName' atanıyor..." -ForegroundColor Yellow
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $principalId `
            -PrincipalId $principalId `
            -ResourceId $graphSp.Id `
            -AppRoleId $appRole.Id | Out-Null
        Write-Host "  -> '$roleName' BAŞARIYLA ATANDI!" -ForegroundColor Green
    }
}

Write-Host "`n==========================================================" -ForegroundColor Green
Write-Host " TÜM İZİNLER TAMAMLANDI! " -ForegroundColor Green
Write-Host " Azure Automation Runbook artık secret olmadan Managed Identity " -ForegroundColor Green
Write-Host " ile Graph API'ye erişebilir ve mail atabilir. " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
