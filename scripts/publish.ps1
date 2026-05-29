<#
publish.ps1 - Cria (se preciso) o repo GitHub 'cadastro', faz commit/push e liga o Pages.
Le o token do Gerenciador de Credenciais do Windows (git:https://github.com).

Uso:
  powershell -ExecutionPolicy Bypass -File publish.ps1 -Message "msg"
#>
param(
  [string]$RepoDir = "C:\Users\COMPRASD\cadastro-site",
  [string]$Message = "Atualiza site de cadastro",
  [string]$Owner   = "MFerrettiGit",
  [string]$Repo    = "cadastro"
)
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$sig = @"
using System;
using System.Runtime.InteropServices;
public class CredVaultCAD {
  [DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
  static extern bool CredRead(string target, uint type, uint flags, out IntPtr cred);
  [StructLayout(LayoutKind.Sequential)]
  struct CREDENTIAL { public uint Flags; public uint Type; public IntPtr TargetName; public IntPtr Comment;
    public long LastWritten; public uint CredentialBlobSize; public IntPtr CredentialBlob;
    public uint Persist; public uint AttributeCount; public IntPtr Attributes; public IntPtr TargetAlias; public IntPtr UserName; }
  public static string Get(string target){ IntPtr p; if(!CredRead(target,1,0,out p)) return "FAIL";
    var c=(CREDENTIAL)Marshal.PtrToStructure(p,typeof(CREDENTIAL)); return Marshal.PtrToStringUni(c.CredentialBlob,(int)c.CredentialBlobSize/2); }
}
"@
Add-Type -TypeDefinition $sig -Language CSharp
$tok = [CredVaultCAD]::Get("git:https://github.com")
if($tok -eq "FAIL"){ throw "Credencial do GitHub nao encontrada (git:https://github.com)." }
$hdr = @{ Authorization = "token $tok"; "User-Agent" = "mferretti-cadastro"; Accept = "application/vnd.github+json" }

# 1) Repo existe?
$exists = $true
try { Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo" -Headers $hdr -Method Get | Out-Null }
catch { $exists = $false }
if(-not $exists){
  Write-Host "Criando repo $Owner/$Repo ..."
  $body = @{ name=$Repo; private=$false; description="Cadastro de produtos M. Ferretti - equipe de vendas"; has_issues=$false; has_wiki=$false } | ConvertTo-Json
  Invoke-RestMethod -Uri "https://api.github.com/user/repos" -Headers $hdr -Method Post -Body $body -ContentType "application/json" | Out-Null
  Start-Sleep -Seconds 2
} else { Write-Host "Repo ja existe." }

# 2) git init/commit/push
Set-Location $RepoDir
if(-not (Test-Path ".git")){ git init | Out-Null; git branch -M main }
git add -A
git -c user.name="MFerrettiGit" -c user.email="compras@mferretti.com.br" commit -m $Message
git push "https://$($Owner):$tok@github.com/$Owner/$Repo.git" main
Write-Host ("PUSH_EXIT=" + $LASTEXITCODE)

# 3) Liga o Pages (branch main / root)
try {
  $pbody = @{ source = @{ branch="main"; path="/" } } | ConvertTo-Json
  Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo/pages" -Headers $hdr -Method Post -Body $pbody -ContentType "application/json" | Out-Null
  Write-Host "Pages habilitado."
} catch {
  if("$($_.Exception.Response.StatusCode.value__)" -eq "409"){ Write-Host "Pages ja estava habilitado." }
  else { Write-Host ("Aviso ao habilitar Pages: " + $_.Exception.Message) }
}
Write-Host ("Site: https://" + $Owner.ToLower() + ".github.io/" + $Repo + "/")
