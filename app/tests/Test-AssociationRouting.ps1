$ErrorActionPreference='Stop'
$root=Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
. (Join-Path $root 'WordDefault.ps1') -LibraryOnly
. (Join-Path $root 'Batch.ps1')
# Load and verify the real bundled component, but force a stop before any write.
$script:batchSnapshot=''
$caught=$false
try { Write-WordLatestAssociation '.docx' 'Word.Document.12' } catch {
 if ($_.Exception.Message -match '缺少设置前记录|新关联算法未通过|本机的新关联校验暂不支持') { $caught=$true }
 else { throw }
}
if (!$caught) { throw 'Missing snapshot must prevent writes.' }
Write-Output 'PASS: bundled new engine loads and refuses a write without a snapshot'

# Mutation backends are fakes from here onward.
$script:isWin7=$false
$script:testScheme=0
$script:legacyCalls=0; $script:latestCalls=0
function Get-WordHashScheme { return $script:testScheme }
function Set-FTA { param($ProgId,$Extension) $script:legacyCalls++ }
function Write-WordLatestAssociation { param($Extension,$ProgId) $script:latestCalls++ }
Write-WordAssociation '.docx' 'Word.Document.12'
if ($legacyCalls -ne 1 -or $latestCalls -ne 0) { throw 'Legacy routing failed' }
$script:testScheme=1
Write-WordAssociation '.docx' 'Word.Document.12'
if ($legacyCalls -ne 1 -or $latestCalls -ne 1) { throw 'New routing failed' }
$script:testScheme=2; $caught=$false
try { Write-WordAssociation '.docx' 'Word.Document.12' } catch { $caught=$true }
if (!$caught -or $legacyCalls -ne 1 -or $latestCalls -ne 1) { throw 'Unknown scheme must stop before mutation' }
$caught=$false
try { Write-WordAssociation '.pdf' 'Word.Document.12' } catch { $caught=$true }
if (!$caught -or $legacyCalls -ne 1 -or $latestCalls -ne 1) { throw 'Out-of-scope type must stop' }
Write-Output 'PASS: legacy/new routing, unknown-scheme rejection and extension allowlist'
