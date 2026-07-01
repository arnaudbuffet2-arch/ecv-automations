$dir = "c:\cerveau 2 obsidian vault\ceveau 2 vault\claude safe"
$pendingFile = Join-Path $dir "pending_followups.json"

$data = Get-Content $pendingFile -Raw | ConvertFrom-Json
$cutoff = (Get-Date).AddHours(-24)

# Nouveaux dans les dernières 24h (via added_at)
$recent = @($data | Where-Object {
    $_.added_at -and ([datetime]$_.added_at) -gt $cutoff
})
$recentCount = $recent.Count

# Brouillons actifs en attente d'envoi
$drafted = @($data | Where-Object { $_.draft_id -and -not $_.sent }).Count

# Coachings en attente de brouillon (futurs, pas encore de draft)
$today = (Get-Date).ToString("yyyy-MM-dd")
$pendingDraft = @($data | Where-Object { -not $_.sent -and -not $_.draft_id }).Count

Add-Type -AssemblyName System.Windows.Forms

if ($recentCount -gt 0) {
    $msg = "$recentCount nouveau(x) coaching(s) detecte(s) (24 dernieres heures).`nBrouillons Gmail en attente d'envoi : $drafted`nCoachings futurs sans brouillon : $pendingDraft"
    [System.Windows.Forms.MessageBox]::Show($msg, "ECV - Suivi coaching", "OK", "Exclamation") | Out-Null
} else {
    $msg = "Aucun nouveau coaching (24 dernieres heures).`nBrouillons Gmail en attente d'envoi : $drafted`nCoachings futurs sans brouillon : $pendingDraft"
    [System.Windows.Forms.MessageBox]::Show($msg, "ECV - Suivi coaching", "OK", "Information") | Out-Null
}
