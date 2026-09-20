<#
  .msg -> .txt 사이드카 변환기 (노트북에서 실행)

  Drive 동기화 폴더 안의 .msg 를 찾아 같은 이름의 .txt 를 옆에 만든다.
  .txt 는 Drive 가 그대로 읽을 수 있으므로, 이후 메일 본문 확인에
  별도 파싱이 필요 없다.

  실행 예
    powershell -ExecutionPolicy Bypass -File msg_to_txt.ps1 `
      -Root "$env:USERPROFILE\Google Drive\내 컴퓨터\내 노트북\01 도서\2026"

  작업 스케줄러에 등록하면 신규 .msg 만 자동으로 변환된다(-Since 사용).
  Outlook 이 설치된 PC 에서만 동작한다.
#>
param(
  [Parameter(Mandatory = $true)][string]$Root,
  [int]$Since = 0,          # 최근 N일 내 수정분만. 0 = 전체
  [switch]$Force            # 이미 .txt 가 있어도 다시 만든다
)

$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")

$files = Get-ChildItem -LiteralPath $Root -Filter *.msg -Recurse -File
if ($Since -gt 0) {
  $cut = (Get-Date).AddDays(-$Since)
  $files = $files | Where-Object { $_.LastWriteTime -ge $cut }
}

$done = 0; $skip = 0; $fail = 0
foreach ($f in $files) {
  $txt = [IO.Path]::ChangeExtension($f.FullName, ".txt")
  if ((Test-Path -LiteralPath $txt) -and -not $Force) { $skip++; continue }
  try {
    $item = $ns.OpenSharedItem($f.FullName)
    $atts = @()
    for ($i = 1; $i -le $item.Attachments.Count; $i++) {
      $atts += $item.Attachments.Item($i).FileName
    }
    $sb = New-Object Text.StringBuilder
    [void]$sb.AppendLine("제목  : " + $item.Subject)
    [void]$sb.AppendLine("발신  : " + $item.SenderName + " <" + $item.SenderEmailAddress + ">")
    [void]$sb.AppendLine("수신  : " + $item.To)
    [void]$sb.AppendLine("참조  : " + $item.CC)
    [void]$sb.AppendLine("일시  : " + $item.SentOn)
    [void]$sb.AppendLine("첨부  : " + ($atts -join ", "))
    [void]$sb.AppendLine("--- 본문 ---")
    [void]$sb.AppendLine($item.Body)
    [IO.File]::WriteAllText($txt, $sb.ToString(), [Text.Encoding]::UTF8)
    [void]$item.Close(1)     # olDiscard
    $done++
  } catch {
    Write-Warning ("변환 실패: " + $f.FullName + " -> " + $_.Exception.Message)
    $fail++
  }
}
Write-Host "변환 $done / 건너뜀 $skip / 실패 $fail"
