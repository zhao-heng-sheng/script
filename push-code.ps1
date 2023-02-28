Function push-code{
    Set-Location E:\project\notes
    $now = Get-Date
    $msg = "==>" + $now.ToString('yyyy年MM月dd日 HH:mm:ss')
    Write-Output $msg >> .\gitpush.log
    git pull >> .\gitpush.log
    git add . >> .\gitpush.log
    git commit -m $msg >> .\gitpush.log
    git push >> .\gitpush.log
}
push-code