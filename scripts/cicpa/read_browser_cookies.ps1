# 关联方识别与核查 — Edge/Chrome Cookie 读取
# 改编自 jackwener/OpenCLI(Apache-2.0)
# 本文件的修改部分以 Apache-2.0 协议发布,与上游保持一致。
#param(
    [Parameter(Mandatory = $true)]
    [string]$WebSocketUrl,
    [ValidateSet("cdp", "bidi")]
    [string]$Protocol = "cdp"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$socket = New-Object System.Net.WebSockets.ClientWebSocket
$cancellation = [Threading.CancellationToken]::None

function Send-BrowserCommand {
    param(
        [int]$Id,
        [string]$Method,
        [hashtable]$Params
    )
    $requestText = @{
        id = $Id
        method = $Method
        params = $Params
    } | ConvertTo-Json -Compress -Depth 8
    $requestBytes = [Text.Encoding]::UTF8.GetBytes($requestText)
    $request = [ArraySegment[byte]]::new($requestBytes)
    $socket.SendAsync(
        $request,
        [Net.WebSockets.WebSocketMessageType]::Text,
        $true,
        $cancellation
    ).GetAwaiter().GetResult()

    do {
        $memory = New-Object IO.MemoryStream
        try {
            do {
                $buffer = New-Object byte[] 65536
                $segment = [ArraySegment[byte]]::new($buffer)
                $received = $socket.ReceiveAsync(
                    $segment,
                    $cancellation
                ).GetAwaiter().GetResult()
                if ($received.MessageType -eq [Net.WebSockets.WebSocketMessageType]::Close) {
                    throw "Browser closed the connection"
                }
                $memory.Write($buffer, 0, $received.Count)
            } while (-not $received.EndOfMessage)
            $text = [Text.Encoding]::UTF8.GetString($memory.ToArray())
            $response = $text | ConvertFrom-Json
        }
        finally {
            $memory.Dispose()
        }
    } while ($response.id -ne $Id)

    if ($response.error) {
        throw ($response.error | ConvertTo-Json -Compress)
    }
    return $response
}

try {
    $socket.ConnectAsync([Uri]$WebSocketUrl, $cancellation).GetAwaiter().GetResult()
    if ($Protocol -eq "bidi") {
        Send-BrowserCommand -Id 1 -Method "session.new" -Params @{
            capabilities = @{}
        } | Out-Null
        $response = Send-BrowserCommand -Id 2 -Method "storage.getCookies" -Params @{}
    }
    else {
        $response = Send-BrowserCommand -Id 1 -Method "Storage.getCookies" -Params @{}
    }
    $cookies = @(
        $response.result.cookies | ForEach-Object {
            $cookieValue = if ($_.value -is [string]) {
                $_.value
            }
            else {
                $_.value.value
            }
            [PSCustomObject]@{
                name = [string]$_.name
                value = [string]$cookieValue
                domain = [string]$_.domain
            }
        }
    )
    ConvertTo-Json -InputObject $cookies -Compress
}
finally {
    if ($socket.State -eq [Net.WebSockets.WebSocketState]::Open) {
        $socket.CloseAsync(
            [Net.WebSockets.WebSocketCloseStatus]::NormalClosure,
            "",
            $cancellation
        ).GetAwaiter().GetResult()
    }
    $socket.Dispose()
}
