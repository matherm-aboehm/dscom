#$thread = [System.Threading.Thread]::CurrentThread
function OnFolderClose {
    # This is run when the folder / VSCode closes
    . dotnet build-server shutdown
}

try {
    While ($True) {
        sleep 10
    }
    #[System.Threading.Thread]::Sleep([System.Threading.Timeout]::Infinite)
}
finally {
    OnFolderClose
}
