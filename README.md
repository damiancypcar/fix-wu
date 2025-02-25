# fix-wu
Fixes Windows Update stuck problem


1. Run PowerShell as Administrator

2. Set the execution policy

    ```powershell
    Set-ExecutionPolicy Bypass -Scope Process -Force
    ```

3. Run script directly from the web

    ```powershell
    irm 'https://raw.githubusercontent.com/damiancypcar/fix-wu/refs/heads/main/fix-wu.ps1' | iex
    ```
