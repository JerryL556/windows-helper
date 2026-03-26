# Windows Helper

Simple Windows desktop helper app with:

- Global hotkey: `Alt+P`
- Hidden startup behavior
- Single-instance protection
- Quick access buttons for common Windows tools

## Run manually

```powershell
python .\windows_helper.py
```

## Install startup launcher

```powershell
powershell -ExecutionPolicy Bypass -File .\install_startup.ps1
```

## Remove startup launcher

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall_startup.ps1
```
