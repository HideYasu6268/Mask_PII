# main_all_in.py(APIキー・署名の実値入り、git管理外)をonefile exeにビルドする。
# main_all_in.pyが無い場合は main_all_in_no_api.py をコピーし、
# _EMBEDDED_SIGNATURE / _EMBEDDED_API_KEYS を実際の値に書き換えてから実行すること。
#
# 実行方法: venvを有効化した状態で `powershell -ExecutionPolicy Bypass -File build_exe.ps1`

venv\Scripts\pyinstaller.exe --noconfirm --clean --windowed --onefile --name Mask_PII `
  --collect-all customtkinter `
  --collect-all tksheet `
  --collect-all spacy `
  --collect-all ginza `
  --collect-all ja_ginza `
  --collect-all sudachipy `
  --collect-all sudachidict_core `
  --collect-all llama_cpp `
  --hidden-import win32timezone `
  --hidden-import win32com.gen_py `
  main_all_in.py
