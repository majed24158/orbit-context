# Testing

## Offline smoke suite

Run on Windows with the dependencies installed:

```powershell
python -m compileall -q app tests
python -m unittest discover -s tests -v
```

These checks do not modify Explorer registration, call Word, send data to Codex or incur model usage.

## Extended checks

The `app/verify_*.py` scripts came from the release's development suite. They create disposable fixtures under `app/qa/`, which is ignored by Git.

- `verify_context.py`: local conversion and temporary isolated registry rollback checks; requires a Windows desktop environment.
- `verify_shell_menu.py`: asks Windows Shell to construct metadata for a unique temporary file extension; does not invoke its verbs.
- `verify_bridge.py`: owned-process cancellation and response-boundary tests without inference.
- `verify_word.py`: real installed Word integration with disposable documents.
- `verify_conversion.py`: **live Codex requests** using synthetic text and real Word rendering. Run only when you intend to use the connected account.
- `verify_editable_docx.py`: detailed native document structure and rendering checks. `--renderer` can select the optional document renderer; Word fallback needs `--word-python` and PDF rendering packages installed.

## Compiled runtime

```powershell
$env:QT_QPA_PLATFORM = 'offscreen'
.\dist\OrbitContext.exe --self-test "$PWD\runtime-check.json"
```

`--word-self-test` requires Word; `--conversion-self-test` makes live Codex requests. Those diagnostics are deliberately excluded from CI. Successful synthetic tests demonstrate the execution path, not OCR accuracy for every document.

The release manifest records the hash of the supplied original Windows archive. Organizing the source does not claim a reproducible byte-identical rebuild or a digitally signed executable.
