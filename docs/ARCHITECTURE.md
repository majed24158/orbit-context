# Architecture

`orbit_context.py` owns settings, operation dialogs, progress and notifications. It dispatches local image/PDF operations to `file_tools.py`, Word export to `word_export.py`, and smart conversion to `document_conversion.py`.

`registration.py` defines owned per-user Explorer cascades, stages the stable executable and rolls back failed registry updates. The manager is not a persistent background service.

`word_export.py` initializes a balanced native COM apartment, creates an independent hidden Word instance, opens input read-only, disables macros and link updates, and exports a PDF without opening it for the user.

`codex_bridge.py` runs an installed Codex executable in a hidden Windows Job. The process starts suspended, is assigned to the Job and is then resumed. Requests use a schema and document-data instructions, restricted tool features, bounded output and a deadline. Model output is parsed as data; it is not executed as Python or shell commands.

`document_conversion.py` renders PDF/image pages, requests structured content and optionally compares an owned Word render with the source in a bounded correction pass. `editable_docx.py` validates the page data and builds native text boxes/tables with local image crops. Final output is published through collision-safe atomic output after temporary cleanup.

External dependencies: desktop Word for native export/review and authenticated Codex CLI for smart conversion. Neither is embedded in the executable. Local-only operations do not require a model connection.
