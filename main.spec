# Compatibility entry point: both spec files build the React desktop application.
from pathlib import Path
exec(compile((Path(SPECPATH) / "build.spec").read_text(encoding="utf-8"), "build.spec", "exec"))
