from __future__ import annotations

import os

from core.paths import resolve_runtime_paths
from core.telemetry import EventLog

from gui.main_window import MainWindow

def main():
    paths = resolve_runtime_paths()
    event_log = EventLog(paths.state_root)
    app = MainWindow(
        project_root=str(paths.resource_root),
        output_root=str(paths.export_root),
        state_root=str(paths.state_root),
        event_log=event_log,
    )
    automation_receipt = os.environ.get("ER_PREDICTOR_AUTOMATION_RECEIPT", "").strip()
    if automation_receipt:
        from core.native_qa import NativePackageQa

        NativePackageQa(app, automation_receipt).start()
    event_log.emit("app.ready", product="ER_Predictor", portable=paths.portable)
    app.mainloop()


if __name__ == "__main__":
    main()