"""Actual LibreOffice edit/resave of SDK-generated PPTX; acceptance tooling only."""
import argparse
import json
import os
from pathlib import Path
import signal
import socket
import subprocess
import tempfile
import time
import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    result = PropertyValue()
    result.Name = name
    result.Value = value
    return result


def walk(page):
    for index in range(page.getCount()):
        shape = page.getByIndex(index)
        yield shape
        if shape.supportsService("com.sun.star.drawing.GroupShape"):
            yield from walk(shape)


def edit(source, target):
    with tempfile.TemporaryDirectory(prefix="dagent-ppt-edit-") as profile:
        with socket.socket() as port:
            port.bind(("127.0.0.1", 0))
            number = port.getsockname()[1]
        process = subprocess.Popen(["libreoffice", "-env:UserInstallation=" + Path(profile).as_uri(),
                                    "--headless", "--norestore", "--nodefault", "--nofirststartwizard",
                                    f"--accept=socket,host=127.0.0.1,port={number};urp;StarOffice.ComponentContext"],
                                   stdout=subprocess.PIPE, stderr=subprocess.PIPE, start_new_session=True)
        document = None
        try:
            resolver = uno.getComponentContext().ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", uno.getComponentContext())
            deadline = time.monotonic() + 25
            while True:
                try:
                    context = resolver.resolve(f"uno:socket,host=127.0.0.1,port={number};urp;StarOffice.ComponentContext")
                    break
                except Exception:
                    if time.monotonic() >= deadline or process.poll() is not None:
                        raise RuntimeError("The isolated LibreOffice edit process did not become ready.")
                    time.sleep(0.1)
            desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
            document = desktop.loadComponentFromURL(source.resolve().as_uri(), "_blank", 0,
                                                    (prop("Hidden", True), prop("ReadOnly", False),
                                                     prop("UpdateDocMode", 0), prop("MacroExecutionMode", 4)))
            if document is None:
                raise RuntimeError("The native exported PPTX did not open in Impress.")
            shapes = list(walk(document.getDrawPages().getByIndex(0)))
            text = next(shape for shape in shapes if hasattr(shape, "getString") and "项目进展" in shape.getString())
            text.setString("外部编辑确认：项目进展")
            position = text.getPosition()
            position.X += 200
            text.setPosition(position)
            table = next(shape for shape in shapes if shape.getShapeType() == "com.sun.star.drawing.TableShape")
            table.Model.getCellByPosition(0, 0).setString("外部表格")
            chart_shape = next(shape for shape in shapes if shape.supportsService("com.sun.star.drawing.OLE2Shape"))
            chart = chart_shape.Model
            provider = chart.getDataProvider()
            old_values = provider.getData()
            values = [list(row) for row in old_values]
            values[0][0] = 321.0
            provider.setData(tuple(tuple(row) for row in values))
            notes = document.getDrawPages().getByIndex(0).getNotesPage()
            note = next(shape for shape in walk(notes) if hasattr(shape, "getString") and "演讲者提示" in shape.getString())
            note.setString("外部备注确认")
            document.storeAsURL(target.resolve().as_uri(), (prop("FilterName", "Impress Office Open XML"), prop("Overwrite", True)))
            return {"text": text.getString(), "table": table.Model.getCellByPosition(0, 0).getString(),
                    "chartFirstValue": provider.getData()[0][0], "notes": note.getString(),
                    "movedXHundredthMm": 200, "pages": document.getDrawPages().getCount()}
        finally:
            if document is not None:
                document.close(True)
            try:
                os.killpg(process.pid, signal.SIGTERM)
            except ProcessLookupError:
                pass
            try:
                process.communicate(timeout=3)
            except subprocess.TimeoutExpired:
                os.killpg(process.pid, signal.SIGKILL)
                process.communicate()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("target", type=Path)
    args = parser.parse_args()
    print(json.dumps(edit(args.source, args.target), ensure_ascii=False))
