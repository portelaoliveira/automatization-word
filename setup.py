from pathlib import Path
import sys
import shutil
import subprocess
import shlex
from cx_Freeze import setup, Executable

name = "SARs"
version = "4.3"
build_folder = Path(f"build/{name} v{version}")

src_path = Path("src")
resources_path = Path("resources")
qss_files = src_path.rglob("*.qss")
includes = [(qss, resources_path / "styles" / qss.name) for qss in qss_files]
includes.append(("venv/Lib/site-packages/pyproj.libs", "lib/pyproj.libs"))
includes.append("templates")
includes.append("tutorial/Tutorial.pdf")
includes.append("targets")
if resources_path.exists():
    includes += [resources_path]
includes.append(("bin/hugin", "bin/hugin"))

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
    "build_exe": build_folder,
    "excludes": [
        "tkinter",
        "PyQt5.QtBluetooth",
        "PyQt5.QtNetwork",
        "PyQt5.QtNfc",
        "PyQt5.QtWebChannel",
        "PyQt5.QtWebEngine",
        "PyQt5.QtWebEngineCore",
        "PyQt5.QtWebEngineWidgets",
        "PyQt5.QtWebKit",
        "PyQt5.QtWebKitWidgets",
        "PyQt5.QtWebsockets",
        "PyQt5.QtSql",
        "PyQt5.QtScript",
    ],
    "include_files": includes,
    "zip_include_packages": "*",
    "zip_exclude_packages": ["numpy", "pyproj"],
}

# base="Win32GUI" should be used only for Windows GUI app
base = None
if sys.platform == "win32":
    base = "Win32GUI"

icon = Path() / "resources" / "icons" / "logo.ico"
if not icon.exists():
    logo = icon.parent / "logo.svg"
    imagemagick = Path() / "bin" / "imagemagick" / "magick.exe"
    cmd = (
        f'"{imagemagick}" convert -background none "{logo}" -define'
        f' icon:auto-resize "{icon}"'
    )
    subprocess.run(shlex.split(cmd), check=True)

setup(
    name=name,
    version=version,
    description=f"{name}!",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "main.py",
            base=base,
            target_name=f"{name}.exe",
            icon=icon,
        )
    ],
)

build_lib = build_folder.parent / "lib"
if build_lib.is_dir():
    shutil.rmtree(build_lib)

pyqt5_lib_path = Path(build_exe_options["build_exe"]) / "lib/PyQt5/Qt5"
pyqt5_lib_files = (
    (
        "venv/Lib/site-packages/PyQt5/Qt5/bin/Qt5Core.dll",
        "lib/PyQt5/Qt5/bin/Qt5Core.dll",
    ),
    (
        "venv/Lib/site-packages/PyQt5/Qt5/bin/Qt5Gui.dll",
        "lib/PyQt5/Qt5/bin/Qt5Gui.dll",
    ),
    (
        "venv/Lib/site-packages/PyQt5/Qt5/bin/Qt5Svg.dll",
        "lib/PyQt5/Qt5/bin/Qt5Svg.dll",
    ),
    (
        "venv/Lib/site-packages/PyQt5/Qt5/bin/Qt5Widgets.dll",
        "lib/PyQt5/Qt5/bin/Qt5Widgets.dll",
    ),
    (
        "venv/Lib/site-packages/PyQt5/Qt5/plugins/iconengines/qsvgicon.dll",
        "lib/PyQt5/Qt5/plugins/iconengines/qsvgicon.dll",
    ),
    (
        "venv/Lib/site-packages/PyQt5/Qt5/plugins/platforms/qwindows.dll",
        "lib/PyQt5/Qt5/plugins/platforms/qwindows.dll",
    ),
)

if pyqt5_lib_path.exists():
    shutil.rmtree(pyqt5_lib_path)
    for f in pyqt5_lib_files:
        in_file = f[0]
        out_file = build_folder / f[1]
        out_file.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(in_file, out_file)

z7_path = Path(f"{build_folder}.7z")

if z7_path.exists():
    usr_in = input(
        f'Arquivo "{z7_path}" já existe, deseja substituir? [S/n] '
    ).strip()
    if usr_in == "" or usr_in[0].lower() == "s":
        z7_path.unlink()
    else:
        print("O executável não foi compactado!")
        exit()

cmd = shlex.split(
    "./bin/7zr.exe a -t7z -m0=lzma -mx=9 -mfb=64 -md=32m -ms=on "
    f'"{z7_path.name}" "{build_folder.name}"'
)

subprocess.run(cmd, cwd=build_folder.parent)
