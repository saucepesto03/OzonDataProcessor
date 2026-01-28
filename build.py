import subprocess
import sys
import shutil
from pathlib import Path

PROJECT_NAME = "ozon_launcher"
ENTRY_POINT = "launcher.py"

def check_pyinstaller():
    try:
        import PyInstaller  # noqa
        return True
    except ImportError:
        return False

def run_build():
    root = Path(__file__).parent.resolve()
    dist = root / "dist"
    build = root / "build"
    spec = root / f"{PROJECT_NAME}.spec"

    print("=" * 60)
    print("СБОРКА OZON LAUNCHER")
    print("=" * 60)

    if not (root / ENTRY_POINT).exists():
        print(f"❌ Не найден файл {ENTRY_POINT}")
        sys.exit(1)

    if not check_pyinstaller():
        print("❌ PyInstaller не установлен")
        print("Установи командой:")
        print("pip install pyinstaller")
        sys.exit(1)

    # Очистка старых сборок
    for path in [dist, build, spec]:
        if path.exists():
            print(f"🧹 Удаляем {path}")
            if path.is_dir():
                shutil.rmtree(path)
            else:
                path.unlink()

    print("🚀 Запуск PyInstaller...\n")

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--console",
        "--clean",
        "--name", PROJECT_NAME,
        ENTRY_POINT
    ]

    result = subprocess.run(cmd)

    if result.returncode != 0:
        print("\n❌ Ошибка сборки")
        sys.exit(1)

    exe_path = dist / f"{PROJECT_NAME}.exe"

    if exe_path.exists():
        print("\n" + "=" * 60)
        print("✅ СБОРКА ЗАВЕРШЕНА УСПЕШНО")
        print("=" * 60)
        print(f"📦 Файл: {exe_path}")
        print("\nМожно копировать exe на другой ПК (Python не нужен)")
    else:
        print("\n⚠️ exe не найден, но PyInstaller отработал без ошибки")

if __name__ == "__main__":
    run_build()
    input("\nНажмите Enter для выхода...")
