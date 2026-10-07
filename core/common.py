"""Utilitas bersama: ekstrak ZIP, cari file, recalc LibreOffice, logger."""
import glob
import os
import shutil
import subprocess
import tempfile
import zipfile


class Log:
    """Pengganti print() di notebook: menampung pesan supaya bisa ditampilkan di web."""

    def __init__(self):
        self.lines = []

    def __call__(self, *parts):
        msg = " ".join(str(p) for p in parts)
        self.lines.append(msg)
        print(msg)

    def text(self):
        return "\n".join(self.lines)


def extract_all(zip_path, extract_to):
    """Ekstrak ZIP (termasuk ZIP bersarang). Kembalikan daftar semua file hasil ekstrak."""
    os.makedirs(extract_to, exist_ok=True)
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(extract_to)
    out = []
    for root, _, files in os.walk(extract_to):
        for f in files:
            full = os.path.join(root, f)
            if f.lower().endswith(".zip"):
                out.extend(extract_all(full, full + "_extracted"))
            else:
                out.append(full)
    return out


def prepare_inputs(input_dir):
    """Ekstrak semua ZIP di input_dir, kembalikan daftar semua file (non-ZIP, bukan file sistem)."""
    files = []
    for p in sorted(glob.glob(os.path.join(input_dir, "*"))):
        if os.path.isdir(p):
            continue
        if p.lower().endswith(".zip"):
            files.extend(extract_all(p, p + "_extracted"))
        else:
            files.append(p)
    return [
        f for f in files
        if "__MACOSX" not in f and not os.path.basename(f).startswith(("~$", "._"))
    ]


def find_soffice():
    path = shutil.which("soffice") or shutil.which("libreoffice")
    if path:
        return path
    for cand in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if os.path.exists(cand):
            return cand
    return None


def recalc_with_libreoffice(path, log=print, timeout=120):
    """Hitung ulang formula lewat LibreOffice headless supaya nilainya ter-cache.
    Kalau LibreOffice tidak ada, dilewati (file tetap valid; Excel menghitung saat dibuka)."""
    soffice = find_soffice()
    if soffice is None:
        log("LibreOffice tidak ditemukan -- lewati recalculate. "
            "Excel akan menghitung formula saat file dibuka.")
        return False
    abs_path = os.path.abspath(path)
    with tempfile.TemporaryDirectory() as tmpdir:
        profile = "file://" + os.path.join(tmpdir, "lo_profile").replace("\\", "/")
        cmd = [soffice, f"-env:UserInstallation={profile}", "--headless", "--calc",
               "--convert-to", "xlsx", "--outdir", tmpdir, abs_path]
        try:
            result = subprocess.run(cmd, timeout=timeout, capture_output=True, text=True)
            converted = os.path.join(tmpdir, os.path.basename(abs_path))
            if result.returncode == 0 and os.path.exists(converted) and os.path.getsize(converted) > 0:
                shutil.move(converted, abs_path)
                log("Recalculate selesai, semua formula sudah ter-cache nilainya.")
                return True
            log(f"Recalculate gagal (returncode={result.returncode}): {result.stderr.strip()[:300]}")
        except Exception as e:  # noqa: BLE001
            log(f"Recalculate gagal ({e}) -- file tetap tersimpan, formula belum ter-cache.")
    return False
