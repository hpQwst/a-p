"""Mede um caso real do pipeline sem alterar os arquivos de origem.

O pico do working set e coletado pelo proprio processo via API do sistema,
porque tracemalloc nao inclui buffers nativos/ZIP. Este script imprime somente
o resultado funcional e o tempo; recebe sempre copias locais dos decks usados
no benchmark.
"""

from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
from io import BytesIO
import json
import os
from pathlib import Path
import time
from zipfile import ZIP_DEFLATED, ZipFile

from ppt_automator.engine import generate_updated_pptx
from worker.processor import analyze_files


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--xlsx", action="append", default=[])
    parser.add_argument("--generate", action="store_true")
    args = parser.parse_args()

    pptx_path = Path(args.pptx).resolve()
    xlsx_paths = [Path(value).resolve() for value in args.xlsx]
    started = time.perf_counter()
    pptx_bytes = pptx_path.read_bytes()
    archive_buffer = BytesIO()
    with ZipFile(archive_buffer, "w", ZIP_DEFLATED) as archive:
        for path in xlsx_paths:
            archive.writestr(path.name, path.read_bytes())
    analysis = analyze_files(pptx_bytes, archive_buffer.getvalue())
    generated_bytes = 0
    if args.generate:
        generated_bytes = len(generate_updated_pptx(pptx_bytes, analysis.plans, targets=analysis.targets))
    elapsed = time.perf_counter() - started
    print(
        json.dumps(
            {
                "pptx": pptx_path.name,
                "pptx_bytes": len(pptx_bytes),
                "xlsx_bytes": sum(path.stat().st_size for path in xlsx_paths),
                "targets": analysis.target_count,
                "sources": analysis.source_count,
                "plans": len(analysis.plans),
                "warnings": len(analysis.warnings),
                "generated_bytes": generated_bytes,
                "elapsed_seconds": round(elapsed, 3),
                "peak_working_set_bytes": _peak_working_set_bytes(),
            },
            ensure_ascii=False,
        ),
        flush=True,
    )


def _peak_working_set_bytes() -> int:
    if os.name != "nt":
        try:
            import resource

            return int(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss) * 1024
        except (ImportError, AttributeError):
            return 0

    class ProcessMemoryCounters(ctypes.Structure):
        _fields_ = [
            ("cb", wintypes.DWORD),
            ("PageFaultCount", wintypes.DWORD),
            ("PeakWorkingSetSize", ctypes.c_size_t),
            ("WorkingSetSize", ctypes.c_size_t),
            ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
            ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
            ("PagefileUsage", ctypes.c_size_t),
            ("PeakPagefileUsage", ctypes.c_size_t),
        ]

    counters = ProcessMemoryCounters()
    counters.cb = ctypes.sizeof(counters)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    psapi = ctypes.WinDLL("psapi", use_last_error=True)
    kernel32.GetCurrentProcess.restype = wintypes.HANDLE
    psapi.GetProcessMemoryInfo.argtypes = [
        wintypes.HANDLE,
        ctypes.POINTER(ProcessMemoryCounters),
        wintypes.DWORD,
    ]
    psapi.GetProcessMemoryInfo.restype = wintypes.BOOL
    handle = kernel32.GetCurrentProcess()
    ok = psapi.GetProcessMemoryInfo(handle, ctypes.byref(counters), counters.cb)
    return int(counters.PeakWorkingSetSize) if ok else 0


if __name__ == "__main__":
    main()
