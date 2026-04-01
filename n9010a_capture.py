# -*- coding: utf-8 -*-
"""
Keysight N9010A screenshot helper.
Designed as a fault snapshot tool so the primary system can capture the current
screen without disturbing the instrument configuration.
"""
from __future__ import annotations

import argparse
import socket
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

DEFAULT_FILENAME_TEMPLATE = "failure_screenshot_%Y%m%d_%H%M%S.png"
INSTRUMENT_PORT = 5025


def send(sock: socket.socket, cmd: str) -> None:
    if not cmd.endswith("\n"):
        cmd += "\n"
    sock.sendall(cmd.encode("ascii", errors="ignore"))


def recv_line(sock: socket.socket, chunk: int = 65536, timeout: Optional[float] = None) -> str:
    if timeout is not None:
        sock.settimeout(timeout)
    data = b""
    while b"\n" not in data:
        part = sock.recv(chunk)
        if not part:
            break
        data += part
    if timeout is not None:
        sock.settimeout(None)
    return data.split(b"\n", 1)[0].decode("utf-8", errors="ignore").strip()


def recv_exact(sock: socket.socket, size: int, timeout: Optional[float] = None) -> bytes:
    if timeout is not None:
        sock.settimeout(timeout)
    buf = bytearray()
    while len(buf) < size:
        part = sock.recv(min(65536, size - len(buf)))
        if not part:
            raise RuntimeError("Socket closed before receiving expected bytes")
        buf.extend(part)
    if timeout is not None:
        sock.settimeout(None)
    return bytes(buf)


def query_bin_block(sock: socket.socket, cmd: str, timeout: Optional[float] = None) -> bytes:
    send(sock, cmd)
    head = recv_exact(sock, 2, timeout=timeout)
    if head[0:1] != b"#" or not head[1:2].isdigit():
        raise RuntimeError(f"Invalid block header: {head!r}")
    digits = int(head[1:2].decode())
    if digits == 0:
        raise NotImplementedError("Indefinite-length block not supported")
    length_bytes = recv_exact(sock, digits, timeout=timeout)
    expected = int(length_bytes.decode())
    payload = recv_exact(sock, expected, timeout=timeout)

    # Drain any trailing newline so future reads behave.
    sock.settimeout(0.2)
    try:
        while True:
            tail = sock.recv(65536)
            if not tail or len(tail) < 65536:
                break
    except socket.timeout:
        pass
    finally:
        sock.settimeout(None)
    return payload


def idn(sock: socket.socket) -> str:
    send(sock, "*IDN?")
    return recv_line(sock)


def opc_wait(sock: socket.socket, timeout: float = 10.0) -> None:
    send(sock, "*OPC?")
    _ = recv_line(sock, timeout=timeout)


def capture_sdump(
    sock: socket.socket,
    center: Optional[float],
    span: Optional[float],
    do_sweep: bool,
    opctimeout: float,
    stream_timeout: float,
    verbose: bool,
) -> bytes:
    configure = any(value is not None for value in (center, span)) or do_sweep
    if configure and verbose:
        print("[INFO] Configuring instrument before SDUMp capture")

    if center is not None:
        send(sock, f":SENS:FREQ:CENT {center}")
    if span is not None:
        send(sock, f":SENS:FREQ:SPAN {span}")
    if do_sweep:
        send(sock, ":INIT:IMM")
        opc_wait(sock, timeout=opctimeout)

    if verbose:
        print("[INFO] Requesting SDUMp screenshot stream")
    return query_bin_block(sock, ":HCOPy:SDUMp:DATA? PNG", timeout=stream_timeout)


def capture_mmem(
    sock: socket.socket,
    center: Optional[float],
    span: Optional[float],
    do_sweep: bool,
    opctimeout: float,
    stream_timeout: float,
    verbose: bool,
) -> bytes:
    if verbose:
        print("[INFO] Configuring instrument before MMEM capture")

    if center is not None:
        send(sock, f":SENS:FREQ:CENT {center}")
    if span is not None:
        send(sock, f":SENS:FREQ:SPAN {span}")

    send(sock, ":STOP")
    time.sleep(0.05)

    if do_sweep:
        send(sock, ":INIT:IMM")
        opc_wait(sock, timeout=opctimeout)
        send(sock, ":STOP")

    filename = r"D:\temp\scr.png"
    send(sock, f':MMEM:STOR:SCR "{filename}"')
    opc_wait(sock, timeout=max(10.0, opctimeout))

    if verbose:
        print(f"[INFO] Reading screenshot from {filename}")
    png = query_bin_block(sock, f':MMEM:DATA? "{filename}"', timeout=max(10.0, stream_timeout))

    try:
        send(sock, f':MMEM:DEL "{filename}"')
    except Exception:
        pass

    return png


def _single_capture(
    ip: str,
    method: str,
    center: Optional[float],
    span: Optional[float],
    timeout: float,
    do_sweep: bool,
    opctimeout: float,
    stream_timeout: float,
    verbose: bool,
) -> bytes:
    with socket.create_connection((ip, INSTRUMENT_PORT), timeout=timeout) as sock:
        sock.settimeout(timeout)
        ident = idn(sock)
        if verbose:
            print(f"[IDN] {ident}")

        if method == "sdump":
            return capture_sdump(sock, center, span, do_sweep, opctimeout, stream_timeout, verbose)
        if method == "mmem":
            return capture_mmem(sock, center, span, do_sweep, opctimeout, stream_timeout, verbose)
        raise ValueError(f"Unsupported capture method: {method}")


def save_screenshot(
    ip: str,
    output_path: Optional[str] = None,
    *,
    center: Optional[float] = None,
    span: Optional[float] = None,
    timeout: float = 5.0,
    do_sweep: bool = False,
    method: str = "auto",
    opctimeout: float = 10.0,
    stream_timeout: float = 10.0,
    verbose: bool = False,
) -> bool:
    if output_path is None:
        timestamp = datetime.now().strftime(DEFAULT_FILENAME_TEMPLATE)
        path = Path.cwd() / timestamp
    else:
        path = Path(output_path)

    target_method = method.lower()
    if target_method not in {"auto", "sdump", "mmem"}:
        raise ValueError("method must be one of: auto, sdump, mmem")

    try:
        png_data: Optional[bytes] = None
        attempt_errors: list[str] = []

        if target_method in {"auto", "sdump"}:
            try:
                png_data = _single_capture(
                    ip=ip,
                    method="sdump",
                    center=center,
                    span=span,
                    timeout=timeout,
                    do_sweep=do_sweep,
                    opctimeout=opctimeout,
                    stream_timeout=stream_timeout,
                    verbose=verbose,
                )
            except Exception as exc:
                attempt_errors.append(f"SDUMp capture error: {exc}")
                if verbose:
                    print(f"[WARN] SDUMp capture failed: {exc}")

        if png_data is None and target_method in {"auto", "mmem"}:
            should_remain_passive = not any(value is not None for value in (center, span)) and not do_sweep
            if should_remain_passive and verbose:
                print("[WARN] MMEM capture will momentarily halt the instrument display (SCPI :STOP).")
            try:
                png_data = _single_capture(
                    ip=ip,
                    method="mmem",
                    center=center,
                    span=span,
                    timeout=timeout,
                    do_sweep=do_sweep,
                    opctimeout=opctimeout,
                    stream_timeout=stream_timeout,
                    verbose=verbose,
                )
            except Exception as exc:
                attempt_errors.append(f"MMEM capture error: {exc}")
                if verbose:
                    print(f"[WARN] MMEM capture failed: {exc}")

        if png_data is None:
            if attempt_errors:
                raise RuntimeError("; ".join(attempt_errors))
            raise RuntimeError("Screenshot capture failed; no data returned.")

        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(png_data)
        if verbose:
            print(f"[OK] Screenshot saved to {path}")
        return True
    except Exception as exc:
        if verbose:
            print(f"[ERROR] Screenshot capture failed: {exc}")
        return False


def main() -> None:
    parser = argparse.ArgumentParser(description="Capture a screenshot from a Keysight N9010A.")
    parser.add_argument("--ip", required=True, help="Instrument IP address")
    parser.add_argument("--out", help="Optional output file path. Auto-generates when omitted.")
    parser.add_argument("--center", type=float, help="Center frequency in Hz (optional)")
    parser.add_argument("--span", type=float, help="Span in Hz (optional)")
    parser.add_argument("--timeout", type=float, default=5.0, help="Socket timeout in seconds")
    parser.add_argument("--sweep", action="store_true", help="Trigger a single sweep before capture")
    parser.add_argument(
        "--method",
        choices=["auto", "sdump", "mmem"],
        default="auto",
        help="Capture method to use",
    )
    parser.add_argument("--opctimeout", type=float, default=10.0, help="*OPC? wait timeout in seconds")
    parser.add_argument(
        "--stream-timeout",
        type=float,
        default=10.0,
        help="Binary data streaming timeout in seconds",
    )
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging")
    args = parser.parse_args()

    success = save_screenshot(
        ip=args.ip,
        output_path=args.out,
        center=args.center,
        span=args.span,
        timeout=args.timeout,
        do_sweep=args.sweep,
        method=args.method,
        opctimeout=args.opctimeout,
        stream_timeout=args.stream_timeout,
        verbose=args.verbose,
    )
    raise SystemExit(0 if success else 1)


if __name__ == "__main__":
    main()
