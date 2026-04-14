"""
network_bridge.py
─────────────────────────────────────────────────────
TCP socket bridge for the Door-Sign-Status system.
Allows the Control PC (Machine A) to push status
updates over Wi-Fi to the Display PC (Machine B).

Uses only Python built-in modules — no extra dependencies.

Server (Display PC):
    server = StatusServer(callback=my_handler)
    server.start()           # non-blocking, runs in daemon thread
    server.stop()

Client (Control PC):
    ok, err = send_status_update("192.168.1.105", "Busy", "3:00 PM")
"""

import socket
import json
import threading
import logging

# ═══════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════

DEFAULT_PORT = 9800
BUFFER_SIZE = 4096
CLIENT_TIMEOUT = 3  # seconds


# ═══════════════════════════════════════════════════════════════════════
#  SERVER — runs on Display PC (Machine B)
# ═══════════════════════════════════════════════════════════════════════

class StatusServer:
    """
    TCP server that listens for JSON status updates and visitor messages.

    When a valid status payload is received, calls:
        callback(status: str, return_time: str, source: str)

    When a valid visitor payload is received, calls:
        visitor_callback(name: str, purpose: str, timestamp: str)

    Callbacks are called from the server's background thread,
    so the caller must handle thread-safety (e.g. use a Queue).
    """

    def __init__(self, host: str = "0.0.0.0", port: int = DEFAULT_PORT,
                 callback=None, visitor_callback=None):
        self.host = host
        self.port = port
        self.callback = callback
        self.visitor_callback = visitor_callback
        self._running = False
        self._thread = None
        self._server_socket = None
        self.last_client_ip = None  # track who connects to us

    def start(self):
        """Start the server in a background daemon thread."""
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(
            target=self._serve_forever,
            daemon=True,
            name="StatusServer"
        )
        self._thread.start()

    def stop(self):
        """Signal the server to stop."""
        self._running = False
        # Create a brief connection to unblock accept()
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(1)
                s.connect(("127.0.0.1", self.port))
        except OSError:
            pass
        if self._server_socket:
            try:
                self._server_socket.close()
            except OSError:
                pass

    def _serve_forever(self):
        """Main server loop — runs in background thread."""
        while self._running:
            try:
                self._server_socket = socket.socket(
                    socket.AF_INET, socket.SOCK_STREAM
                )
                self._server_socket.setsockopt(
                    socket.SOL_SOCKET, socket.SO_REUSEADDR, 1
                )
                self._server_socket.settimeout(5)  # periodic check for _running
                self._server_socket.bind((self.host, self.port))
                self._server_socket.listen(5)

                logging.info(
                    f"StatusServer listening on {self.host}:{self.port}"
                )

                while self._running:
                    try:
                        client, addr = self._server_socket.accept()
                    except socket.timeout:
                        continue  # check _running flag and retry
                    except OSError:
                        break

                    # Handle the client in the same thread (fast JSON payloads)
                    try:
                        self._handle_client(client, addr)
                    except Exception as e:
                        logging.warning(f"Error handling client {addr}: {e}")
                    finally:
                        try:
                            client.close()
                        except OSError:
                            pass

            except OSError as e:
                logging.error(f"StatusServer error: {e}")
                if self._running:
                    import time
                    time.sleep(2)  # wait before retry
            finally:
                if self._server_socket:
                    try:
                        self._server_socket.close()
                    except OSError:
                        pass

    def _handle_client(self, client: socket.socket, addr):
        """Read JSON payload from a single client connection."""
        client.settimeout(5)
        data = b""
        while True:
            chunk = client.recv(BUFFER_SIZE)
            if not chunk:
                break
            data += chunk
            # Simple protocol: one JSON object per connection
            if len(data) > BUFFER_SIZE * 4:
                break  # safety limit

        if not data:
            return

        try:
            payload = json.loads(data.decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError) as e:
            logging.warning(f"Invalid payload from {addr}: {e}")
            # Send error response
            try:
                client.sendall(json.dumps({"ok": False, "error": str(e)}).encode())
            except OSError:
                pass
            return

        msg_type = payload.get("type", "status")

        # Track the client IP (used by display PC to send messages back)
        self.last_client_ip = addr[0]

        # Send acknowledgement
        try:
            client.sendall(json.dumps({"ok": True}).encode())
        except OSError:
            pass

        if msg_type == "visitor":
            # Visitor message from display PC → control PC
            name = payload.get("name", "")
            purpose = payload.get("purpose", "")
            timestamp = payload.get("timestamp", "")

            logging.info(
                f"Visitor message from {addr}: name={name!r}, "
                f"purpose={purpose!r}"
            )

            if self.visitor_callback:
                try:
                    self.visitor_callback(name, purpose, timestamp)
                except Exception as e:
                    logging.error(f"Visitor callback error: {e}")
        else:
            # Status update (default)
            status = payload.get("status", "")
            return_time = payload.get("return_time", "")
            source = payload.get("source", "network")

            logging.info(
                f"Received from {addr}: status={status!r}, "
                f"return_time={return_time!r}"
            )

            if self.callback:
                try:
                    self.callback(status, return_time, source)
                except Exception as e:
                    logging.error(f"Callback error: {e}")

    def get_local_ip(self) -> str:
        """Get this machine's LAN IP address."""
        return get_local_ip()


# ═══════════════════════════════════════════════════════════════════════
#  CLIENT — runs on Control PC (Machine A)
# ═══════════════════════════════════════════════════════════════════════

def send_status_update(ip: str, status: str, return_time: str = "",
                       port: int = DEFAULT_PORT,
                       source: str = "manual") -> tuple:
    """
    Send a status update to the Display PC via TCP.

    Args:
        ip:          IP address of the Display PC
        status:      Status text (e.g. "Busy", "Available")
        return_time: Return time string (e.g. "3:00 PM")
        port:        TCP port (default 9800)
        source:      Source tag ("manual", "outlook", etc.)

    Returns:
        (success: bool, error_message: str)
    """
    payload = json.dumps({
        "type": "status",
        "status": status,
        "return_time": return_time,
        "source": source,
    }).encode("utf-8")

    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(CLIENT_TIMEOUT)
            sock.connect((ip, port))
            sock.sendall(payload)
            sock.shutdown(socket.SHUT_WR)  # signal end of data

            # Wait for acknowledgement
            response = b""
            try:
                response = sock.recv(BUFFER_SIZE)
            except socket.timeout:
                pass

            if response:
                try:
                    ack = json.loads(response.decode("utf-8"))
                    if ack.get("ok"):
                        return (True, "")
                    return (False, ack.get("error", "Unknown server error"))
                except (json.JSONDecodeError, UnicodeDecodeError):
                    pass

            return (True, "")  # no response = assume delivered

    except ConnectionRefusedError:
        return (False, "Connection refused — Display PC may be offline")
    except TimeoutError:
        return (False, "Connection timed out — check IP address and firewall")
    except socket.timeout:
        return (False, "Connection timed out — check IP address and firewall")
    except OSError as e:
        return (False, f"Network error: {e}")


# ═══════════════════════════════════════════════════════════════════════
#  CLIENT — send visitor message from Display PC → Control PC
# ═══════════════════════════════════════════════════════════════════════

def send_visitor_message(ip: str, name: str, purpose: str,
                         timestamp: str = "",
                         port: int = DEFAULT_PORT) -> tuple:
    """
    Send a visitor message from the Display PC to the Control PC.

    Args:
        ip:        IP address of the Control PC
        name:      Visitor's name
        purpose:   Visitor's message
        timestamp: Time the message was left
        port:      TCP port (default 9800)

    Returns:
        (success: bool, error_message: str)
    """
    if not timestamp:
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    payload = json.dumps({
        "type": "visitor",
        "name": name,
        "purpose": purpose,
        "timestamp": timestamp,
    }).encode("utf-8")

    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(CLIENT_TIMEOUT)
            sock.connect((ip, port))
            sock.sendall(payload)
            sock.shutdown(socket.SHUT_WR)

            response = b""
            try:
                response = sock.recv(BUFFER_SIZE)
            except socket.timeout:
                pass

            if response:
                try:
                    ack = json.loads(response.decode("utf-8"))
                    if ack.get("ok"):
                        return (True, "")
                    return (False, ack.get("error", "Unknown server error"))
                except (json.JSONDecodeError, UnicodeDecodeError):
                    pass

            return (True, "")

    except (ConnectionRefusedError, TimeoutError, socket.timeout, OSError) as e:
        return (False, str(e))


# ═══════════════════════════════════════════════════════════════════════
#  UTILITY
# ═══════════════════════════════════════════════════════════════════════

def get_local_ip() -> str:
    """
    Get this machine's LAN IP address.
    Uses a UDP connect trick — no data is actually sent.
    """
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"


# ═══════════════════════════════════════════════════════════════════════
#  SELF-TEST — run this file directly to verify loopback
# ═══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import time

    print("=" * 60)
    print("  Network Bridge — Loopback Self-Test")
    print("=" * 60)

    results = []

    def on_status(status, return_time, source):
        results.append((status, return_time, source))
        print(f"  [SERVER] Received: {status!r}  return_time={return_time!r}")

    # Start server
    server = StatusServer(callback=on_status)
    server.start()
    time.sleep(0.5)

    local_ip = get_local_ip()
    print(f"\n  Local IP: {local_ip}")
    print(f"  Port:     {DEFAULT_PORT}\n")

    # Test 1: Basic status
    print("  Test 1: Send basic status...")
    ok, err = send_status_update("127.0.0.1", "Busy", "3:00 PM")
    time.sleep(0.3)
    assert ok, f"Test 1 failed: {err}"
    assert len(results) == 1
    assert results[-1] == ("Busy", "3:00 PM", "manual")
    print("  ✓ Test 1 passed\n")

    # Test 2: Status with no return time
    print("  Test 2: Send status without return time...")
    ok, err = send_status_update("127.0.0.1", "Available")
    time.sleep(0.3)
    assert ok, f"Test 2 failed: {err}"
    assert len(results) == 2
    assert results[-1] == ("Available", "", "manual")
    print("  ✓ Test 2 passed\n")

    # Test 3: Status with source
    print("  Test 3: Send status with outlook source...")
    ok, err = send_status_update("127.0.0.1", "In a Meeting", source="outlook")
    time.sleep(0.3)
    assert ok, f"Test 3 failed: {err}"
    assert results[-1][2] == "outlook"
    print("  ✓ Test 3 passed\n")

    # Test 4: Connection to bad IP
    print("  Test 4: Connection to unreachable IP (should fail gracefully)...")
    ok, err = send_status_update("192.168.255.254", "Test", port=19999)
    assert not ok
    print(f"  ✓ Test 4 passed — error: {err}\n")

    server.stop()
    time.sleep(0.5)

    print("=" * 60)
    print("  All tests passed! ✓")
    print("=" * 60)
