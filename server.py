import socket
import uvicorn


def is_port_in_use(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        return sock.connect_ex(("localhost", port)) == 0


if __name__ == "__main__":
    port = 7656
    if not is_port_in_use(port):
        uvicorn.run("app.main:app", host="localhost", port=port, reload=True)
    else:
        print(f"Port {port} is already in use. Please free up the port and try again.")
