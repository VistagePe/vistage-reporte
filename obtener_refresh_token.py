"""
Ejecutá este script UNA SOLA VEZ en tu computadora para obtener el Refresh Token.
El Refresh Token es lo que necesitás pegar en Streamlit Cloud.
"""

import requests, webbrowser, threading, time
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, urlencode

# ── Pegá acá tus credenciales ─────────────────────────────────────
CLIENT_ID     = "TU_CLIENT_ID_AQUI"
CLIENT_SECRET = "TU_CLIENT_SECRET_AQUI"
# ─────────────────────────────────────────────────────────────────

REDIRECT_URI = "http://localhost:8080/callback"
SCOPES       = "ZohoCreator.report.READ,ZohoCreator.form.READ"

auth_code = None

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        global auth_code
        params = parse_qs(urlparse(self.path).query)
        auth_code = params.get("code", [None])[0]
        self.send_response(200)
        self.send_header("Content-type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"<h2>Listo! Ya podes cerrar esta pestana.</h2>")
    def log_message(self, *a): pass

def main():
    if CLIENT_ID == "TU_CLIENT_ID_AQUI":
        print("Completa CLIENT_ID y CLIENT_SECRET en este archivo primero.")
        input("Enter para salir...")
        return

    url = "https://accounts.zoho.com/oauth/v2/auth?" + urlencode({
        "response_type": "code", "client_id": CLIENT_ID,
        "scope": SCOPES, "redirect_uri": REDIRECT_URI, "access_type": "offline"
    })

    server = HTTPServer(("localhost", 8080), Handler)
    threading.Thread(target=lambda: server.handle_request(), daemon=True).start()

    print("Abriendo navegador para autorizar...")
    webbrowser.open(url)

    for _ in range(120):
        if auth_code: break
        time.sleep(1)

    server.server_close()

    if not auth_code:
        print("No se recibio autorizacion. Intentá de nuevo.")
        input()
        return

    r = requests.post("https://accounts.zoho.com/oauth/v2/token", data={
        "code": auth_code, "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
        "redirect_uri": REDIRECT_URI, "grant_type": "authorization_code"
    })
    data = r.json()

    if "refresh_token" in data:
        print("\n" + "="*60)
        print("REFRESH TOKEN OBTENIDO EXITOSAMENTE")
        print("="*60)
        print(f"\nREFRESH TOKEN:\n{data['refresh_token']}\n")
        print("Copiá este valor y pegalo en Streamlit Cloud como:")
        print('ZOHO_REFRESH_TOKEN = "...el token..."')
        print("="*60)
    else:
        print(f"Error: {data}")

    input("\nPresiona Enter para cerrar...")

if __name__ == "__main__":
    main()
