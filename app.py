import http.server
import socketserver
import webbrowser
import os
import sys
import threading
import time

PORT = 5000
HOST = "127.0.0.1"

class CustomHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # Add CORS and caching headers for smooth local development
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        super().end_headers()

def open_browser(url):
    time.sleep(1.2)
    print(f"Opening {url} in your default browser...")
    webbrowser.open(url)

def main():
    # Change working directory to script location
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    server_address = (HOST, PORT)
    
    try:
        httpd = socketserver.TCPServer(server_address, CustomHandler)
    except OSError:
        # Fallback to port 8080 if 5000 is occupied
        server_address = (HOST, 8080)
        httpd = socketserver.TCPServer(server_address, CustomHandler)
        port_used = 8080
    else:
        port_used = PORT

    url = f"http://{HOST}:{port_used}"
    
    print("\n" + "=" * 60)
    print("  STUDIO SQL - BASIC DATABASE & FILE STUDIO SERVER")
    print("=" * 60)
    print(f"  Status   : Server is Live and Running!")
    print(f"  URL      : {url}")
    print(f"  Directory: {os.getcwd()}")
    print("=" * 60)
    print("  Press Ctrl+C in this console to stop the server.\n")

    # Launch browser in a separate thread
    threading.Thread(target=open_browser, args=(url,), daemon=True).start()

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping Studio SQL Server... Goodbye!")
        httpd.server_close()
        sys.exit(0)

if __name__ == "__main__":
    main()
