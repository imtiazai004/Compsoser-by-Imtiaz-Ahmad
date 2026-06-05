import subprocess, socket, time

# Start Streamlit in background
st = subprocess.Popen([
    "streamlit", "run", "app.py",
    "--server.port=8501",
    "--server.address=0.0.0.0",
    "--server.headless=true",
    "--browser.gatherUsageStats=false"
])

# Wait until Streamlit is actually ready before starting nginx
print("Waiting for Streamlit to start...", flush=True)
for _ in range(60):
    try:
        socket.create_connection(("localhost", 8501), timeout=1).close()
        print("Streamlit ready — starting nginx.", flush=True)
        break
    except Exception:
        time.sleep(1)

# Start nginx in foreground (keeps container alive)
subprocess.run(["nginx", "-g", "daemon off;"])
