import subprocess
import os
import sys
import time
import signal

def run_backend():
    print("Starting Backend (Flask on port 1501)...")
    return subprocess.Popen([sys.executable, "backend/app.py"])

def run_frontend():
    print("Starting Frontend (Vite)...")
    # On Windows, we need shell=True to run npm
    return subprocess.Popen(["npm", "run", "dev"], cwd="frontend", shell=True)

if __name__ == "__main__":
    backend_proc = run_backend()
    frontend_proc = run_frontend()
    
    print("\n" + "="*50)
    print("SERVICES ARE RUNNING")
    print("Backend: http://localhost:1501")
    print("Frontend: http://localhost:1500")
    print("Press CTRL+C to stop both services")
    print("="*50 + "\n")
    
    try:
        while True:
            time.sleep(1)
            # Check if processes are still running
            if backend_proc.poll() is not None:
                print("Backend service stopped unexpectedly.")
                break
            if frontend_proc.poll() is not None:
                print("Frontend service stopped unexpectedly.")
                break
    except KeyboardInterrupt:
        print("\nShutting down services...")
    finally:
        # Gracefully terminate both processes
        backend_proc.terminate()
        frontend_proc.terminate()
        
        # Wait for them to exit
        backend_proc.wait()
        frontend_proc.wait()
        print("All services stopped.")
