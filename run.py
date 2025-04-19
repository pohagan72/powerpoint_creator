import os
from waitress import serve
from app import app # Import the Flask app object from your app.py file

if __name__ == '__main__':
    host = '0.0.0.0' # Listen on all network interfaces
    port = 5000      # Port the app will run on
    threads = 8      # Number of threads to handle concurrent requests (adjust as needed)

    print(f"Starting Waitress server on http://{host}:{port}")
    serve(app, host=host, port=port, threads=threads)