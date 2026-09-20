import webbrowser
from app import create_app

app = create_app()

if __name__ == '__main__':
    webbrowser.open("http://127.0.0.1:5000")
    # Run the application on the local network (accessible to other devices on the same Wi-Fi/LAN)
    app.run(host='0.0.0.0', port=5000, debug=False)