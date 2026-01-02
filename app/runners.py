import webview

import threading, time

from waitress import serve
import webview

PORT = 8012

def run_flask(app):
    # Run the Flask app with the specified port
    # Set host to '127.0.0.1' for local access only
    serve(app, host='127.0.0.1', port=PORT)
    # app.run(host='127.0.0.1', port=PORT, debug=False, use_reloader=False)

def run_gui(app):

    title = "Excel Comparison Dashboard"

    webview.settings["ALLOW_DOWNLOADS"] = True

    # Start the Flask server in a separate thread
    t = threading.Thread(target=lambda : run_flask(app))
    t.daemon = True
    t.start()

    # Wait a moment for the server to start
    time.sleep(1) 


    try:
        # Create the pywebview window, pointing to the specific port
        url = f"http://127.0.0.1:{PORT}"
        webview.create_window(title, url, maximized=True)
        # Start pywebview's main loop
        webview.start()

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        input("Press Enter to exit...")


def run_browser(app):
    serve(app, host="127.0.0.1", port=5000)  # Run with production mode


def run_dev(app):
    app.run(debug=True, use_reloader=True)
