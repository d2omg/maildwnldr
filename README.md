## Email Downloader

This is a simple email downloader application built using Python and the Tkinter library for the GUI. The program allows users to connect to an IMAP mail server, select specific folders, and download the emails in EML format to the local machine.

### Features

- Connect to an IMAP mail server and authenticate with username and password.
- Browse and select specific folders from the mail server.
- Download emails from the selected folders in EML format.
- Progress bar to track the download status.
- Option to save the mail server configuration for future use.

### Requirements

- Python 3.x
- Tkinter
- imapclient
- requests

### Installation

1. Clone the repository or download the code as a zip file and extract it to your local machine.
2. Make sure you have Python 3.x installed on your system.
3. Install the required dependencies by running the following command:

   ```
   pip install -r requirements.txt
   ```

### Usage

1. Run the program by executing the `main.py` script.
2. The main window will appear, prompting you to enter the mail server address, username, and password.
3. Click on the "Connect" button to establish a connection to the mail server.
4. If the connection is successful, a new window will open to display a list of available folders on the server.
5. Select the folders from which you want to download emails.
6. Click on the "Download" button to start the download process.
7. A progress window will appear, showing the status of the download.
8. You can cancel the download at any time by clicking the "Cancel" button in the progress window.
9. Once the download is complete, a message will appear, indicating the total number of downloaded emails and folders.
10. You can close the application by closing the main window or clicking on the "X" button.

### Configuration

The program allows you to save the mail server configuration for future use. When you run the program again, it will automatically load the saved configuration from the `config.json` file, if available.

### Disclaimer

Please use this application responsibly and only for authorized access to your own mail server or with the explicit permission of the mail server owner. The developer is not responsible for any misuse or unauthorized access to mail servers using this application.
