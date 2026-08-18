# DIY Video Compilation Maker

An automated Python script designed to build long-form video compilations from a Google Drive repository. The script processes individual clips, standardizes their format (resolution, framerate, and audio) using FFmpeg, concats them into a single file, mixes in a background music (BGM) track, and outputs a finished video alongside an auto-generated thumbnail grid.

## Prerequisites

Before running this script, ensure you have the following installed on your machine:
*   **Python 3.7+**
*   **FFmpeg:** This must be installed and accessible in your system's PATH.

## Installation & Setup

1. **Clone or Download the Repository:** Save `compilation_maker.py` to your local machine.
2. **Install Python Dependencies:**
   Run the following command in your terminal to install the necessary libraries:
   `pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib openpyxl requests pillow numpy`
3. **Configure Local Paths:**
   Open `compilation_maker.py` and update the following variables under the `CONFIG` section with paths specific to your environment:
   *   `BGM_FOLDER`: Set this to the folder containing your MP3/WAV background music files.
   *   `EXCEL_PATH`: Set this to the location where your log/input Excel spreadsheet will be saved and read from.
4. **Google Drive API Credentials:**
   *   You need a Google Cloud project with the Google Drive API enabled.
   *   Generate OAuth 2.0 Client ID credentials and download the JSON file.
   *   Rename the downloaded file to `client_secret.json` and place it in the same directory as `compilation_maker.py`.
   *   **Important:** Do NOT commit `client_secret.json` or the auto-generated `token.pickle` to version control. 

## How to Use

1. **Populate the Input Sheet:** 
   *   Run the script once to automatically generate the Excel file template.
   *   In the generated Excel file, go to the `input` tab.
   *   Add your target videos (Channel Name, Filename, Google Drive Link).
2. **Run the Script:**
   `python compilation_maker.py`
3. **Follow the On-Screen Prompts:**
   *   Choose your aspect ratio (16:9 Horizontal or 9:16 Vertical).
   *   Enter the target channel names separated by commas (or leave blank).
   *   Specify how many compilation videos you want to generate.
   
The script will download the clips, normalize them to prevent frame/audio drops, generate the video, add music, and create a 3-panel thumbnail grid in your output folder.
