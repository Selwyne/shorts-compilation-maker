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

## Configuration & Customization

If you need to tweak the core output settings, you can modify the variables directly inside `compilation_maker.py`:
*   **Audio Adjustments:** To update the background music (BGM) volume or the original file volume, change the values on **lines 33-36**.
*   **Video Duration:** To adjust the target length of your final compiled video (in seconds), update the value on **line 41**.
*   **Max Usage Count:** To allow source clips to be reused in multiple compilations, update `MAX_USAGE_COUNT` on **line 42**.

---

## How It Works (Under the Hood)

The script relies on several automated logic checks to streamline the generation process:

*   **Blank Channel Targeting:** The script filters videos by the channel name provided in the Excel sheet. If you press Enter (leave the prompt blank) or type `null` when asked for a target channel, the script will specifically seek out and process only the videos in your sheet that have a **blank/empty** channel name.
*   **Max Usage Allocation:** By default, `MAX_USAGE_COUNT` is set to `1`. This means once a clip is used in a final video, it is logged and will never be used again. If you change this variable to a higher number (e.g., `3`), the script can pull that exact same clip for up to 3 entirely different compilation outputs before retiring it. 
*   **Infinite Video Generation:** When prompted for "How many videos to generate?", if you leave the input blank, the script enters an infinite loop. It will continue building back-to-back compilation videos non-stop until it completely exhausts the usable clips in your input tab (or until the remaining clips fail to meet the minimum duration). 
*   **Organized Output Structure:** When a video is successfully rendered, the script creates a dedicated folder in your destination path. This folder is named after the target channel (e.g., `ChannelName_Timestamp`). Inside this specific folder, you will find both your final compiled `.mp4` video and the generated 16:9 `.jpg` thumbnail grid.

---

## How to Use

1. **Populate the Input Sheet:** 
   *   Run the script once to automatically generate the Excel file template.
   *   In the generated Excel file, go to the `input` tab.
   *   Add your target videos (Channel Name, Filename, Google Drive Link).
   
   ### Excel Sheet Formatting

   When you run the script for the first time, it will automatically generate a `long-form_logs.xlsx` file. Open this file and go to the **input** tab. 

   You must format your data exactly like this so the script can read the Google Drive links:

   | channel name | filename          | link                                                               |
   | :---         | :---              | :---                                                               |
   | Channel_Name 1       | video_1_craft.mp4 | https://drive.google.com/file/d/1A2B3C4D5E6F7G8H9I0J/view?usp=sharing |
   | Channel_Name 2      | video_2_hack.mp4  | https://drive.google.com/file/d/0J9I8H7G6F5E4D3C2B1A/view?usp=sharing |

   *   **channel name:** Make sure this matches the channel names you input in the terminal exactly.
   *   **filename:** (Optional) Just for your own reference.
   *   **link:** Must be a valid, accessible Google Drive link.

2. **Run the Script:**
   `python compilation_maker.py`
   
3. **Follow the On-Screen Prompts:**
   *   Choose your aspect ratio (16:9 Horizontal or 9:16 Vertical).
   *   Enter the target channel names separated by commas (or leave blank).
   *   Specify how many compilation videos you want to generate.
   
The script will download the clips, normalize them to prevent frame/audio drops, generate the video, add music, and create a 3-panel thumbnail grid in your output folder.
