# Notepad++ Automated UI Tests

This project contains an automated UI test suite for Notepad++. It uses Python with PyAutoGUI for UI interaction and PyTest as the testing framework. The tests cover basic find, replace, and replace all functionalities, as well as dialog interactions.

## Features Tested

* **Find Functionality:**
    * Positive test: Finding an existing word.
    * Negative test: Attempting to find a non-existing word and verifying the "not found" indication.
* **Replace Functionality (Single Instance):**
    * Replacing an existing word.
    * Replacing a different existing word.
    * Attempting to replace a non-existent word (text should remain unchanged).
    * Replacing a word with an empty string.
* **Replace All Functionality:**
    * Replacing all occurrences of a word (case-insensitive).
    * Replacing all occurrences of a different word.
    * Replacing all occurrences of a word, demonstrating case-insensitivity.
    * Attempting to replace all occurrences of a non-existent word.
* **Replace Dialog Interaction:**
    * Opening the Replace dialog.
    * Closing the Replace dialog using its dedicated close button.

## Prerequisites

* **Python:** Version 3.7+ recommended.
* **Notepad++:** Installed on the system. The script is configured for a specific installation path but can be adjusted.
* **pip:** Python package installer.
* **Monitor/Display:** The tests interact with the UI, so a display is required. It's best to run these on a machine where the Notepad++ window can be active and unobstructed.
* **Screen Resolution & Theme:** Image recognition is sensitive to screen resolution, DPI scaling, and application themes. The provided UI element images might need to be recaptured on your specific system if they are not recognized.

## Setup

1.  **Clone the Repository (or download the script):**
    If this were a full project:
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```
    Otherwise, ensure you have the Python script (`.py` file).

2.  **Install Python:**
    If you don't have Python installed, download it from [python.org](https://www.python.org/) and ensure it's added to your system's PATH.

3.  **Create a Virtual Environment (Recommended):**
    ```bash
    python -m venv venv
    # On Windows
    .\venv\Scripts\activate
    # On macOS/Linux
    source venv/bin/activate
    ```

4.  **Install Dependencies:**
    Create a `requirements.txt` file with the following content:
    ```txt
    pytest
    pyautogui
    pyperclip
    opencv-python
    ```
    Then install them:
    ```bash
    pip install -r requirements.txt
    ```
    * `pytest`: The testing framework.
    * `pyautogui`: For GUI automation.
    * `pyperclip`: For copy/paste operations.
    * `opencv-python`: Used by PyAutoGUI for more reliable image recognition (confidence scores, grayscale).

5.  **Install Notepad++:**
    Download and install Notepad++ from its [official website](https://notepad-plus-plus.org/).

6.  **Configure Notepad++ Path:**
    Open the test script (e.g., `notepad_plus_plus_tests.py`) and update the `NOTEPAD_PLUS_PLUS_PATH` variable if your Notepad++ installation path is different:
    ```python
    NOTEPAD_PLUS_PLUS_PATH = r"C:\Program Files\Notepad++\notepad++.exe" # Adjust if necessary
    ```

7.  **Prepare UI Element Images:**
    * This script relies on image recognition. You need a folder named `ui_elements` in the same directory as the test script.
    * This folder must contain the necessary PNG images of Notepad++ UI elements (buttons, menu items, dialog indicators). The script references images like:
        * `search_menu_item.png`
        * `replace_submenu_item.png`
        * `find_next_button.png`
        * `replace_action_button.png`
        * `replace_all_button.png`
        * `replace_dialog_close_button.png`
        * `dont_save_button.png`
        * `find_success_indicator.png` (e.g., highlighted text or a specific part of the UI indicating success)
        * `find_text_not_found_dialog.png` (e.g., the "text not found" pop-up message)
    * **Important:** You will likely need to capture these images from your own Notepad++ instance, as UI appearance can vary based on version, theme, and OS settings. Ensure they are cropped precisely to the element you want to identify.

8.  **Create Screenshot Directory:**
    The script will automatically create a directory named `test_screenshots` (if it doesn't exist) in the same location as the script. All screenshots taken during the test execution (success, error, debug) will be saved here.

## Directory Structure

Your project directory should look something like this:
Okay, here's a README.md file in English for the Python test script. This README assumes the script is named something like notepad_plus_plus_tests.py.

Markdown

# Notepad++ Automated UI Tests

This project contains an automated UI test suite for Notepad++. It uses Python with PyAutoGUI for UI interaction and PyTest as the testing framework. The tests cover basic find, replace, and replace all functionalities, as well as dialog interactions.

## Features Tested

* **Find Functionality:**
    * Positive test: Finding an existing word.
    * Negative test: Attempting to find a non-existing word and verifying the "not found" indication.
* **Replace Functionality (Single Instance):**
    * Replacing an existing word.
    * Replacing a different existing word.
    * Attempting to replace a non-existent word (text should remain unchanged).
    * Replacing a word with an empty string.
* **Replace All Functionality:**
    * Replacing all occurrences of a word (case-insensitive).
    * Replacing all occurrences of a different word.
    * Replacing all occurrences of a word, demonstrating case-insensitivity.
    * Attempting to replace all occurrences of a non-existent word.
* **Replace Dialog Interaction:**
    * Opening the Replace dialog.
    * Closing the Replace dialog using its dedicated close button.

## Prerequisites

* **Python:** Version 3.7+ recommended.
* **Notepad++:** Installed on the system. The script is configured for a specific installation path but can be adjusted.
* **pip:** Python package installer.
* **Monitor/Display:** The tests interact with the UI, so a display is required. It's best to run these on a machine where the Notepad++ window can be active and unobstructed.
* **Screen Resolution & Theme:** Image recognition is sensitive to screen resolution, DPI scaling, and application themes. The provided UI element images might need to be recaptured on your specific system if they are not recognized.

## Setup

1.  **Clone the Repository (or download the script):**
    If this were a full project:
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```
    Otherwise, ensure you have the Python script (`.py` file).

2.  **Install Python:**
    If you don't have Python installed, download it from [python.org](https://www.python.org/) and ensure it's added to your system's PATH.

3.  **Create a Virtual Environment (Recommended):**
    ```bash
    python -m venv venv
    # On Windows
    .\venv\Scripts\activate
    # On macOS/Linux
    source venv/bin/activate
    ```

4.  **Install Dependencies:**
    Create a `requirements.txt` file with the following content:
    ```txt
    pytest
    pyautogui
    pyperclip
    opencv-python
    ```
    Then install them:
    ```bash
    pip install -r requirements.txt
    ```
    * `pytest`: The testing framework.
    * `pyautogui`: For GUI automation.
    * `pyperclip`: For copy/paste operations.
    * `opencv-python`: Used by PyAutoGUI for more reliable image recognition (confidence scores, grayscale).

5.  **Install Notepad++:**
    Download and install Notepad++ from its [official website](https://notepad-plus-plus.org/).

6.  **Configure Notepad++ Path:**
    Open the test script (e.g., `notepad_plus_plus_tests.py`) and update the `NOTEPAD_PLUS_PLUS_PATH` variable if your Notepad++ installation path is different:
    ```python
    NOTEPAD_PLUS_PLUS_PATH = r"C:\Program Files\Notepad++\notepad++.exe" # Adjust if necessary
    ```

7.  **Prepare UI Element Images:**
    * This script relies on image recognition. You need a folder named `ui_elements` in the same directory as the test script.
    * This folder must contain the necessary PNG images of Notepad++ UI elements (buttons, menu items, dialog indicators). The script references images like:
        * `search_menu_item.png`
        * `replace_submenu_item.png`
        * `find_next_button.png`
        * `replace_action_button.png`
        * `replace_all_button.png`
        * `replace_dialog_close_button.png`
        * `dont_save_button.png`
        * `find_success_indicator.png` (e.g., highlighted text or a specific part of the UI indicating success)
        * `find_text_not_found_dialog.png` (e.g., the "text not found" pop-up message)
    * **Important:** You will likely need to capture these images from your own Notepad++ instance, as UI appearance can vary based on version, theme, and OS settings. Ensure they are cropped precisely to the element you want to identify.

8.  **Create Screenshot Directory:**
    The script will automatically create a directory named `test_screenshots` (if it doesn't exist) in the same location as the script. All screenshots taken during the test execution (success, error, debug) will be saved here.

## Directory Structure

Your project directory should look something like this:

/your_project_folder
|-- notepad_plus_plus_tests.py  # Your main test script
|-- /ui_elements/               # Folder for source UI element images
|   |-- search_menu_item.png
|   |-- replace_submenu_item.png
|   |-- ... (other UI images)
|-- /test_screenshots/          # Folder where test result screenshots will be saved (created automatically)
|-- requirements.txt
|-- README.md


## Running the Tests

1.  **Navigate to the script's directory** in your terminal or command prompt.
2.  **Ensure your virtual environment is activated** (if you created one).
3.  **Do not interact with the mouse or keyboard** while the tests are running, as PyAutoGUI controls them. Ensure the Notepad++ window (once opened by the script) remains active and unobstructed.
4.  **Run PyTest:**
    ```bash
    pytest
    ```
    Or, to be more specific if you have multiple test files:
    ```bash
    pytest notepad_plus_plus_tests.py
    ```
    PyTest will discover and run the test functions in the script. You'll see output indicating the status of each test.

## Notes and Troubleshooting

* **Image Recognition Failures:** If tests fail because images are not found, try re-capturing the relevant images from the `ui_elements` folder on your system with your current Notepad++ theme and resolution. Ensure screenshots are clear and tightly cropped.
* **Timing Issues:** If tests are flaky, you might need to adjust the `ACTION_DELAY` or `INITIAL_APP_WAIT_TIME` constants in the script.
* **Screen Resolution/Scaling:** High DPI screens or custom scaling can affect PyAutoGUI's coordinate system and image matching. It's generally best to run these tests with 100% scaling.
* **Notepad++ Language:** The script is primarily designed for an English version of Notepad++, though some dialog title checks include Russian alternatives for robustness. If your Notepad++ uses a different language, image matching might be more reliable than title checks for dialogs.
* **Fail-Safe:** PyAutoGUI has a fail-safe feature: rapidly move your mouse to any corner of the screen to stop execution if something goes wrong.
