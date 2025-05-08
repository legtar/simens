import pytest
import pyautogui
import time
import subprocess
import os
import re
import pyperclip

# --- Configuration ---
NOTEPAD_PLUS_PLUS_PATH = r"C:\Program Files\Notepad++\notepad++.exe"
TEXT_TO_TYPE = """     Integrated circuit design, Semiconductor design, chip design or IC design, is a sub-field of Electronics Engineering, encompassing the particular logic and circuit design techniques required to design integrated circuits, or ICs.

     ICs consist of miniaturized electronics components built into an electrical network on a monolithic semiconductor substrate by photolithography."""
TEXT_AS_READ_FROM_NOTEPAD_NORMALIZED = """     Integrated circuit design, Semiconductor design, chip design or IC design, is a sub-field of Electronics Engineering, encompassing the particular logic and circuit design techniques required to design integrated circuits, or ICs.
\t \n\t      ICs consist of miniaturized electronics components built into an electrical network on a monolithic semiconductor substrate by photolithography."""

# Directories
UI_ELEMENTS_DIR = "ui_elements"  # For source UI images
SCREENSHOTS_DIR = "test_screenshots" # For saved screenshots from tests

# UI Images (source images)
SEARCH_MENU_IMAGE = os.path.join(UI_ELEMENTS_DIR, "search_menu_item.png")
REPLACE_SUBMENU_IMAGE = os.path.join(UI_ELEMENTS_DIR, "replace_submenu_item.png")
FIND_NEXT_BUTTON_IMAGE = os.path.join(UI_ELEMENTS_DIR, "find_next_button.png")
REPLACE_ACTION_BUTTON_IMAGE = os.path.join(UI_ELEMENTS_DIR, "replace_action_button.png")
REPLACE_ALL_BUTTON_IMAGE = os.path.join(UI_ELEMENTS_DIR, "replace_all_button.png")
REPLACE_DIALOG_CLOSE_BUTTON_IMAGE = os.path.join(UI_ELEMENTS_DIR, "replace_dialog_close_button.png")
DONT_SAVE_BUTTON_IMAGE = os.path.join(UI_ELEMENTS_DIR, "dont_save_button.png")

# Settings
INITIAL_APP_WAIT_TIME = 2
ACTION_DELAY = 0.7
UI_IMAGE_CONFIDENCE = 0.9
VALIDATION_IMAGE_CONFIDENCE = 0.85

# Global for Popen process
launched_notepad_process = None

# Ensure directories exist
os.makedirs(UI_ELEMENTS_DIR, exist_ok=True)
os.makedirs(SCREENSHOTS_DIR, exist_ok=True)


# Test scenarios data for Find
TEST_SCENARIOS_FIND = [
    {
        "name": "positive_find_scenario",
        "word_to_find": "chip",
        "validation_image": os.path.join(UI_ELEMENTS_DIR, "find_success_indicator.png"), # Source UI image
        "screenshot_name": os.path.join(SCREENSHOTS_DIR, "notepad_find_positive_test.png"), # Test output screenshot
    },
    {
        "name": "negative_find_scenario",
        "word_to_find": "no exist",
        "validation_image": os.path.join(UI_ELEMENTS_DIR, "find_text_not_found_dialog.png"), # Source UI image
        "screenshot_name": os.path.join(SCREENSHOTS_DIR, "notepad_find_negative_test.png"), # Test output screenshot
    }
]
REPLACE_TEST_SCENARIOS = [
    {
        "name": "positive_replace_once",
        "word_to_find": "chip",
        "replace_with_word": "MICROCHIP",
        "expected_text": TEXT_AS_READ_FROM_NOTEPAD_NORMALIZED.replace("chip", "MICROCHIP", 1),
        "screenshot_name": os.path.join(SCREENSHOTS_DIR, "notepad_replace_positive_once.png"),
    },
    {
        "name": "positive_replace_design_once",
        "word_to_find": "design",
        "replace_with_word": "PLAN",
        "expected_text": TEXT_AS_READ_FROM_NOTEPAD_NORMALIZED.replace("design", "PLAN", 1),
        "screenshot_name": os.path.join(SCREENSHOTS_DIR, "notepad_replace_design_once.png"),
    },
    {
        "name": "negative_replace_nonexistent_word",
        "word_to_find": "nonexistentword",
        "replace_with_word": "ANYTHING",
        "expected_text": TEXT_AS_READ_FROM_NOTEPAD_NORMALIZED,
        "screenshot_name": os.path.join(SCREENSHOTS_DIR, "notepad_replace_nonexistent.png"),
    },
    {
        "name": "positive_replace_with_empty_string",
        "word_to_find": "Semiconductor",
        "replace_with_word": "",
        "expected_text": TEXT_AS_READ_FROM_NOTEPAD_NORMALIZED.replace("Semiconductor", "", 1),
        "screenshot_name": os.path.join(SCREENSHOTS_DIR, "notepad_replace_with_empty.png"),
    }
]

REPLACE_ALL_TEST_SCENARIOS = [
    {
        "name": "positive_replace_all_design",
        "word_to_find": "design",
        "replace_with_word": "LAYOUT",
        "expected_text": re.sub(re.escape("design"), "LAYOUT", TEXT_AS_READ_FROM_NOTEPAD_NORMALIZED, flags=re.IGNORECASE),
        "screenshot_name": os.path.join(SCREENSHOTS_DIR, "notepad_replace_all_design.png"),
    },
    {
        "name": "positive_replace_all_circuit",
        "word_to_find": "circuit",
        "replace_with_word": "NETWORK",
        "expected_text": re.sub(re.escape("circuit"), "NETWORK", TEXT_AS_READ_FROM_NOTEPAD_NORMALIZED, flags=re.IGNORECASE),
        "screenshot_name": os.path.join(SCREENSHOTS_DIR, "notepad_replace_all_circuit.png"),
    },
     {
        "name": "positive_replace_all_semiconductor",
        "word_to_find": "Semiconductor",
        "replace_with_word": "TransistorBased",
        "expected_text": re.sub(re.escape("Semiconductor"), "TransistorBased", TEXT_AS_READ_FROM_NOTEPAD_NORMALIZED, flags=re.IGNORECASE),
        "screenshot_name": os.path.join(SCREENSHOTS_DIR, "notepad_replace_all_semiconductor.png"),
    },
    {
        "name": "negative_replace_all_nonexistent",
        "word_to_find": "wordnotpresent",
        "replace_with_word": "SOMETHING",
        "expected_text": TEXT_AS_READ_FROM_NOTEPAD_NORMALIZED,
        "screenshot_name": os.path.join(SCREENSHOTS_DIR, "notepad_replace_all_nonexistent.png"),
    },
]

def open_and_prepare_notepad():
    """Launch/activate Notepad++, maximize it, and return the window object."""
    global launched_notepad_process
    process_obj_from_popen = None

    try:
        print(f"Attempting to launch/activate Notepad++ at path: {NOTEPAD_PLUS_PLUS_PATH}")
        if not os.path.exists(NOTEPAD_PLUS_PLUS_PATH):
            message = f"Notepad++ executable not found at: {NOTEPAD_PLUS_PLUS_PATH}"
            print(message)
            pytest.fail(message)
            return None
        process_obj_from_popen = subprocess.Popen(NOTEPAD_PLUS_PLUS_PATH)
        if launched_notepad_process is None:
            time.sleep(1)
            if process_obj_from_popen.poll() is None:
                launched_notepad_process = process_obj_from_popen
            else:
                process_obj_from_popen = None
                print("Popen process terminated, Notepad++ was likely already open.")
    except Exception as e:
        print(f"Error when trying to launch Notepad++: {e}.")

    print(f"Waiting {INITIAL_APP_WAIT_TIME} sec. for Notepad++ window to appear/activate...")
    time.sleep(INITIAL_APP_WAIT_TIME)

    npp_windows = pyautogui.getWindowsWithTitle("Notepad++")

    if not npp_windows:
        message = "Notepad++ window not found after launch/activation attempt."
        print(message)
        if process_obj_from_popen and process_obj_from_popen.poll() is None:
            process_obj_from_popen.kill()
        elif launched_notepad_process and launched_notepad_process.poll() is None:
            launched_notepad_process.kill()
        pytest.fail(message)
        return None

    npp_window = npp_windows[0]
    print(f"Found Notepad++ window: {npp_window.title}")

    try:
        if not npp_window.isActive:
            print("Activating Notepad++ window...")
            npp_window.activate()
            time.sleep(ACTION_DELAY / 2)
        if not npp_window.isMaximized:
            print("Maximizing Notepad++ window...")
            npp_window.maximize()
            time.sleep(ACTION_DELAY / 2)
        print("Notepad++ window is ready.")
        return npp_window
    except Exception as e:
        message = f"Error during Notepad++ window activation/maximization: {e}"
        print(message)
        pytest.fail(message)
        return None


@pytest.fixture(scope="module")
def notepad_is_ready():
    """Module-level fixture to launch and close Notepad++."""
    global launched_notepad_process

    print("SETUP (module): Launching and preparing Notepad++...")
    main_npp_window = open_and_prepare_notepad()
    if not main_npp_window:
        pytest.fail("Failed to open and prepare Notepad++ in module setup.")
        return

    yield main_npp_window

    print("\nTEARDOWN (module): Closing Notepad++...")
    try:
        current_npp_windows = pyautogui.getWindowsWithTitle("Notepad++")
        if current_npp_windows:
            target_window = current_npp_windows[0]
            if hasattr(target_window, 'isActive') and not target_window.isActive:
                target_window.activate()
                time.sleep(0.2)

            print("Sending Alt+F4 to close Notepad++...")
            pyautogui.hotkey('alt', 'f4')
            time.sleep(1)

            possible_dont_save_titles = ["Notepad++", "Сохранить файл", "Save file"]
            save_dialog = None
            for title_part in possible_dont_save_titles:
                dialogs = pyautogui.getWindowsWithTitle(title_part)
                if dialogs:
                    for d in dialogs:
                        if d.title != target_window.title:
                            dont_save_button_loc = pyautogui.locateOnScreen(DONT_SAVE_BUTTON_IMAGE, confidence=0.8, # DONT_SAVE_BUTTON_IMAGE is a source UI element
                                                                            grayscale=True,
                                                                            region=(d.left, d.top, d.width, d.height))
                            if dont_save_button_loc:
                                pyautogui.click(dont_save_button_loc)
                                print("Clicked 'Don't Save' button based on image.")
                                save_dialog = d
                                break
                            else:
                                print(f"Save dialog '{d.title}' detected, trying to press 'n' for 'No'.")
                                pyautogui.press('n')
                                save_dialog = d
                                break
                    if save_dialog:
                        break
            if not save_dialog:
                print(
                    "No save dialog detected, or 'Don't Save' button not found by image, attempting 'n' key press as fallback.")
                pyautogui.press('n')
            time.sleep(0.5)

        if launched_notepad_process and launched_notepad_process.poll() is None:
            print("Terminating Notepad++ process...")
            launched_notepad_process.terminate()
            try:
                launched_notepad_process.wait(timeout=3)
            except subprocess.TimeoutExpired:
                launched_notepad_process.kill()
    except Exception as e:
        print(f"Error during TEARDOWN (module): {e}")
        if launched_notepad_process and launched_notepad_process.poll() is None:
            launched_notepad_process.kill()


@pytest.fixture
def new_file_setup_teardown(notepad_is_ready):
    """Function-level fixture to create a new file and close it after the test."""
    if not (notepad_is_ready and hasattr(notepad_is_ready, 'activate')):
        pytest.skip("Notepad++ window object is not valid. Skipping new file setup.")
        return

    try:
        if not notepad_is_ready.isActive:
            print("SETUP (function): Activating Notepad++ window...")
            notepad_is_ready.activate()
            time.sleep(0.3)

        print("SETUP (function): Creating new file (Ctrl+N)...")
        pyautogui.hotkey('ctrl', 'n')
        time.sleep(ACTION_DELAY)
    except Exception as e:
        pytest.fail(f"Failed during new_file_setup_teardown [SETUP]: {e}")
        return

    yield

    print("TEARDOWN (function): Closing current file...")
    try:
        if not (notepad_is_ready and hasattr(notepad_is_ready, 'activate')):
            print("TEARDOWN (function): Notepad++ window object is invalid.")
            return

        if not notepad_is_ready.isActive:
            notepad_is_ready.activate()
            time.sleep(0.3)

        active_window = pyautogui.getActiveWindow()
        if active_window and active_window.title != notepad_is_ready.title:
            print(f"TEARDOWN (function): Closing active dialog: {active_window.title}")
            pyautogui.press('esc')
            time.sleep(0.5)
            active_window = pyautogui.getActiveWindow()
            if active_window and active_window.title != notepad_is_ready.title:
                active_window.close()
                time.sleep(0.5)

        if not notepad_is_ready.isActive:
            notepad_is_ready.activate()
            time.sleep(0.3)

        print("TEARDOWN (function): Closing current file tab (Ctrl+W)...")
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(ACTION_DELAY)
        pyautogui.press('n')  # Don't save
        time.sleep(ACTION_DELAY / 2)
    except Exception as e:
        print(f"ERROR during TEARDOWN (function): {e}")


def get_opencv_args():
    """Determine if OpenCV is available and return appropriate arguments."""
    extra_args = {}
    try:
        import cv2
        extra_args['ui'] = {'confidence': UI_IMAGE_CONFIDENCE, 'grayscale': True}
        extra_args['validation'] = {'confidence': VALIDATION_IMAGE_CONFIDENCE, 'grayscale': True}
        print("INFO: OpenCV found, using 'confidence' and 'grayscale' for image search.")
    except ImportError:
        print("WARN: OpenCV not found. Image matching will be stricter.")
        extra_args['ui'] = {}
        extra_args['validation'] = {}
    return extra_args


def get_image_paths(scenario_type="find", scenario=None):
    """
    Get full paths for required UI element images.
    Source UI images are expected in UI_ELEMENTS_DIR.
    Validation images (if specified in scenario) are also expected to be source UI images.
    """
    paths = {
        'search_menu': SEARCH_MENU_IMAGE,
        'replace_submenu': REPLACE_SUBMENU_IMAGE,
        'find_next_button': FIND_NEXT_BUTTON_IMAGE,
        'replace_action_button': REPLACE_ACTION_BUTTON_IMAGE,
        'replace_all_button': REPLACE_ALL_BUTTON_IMAGE,
        'replace_dialog_close_button': REPLACE_DIALOG_CLOSE_BUTTON_IMAGE
    }

    if scenario_type == "find" and scenario and 'validation_image' in scenario:
        # validation_image path comes directly from scenario, already joined with UI_ELEMENTS_DIR
        paths['validation'] = scenario['validation_image']

    required_keys_for_test = []
    if scenario_type == "find":
        required_keys_for_test = ['search_menu', 'replace_submenu', 'find_next_button']
        if 'validation' in paths:
            required_keys_for_test.append('validation')
    elif scenario_type == "replace":
        required_keys_for_test = ['search_menu', 'replace_submenu', 'find_next_button', 'replace_action_button']
    elif scenario_type == "replace_all":
        required_keys_for_test = ['search_menu', 'replace_submenu', 'replace_all_button']
    elif scenario_type == "close_replace_dialog":
        required_keys_for_test = ['search_menu', 'replace_submenu', 'replace_dialog_close_button']
    else:
        required_keys_for_test = ['search_menu', 'replace_submenu']

    for key in required_keys_for_test:
        image_file_path = paths.get(key)
        if not image_file_path:
             pytest.fail(
                f"Logical error: Image key '{key}' expected in 'paths' for type '{scenario_type}', but is missing.")

        if not os.path.exists(image_file_path): # This checks for source UI images
            error_screenshot_name = os.path.join(SCREENSHOTS_DIR, f"error_source_ui_image_not_found_{os.path.basename(image_file_path)}.png")
            try:
                current_screenshot = pyautogui.screenshot()
                if current_screenshot:
                    current_screenshot.save(error_screenshot_name)
                    print(f"ERROR: Screenshot for missing source UI image saved to: {error_screenshot_name}")
                else:
                    print(f"ERROR: Could not take screenshot for missing source UI image.")
            except Exception as e_scr:
                print(f"ERROR: Could not take screenshot for missing source UI image: {e_scr}")
            pytest.fail(f"Source UI Image '{os.path.basename(image_file_path)}' for type '{scenario_type}' not found at: {image_file_path}")
    return paths


def navigate_to_replace_dialog(image_paths, opencv_args):
    """Navigate to the Replace dialog using menu images."""
    search_menu_location = None
    try:
        search_menu_location = pyautogui.locateCenterOnScreen(
            image_paths['search_menu'], **opencv_args['ui'])
    except Exception as e:
        print(f"Error locating search_menu_item: {e}. Trying without grayscale.")
        temp_args = opencv_args['ui'].copy()
        temp_args.pop('grayscale', None)
        search_menu_location = pyautogui.locateCenterOnScreen(image_paths['search_menu'], **temp_args)

    if not search_menu_location:
        pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, "error_search_menu_not_found.png"))
        pytest.fail("Failed to find 'Search' menu item image")

    pyautogui.click(search_menu_location)
    time.sleep(ACTION_DELAY / 2)

    replace_submenu_location = None
    try:
        replace_submenu_location = pyautogui.locateCenterOnScreen(
            image_paths['replace_submenu'], **opencv_args['ui'])
    except Exception as e:
        print(f"Error locating replace_submenu_item: {e}. Trying without grayscale.")
        temp_args = opencv_args['ui'].copy()
        temp_args.pop('grayscale', None)
        replace_submenu_location = pyautogui.locateCenterOnScreen(image_paths['replace_submenu'], **temp_args)

    if not replace_submenu_location:
        pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, "error_replace_submenu_not_found.png"))
        pytest.fail("Failed to find 'Replace...' submenu item image")

    pyautogui.click(replace_submenu_location)
    time.sleep(ACTION_DELAY)
    active_dialog = pyautogui.getActiveWindow()
    expected_dialog_titles_lower = ["replace", "заменить", "find / replace"]
    if not active_dialog or not any(title_part in active_dialog.title.lower() for title_part in expected_dialog_titles_lower):
        print(
            f"Warning: Replace dialog may not be active. Current active: {active_dialog.title if active_dialog else 'None'}")


def validate_find_result(npp_window, validation_image_path, opencv_args):
    """Validate the result of a Find operation using the validation image."""
    time.sleep(0.5)
    win_left, win_top, win_width, win_height = npp_window.left, npp_window.top, npp_window.width, npp_window.height
    screen_width, screen_height = pyautogui.size()
    safe_left = max(0, win_left);
    safe_top = max(0, win_top)
    safe_width = min(win_left + win_width, screen_width) - safe_left
    safe_height = min(win_top + win_height, screen_height) - safe_top

    if safe_width <= 0 or safe_height <= 0:
        pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, "error_window_region_invalid.png"))
        pytest.fail(f"Invalid window region for validation: L{safe_left} T{safe_top} W{safe_width} H{safe_height}")
    search_region = (safe_left, safe_top, safe_width, safe_height)

    try:
        debug_img = pyautogui.screenshot(region=search_region)
        debug_img.save(os.path.join(SCREENSHOTS_DIR, "debug_validation_search_region.png"))
    except Exception as e:
        print(f"DEBUG: Could not save debug screenshot for validation region: {e}")

    indicator_location = None
    try:
        indicator_location = pyautogui.locateOnScreen(
            validation_image_path, region=search_region, **opencv_args['validation'])
    except Exception as e:
        print(f"Error locating validation_image: {e}. Trying without grayscale.")
        temp_args = opencv_args['validation'].copy();
        temp_args.pop('grayscale', None)
        indicator_location = pyautogui.locateOnScreen(validation_image_path, region=search_region, **temp_args)

    if not indicator_location:
        pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_validation_failed_{os.path.basename(validation_image_path)}.png"))
        pytest.fail(
            f"VALIDATION FAILED: Indicator image '{os.path.basename(validation_image_path)}' not found in region {search_region}")
    print(f"SUCCESS: Validation image '{os.path.basename(validation_image_path)}' found at {indicator_location}")


@pytest.mark.parametrize("scenario", TEST_SCENARIOS_FIND)
def test_notepad_find(notepad_is_ready, new_file_setup_teardown, scenario):
    """Parameterized test for Notepad++ Find functionality."""
    npp_window = notepad_is_ready
    opencv_args = get_opencv_args()
    image_paths = get_image_paths(scenario_type="find", scenario=scenario) # Gets UI_ELEMENTS

    try:
        print("Typing text...")
        pyautogui.write(TEXT_TO_TYPE, interval=0.001)
        time.sleep(ACTION_DELAY)

        print("Opening 'Replace' dialog (used for Find as well)...")
        navigate_to_replace_dialog(image_paths, opencv_args)

        print(f"Typing '{scenario['word_to_find']}' into 'Find what' field...")
        pyautogui.write(scenario['word_to_find'], interval=0.005)
        time.sleep(ACTION_DELAY / 2)

        print("Locating and clicking 'Find Next' button...")
        find_next_button_location = None
        replace_dialog_window = pyautogui.getActiveWindow()
        search_region_dialog = None
        expected_dialog_titles = ["replace", "find", "заменить", "найти", "find / replace"]
        if replace_dialog_window and any(
                title in replace_dialog_window.title.lower() for title in expected_dialog_titles):
            search_region_dialog = (replace_dialog_window.left, replace_dialog_window.top, replace_dialog_window.width,
                                    replace_dialog_window.height)
            print(
                f"Searching for Find Next button within dialog: {replace_dialog_window.title} region: {search_region_dialog}")
        else:
            print(
                f"Warning: Could not determine specific dialog window for Find Next. Active: {replace_dialog_window.title if replace_dialog_window else 'None'}. Searching whole screen.")

        try:
            find_next_button_location = pyautogui.locateCenterOnScreen(
                image_paths['find_next_button'], region=search_region_dialog, **opencv_args['ui'])
        except Exception as e:
            print(f"Error locating find_next_button: {e}. Trying without grayscale.")
            temp_args = opencv_args['ui'].copy();
            temp_args.pop('grayscale', None)
            find_next_button_location = pyautogui.locateCenterOnScreen(
                image_paths['find_next_button'], region=search_region_dialog, **temp_args)

        if not find_next_button_location:
            pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_find_next_button_not_found_{scenario['name']}.png"))
            pytest.fail("Failed to find 'Find Next' button image.")

        pyautogui.click(find_next_button_location)
        time.sleep(1.0)

        print(f"Validating result using '{os.path.basename(scenario['validation_image'])}'...")
        validate_find_result(npp_window, image_paths['validation'], opencv_args) # validation_image is a source UI image

        pyautogui.screenshot(scenario['screenshot_name']) # Saves to SCREENSHOTS_DIR via scenario dict
        assert os.path.exists(scenario['screenshot_name']), f"Screenshot was not created: {scenario['screenshot_name']}"

        print("Closing dialog (ESC)...")
        pyautogui.press('esc')
        time.sleep(ACTION_DELAY / 2)

    except pyautogui.FailSafeException:
        pytest.fail("PyAutoGUI fail-safe triggered (mouse moved to a corner)")
    except Exception as e:
        pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_test_find_{scenario['name']}.png"))
        print(f"Error during test '{scenario['name']}': {e}")
        raise


@pytest.mark.parametrize("scenario", REPLACE_TEST_SCENARIOS)
def test_notepad_replace(notepad_is_ready, new_file_setup_teardown, scenario):
    """Parameterized test for Notepad++ Replace functionality (single replace)."""
    npp_window = notepad_is_ready
    opencv_args = get_opencv_args()
    image_paths = get_image_paths(scenario_type="replace", scenario=scenario)

    try:
        if not npp_window.isActive:
            npp_window.activate()
            time.sleep(0.5)

        print("Typing initial text for replace test...")
        pyautogui.write(TEXT_TO_TYPE, interval=0.005)
        time.sleep(ACTION_DELAY / 2)

        print("Moving cursor to the beginning of the document (Ctrl+Home)...")
        pyautogui.hotkey('ctrl', 'home')
        time.sleep(0.3)

        print("Opening 'Replace' dialog...")
        navigate_to_replace_dialog(image_paths, opencv_args)

        print(f"Typing '{scenario['word_to_find']}' into 'Find what' field...")
        pyautogui.write(scenario['word_to_find'], interval=0.04)
        time.sleep(0.4)

        pyautogui.press('tab')
        time.sleep(0.4)

        print("Clearing 'Replace with' field (Ctrl+A, Del)...")
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.press('delete')
        time.sleep(0.2)

        print(f"Typing '{scenario['replace_with_word']}' into 'Replace with' field...")
        pyautogui.write(scenario['replace_with_word'], interval=0.04)
        time.sleep(0.5)

        replace_dialog_window = pyautogui.getActiveWindow()
        search_region_dialog_replace = None
        expected_dialog_titles = ["replace", "заменить", "find / replace"]
        if replace_dialog_window and any(
                title in replace_dialog_window.title.lower() for title in expected_dialog_titles):
            search_region_dialog_replace = (
            replace_dialog_window.left, replace_dialog_window.top, replace_dialog_window.width,
            replace_dialog_window.height)
        else:
            print(
                f"Warning: Could not reliably determine Replace dialog window. Active: {replace_dialog_window.title if replace_dialog_window else 'None'}. Using npp_window region as fallback.")
            search_region_dialog_replace = (npp_window.left, npp_window.top, npp_window.width, npp_window.height)

        print("Locating and clicking 'Find Next' button in Replace dialog...")
        find_next_button_location_in_replace_dialog = None
        try:
            find_next_button_location_in_replace_dialog = pyautogui.locateCenterOnScreen(
                image_paths['find_next_button'],
                region=search_region_dialog_replace,
                **opencv_args['ui']
            )
        except Exception as e:
            print(f"Error locating 'Find Next' (in Replace dialog): {e}. Trying without grayscale.")
            temp_args_fn = opencv_args['ui'].copy()
            temp_args_fn.pop('grayscale', None)
            find_next_button_location_in_replace_dialog = pyautogui.locateCenterOnScreen(
                image_paths['find_next_button'], region=search_region_dialog_replace, **temp_args_fn)

        if not find_next_button_location_in_replace_dialog:
            pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_find_next_in_replace_dialog_not_found_{scenario['name']}.png"))
            if scenario['expected_text'] != TEXT_AS_READ_FROM_NOTEPAD_NORMALIZED:
                pytest.fail(
                    f"Failed to find 'Find Next' button in Replace dialog for positive scenario '{scenario['name']}'.")
            else:
                print(
                    f"INFO: 'Find Next' button not found for negative scenario '{scenario['name']}' (word might not exist). Proceeding to check text.")

        if find_next_button_location_in_replace_dialog:
            print(f"Clicking 'Find Next' button in Replace dialog at {find_next_button_location_in_replace_dialog}")
            pyautogui.click(find_next_button_location_in_replace_dialog)
            time.sleep(0.8)

            debug_dialog_screenshot_name_after_find = os.path.join(SCREENSHOTS_DIR, f"debug_replace_dialog_after_find_next_{scenario['name']}.png")
            if search_region_dialog_replace:
                 pyautogui.screenshot(debug_dialog_screenshot_name_after_find, region=search_region_dialog_replace)
                 print(f"DEBUG: Screenshot of replace dialog after 'Find Next' saved to {debug_dialog_screenshot_name_after_find}")
            else:
                 pyautogui.screenshot(debug_dialog_screenshot_name_after_find)
                 print(f"DEBUG: Screenshot of full screen (replace dialog after 'Find Next') saved to {debug_dialog_screenshot_name_after_find}")


            print("Locating and clicking 'Replace' (action) button...")
            replace_button_location = None
            try:
                replace_button_location = pyautogui.locateCenterOnScreen(
                    image_paths['replace_action_button'], region=search_region_dialog_replace, **opencv_args['ui'])
            except Exception as e:
                print(f"Error locating 'Replace' action button: {e}. Trying without grayscale.")
                temp_args_rep = opencv_args['ui'].copy()
                temp_args_rep.pop('grayscale', None)
                replace_button_location = pyautogui.locateCenterOnScreen(
                    image_paths['replace_action_button'], region=search_region_dialog_replace, **temp_args_rep)

            if not replace_button_location:
                pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_replace_action_button_not_found_{scenario['name']}.png"))
                pytest.fail(f"Failed to find 'Replace' (action) button image for scenario '{scenario['name']}'.")

            print(f"Clicking 'Replace' action button at {replace_button_location}")
            pyautogui.click(replace_button_location)
            time.sleep(1.5)

        current_active_dialog = pyautogui.getActiveWindow()
        if current_active_dialog and any(
                title in current_active_dialog.title.lower() for title in expected_dialog_titles):
            print("Closing 'Replace' dialog (ESC) after operations or if 'Find Next' was skipped...")
            pyautogui.press('esc')
            time.sleep(ACTION_DELAY / 2)
        elif not find_next_button_location_in_replace_dialog:
            print("Closing 'Replace' dialog (ESC) because 'Find Next' was not performed...")
            pyautogui.press('esc')
            time.sleep(ACTION_DELAY / 2)

        print("Validating text in Notepad++ editor...")
        if not npp_window.isActive:
            npp_window.activate()
            time.sleep(0.3)

        try:
            pyperclip.copy('')
        except pyperclip.PyperclipException as e:
            print(f"Note: pyperclip.copy('') failed: {e}")

        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)

        try:
            retrieved_text = pyperclip.paste()
        except pyperclip.PyperclipException as e:
            pytest.fail(f"Failed to paste text from clipboard: {e}.")

        retrieved_text_normalized = retrieved_text.replace('\r\n', '\n')
        expected_text_normalized = scenario['expected_text']
        if '\r\n' in expected_text_normalized:
            expected_text_normalized = expected_text_normalized.replace('\r\n', '\n')

        assert retrieved_text_normalized == expected_text_normalized, \
            f"Text after replace does not match expected for scenario '{scenario['name']}'. \nExpected:\n{expected_text_normalized}\nGot:\n{retrieved_text_normalized}"

        print(f"SUCCESS: Text validation passed for scenario '{scenario['name']}'.")

        pyautogui.screenshot(scenario['screenshot_name']) # Saves to SCREENSHOTS_DIR
        assert os.path.exists(scenario['screenshot_name']), f"Screenshot was not created: {scenario['screenshot_name']}"

    except pyautogui.FailSafeException:
        pytest.fail("PyAutoGUI fail-safe triggered")
    except Exception as e:
        pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_test_replace_{scenario['name']}.png"))
        print(f"Error during replace test '{scenario['name']}': {e}")
        raise


@pytest.mark.parametrize("scenario", REPLACE_ALL_TEST_SCENARIOS)
def test_notepad_replace_all(notepad_is_ready, new_file_setup_teardown, scenario):
    """Parameterized test for Notepad++ Replace All functionality."""
    npp_window = notepad_is_ready
    opencv_args = get_opencv_args()
    image_paths = get_image_paths(scenario_type="replace_all", scenario=scenario)

    try:
        if not npp_window.isActive:
            npp_window.activate()
            time.sleep(0.5)

        print("Typing initial text for replace_all test...")
        pyautogui.write(TEXT_TO_TYPE, interval=0.005)
        time.sleep(ACTION_DELAY / 2)

        print("Moving cursor to the beginning of the document (Ctrl+Home)...")
        pyautogui.hotkey('ctrl', 'home')
        time.sleep(0.3)

        print("Opening 'Replace' dialog...")
        navigate_to_replace_dialog(image_paths, opencv_args)

        print(f"Typing '{scenario['word_to_find']}' into 'Find what' field...")
        pyautogui.write(scenario['word_to_find'], interval=0.04)
        time.sleep(0.4)

        pyautogui.press('tab')
        time.sleep(0.4)

        print("Clearing 'Replace with' field (Ctrl+A, Del)...")
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.press('delete')
        time.sleep(0.2)

        print(f"Typing '{scenario['replace_with_word']}' into 'Replace with' field...")
        pyautogui.write(scenario['replace_with_word'], interval=0.04)
        time.sleep(0.5)

        replace_dialog_window = pyautogui.getActiveWindow()
        search_region_dialog_replace = None
        expected_dialog_titles = ["replace", "заменить", "find / replace"]
        if replace_dialog_window and any(
                title in replace_dialog_window.title.lower() for title in expected_dialog_titles):
            search_region_dialog_replace = (
            replace_dialog_window.left, replace_dialog_window.top, replace_dialog_window.width,
            replace_dialog_window.height)
        else:
            print(f"Warning: Could not reliably determine Replace dialog window. Using npp_window region.")
            search_region_dialog_replace = (npp_window.left, npp_window.top, npp_window.width, npp_window.height)

        print("Locating and clicking 'Replace All' button...")
        replace_all_button_location = None
        try:
            replace_all_button_location = pyautogui.locateCenterOnScreen(
                image_paths['replace_all_button'],
                region=search_region_dialog_replace,
                **opencv_args['ui']
            )
        except Exception as e:
            print(f"Error locating 'Replace All' button: {e}. Trying without grayscale.")
            temp_args_ra = opencv_args['ui'].copy()
            temp_args_ra.pop('grayscale', None)
            replace_all_button_location = pyautogui.locateCenterOnScreen(
                image_paths['replace_all_button'], region=search_region_dialog_replace, **temp_args_ra)

        if not replace_all_button_location:
            pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_replace_all_button_not_found_{scenario['name']}.png"))
            pytest.fail(f"Failed to find 'Replace All' button image for scenario '{scenario['name']}'.")

        print(f"Clicking 'Replace All' button at {replace_all_button_location}")
        pyautogui.click(replace_all_button_location)
        time.sleep(1.5)

        print("Attempting to close potential 'Replace All' confirmation dialog (pressing Enter)...")
        pyautogui.press('enter')
        time.sleep(0.5)

        current_active_dialog = pyautogui.getActiveWindow()
        if current_active_dialog and any(
                title in current_active_dialog.title.lower() for title in expected_dialog_titles):
            print("Closing 'Replace' dialog (ESC) after 'Replace All' operation...")
            pyautogui.press('esc')
            time.sleep(ACTION_DELAY / 2)
        else:
            print("INFO: 'Replace' dialog seems already closed or not active after 'Replace All'.")

        print("Validating text in Notepad++ editor after Replace All...")
        if not npp_window.isActive:
            npp_window.activate()
            time.sleep(0.3)

        try:
            pyperclip.copy('')
        except pyperclip.PyperclipException as e:
            print(f"Note: pyperclip.copy('') failed: {e}")

        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)

        try:
            retrieved_text = pyperclip.paste()
        except pyperclip.PyperclipException as e:
            pytest.fail(f"Failed to paste text from clipboard: {e}.")

        retrieved_text_normalized = retrieved_text.replace('\r\n', '\n')
        expected_text_normalized = scenario['expected_text']
        if '\r\n' in expected_text_normalized:
            expected_text_normalized = expected_text_normalized.replace('\r\n', '\n')

        assert retrieved_text_normalized == expected_text_normalized, \
            f"Text after replace all does not match expected for scenario '{scenario['name']}'. \nExpected:\n{expected_text_normalized}\nGot:\n{retrieved_text_normalized}"

        print(f"SUCCESS: Text validation passed for Replace All scenario '{scenario['name']}'.")

        pyautogui.screenshot(scenario['screenshot_name']) # Saves to SCREENSHOTS_DIR
        assert os.path.exists(scenario['screenshot_name']), f"Screenshot was not created: {scenario['screenshot_name']}"

    except pyautogui.FailSafeException:
        pytest.fail("PyAutoGUI fail-safe triggered")
    except Exception as e:
        pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_test_replace_all_{scenario['name']}.png"))
        print(f"Error during replace_all test '{scenario['name']}': {e}")
        raise


def test_notepad_replace_dialog_close_button(notepad_is_ready, new_file_setup_teardown):
    """Tests closing the Replace dialog using its Close/Cancel button after typing standard text."""
    npp_window = notepad_is_ready
    opencv_args = get_opencv_args()
    image_paths = get_image_paths(scenario_type="close_replace_dialog") # Gets UI_ELEMENTS
    test_name = "replace_dialog_close_test"
    screenshot_success_name = os.path.join(SCREENSHOTS_DIR, f"{test_name}_success.png")


    try:
        if not npp_window.isActive:
            npp_window.activate()
            time.sleep(0.5)

        print(f"Typing standard text ({len(TEXT_TO_TYPE)} chars) for close dialog test...")
        pyautogui.write(TEXT_TO_TYPE, interval=0.005)
        time.sleep(ACTION_DELAY / 2)
        print("Moving cursor to the beginning (Ctrl+Home)...")
        pyautogui.hotkey('ctrl', 'home')
        time.sleep(0.3)

        print("Opening 'Replace' dialog...")
        navigate_to_replace_dialog(image_paths, opencv_args)

        replace_dialog_window = pyautogui.getActiveWindow()
        search_region_dialog_close = None
        possible_dialog_titles = ["Replace", "Заменить", "Find / Replace"]

        if replace_dialog_window and any(title.lower() in replace_dialog_window.title.lower() for title in possible_dialog_titles):
             search_region_dialog_close = (
                 replace_dialog_window.left, replace_dialog_window.top,
                 replace_dialog_window.width, replace_dialog_window.height
             )
             print(f"Searching for Close button within dialog: {replace_dialog_window.title} region: {search_region_dialog_close}")
        else:
             pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_{test_name}_dialog_not_active_before_close.png"))
             pytest.fail(f"Replace dialog window not found or not active before attempting to close. Active: {replace_dialog_window.title if replace_dialog_window else 'None'}")

        print("Locating and clicking 'Close' button in Replace dialog...")
        close_button_location = None
        try:
            close_button_location = pyautogui.locateCenterOnScreen(
                image_paths['replace_dialog_close_button'], # This is a source UI_ELEMENT image
                region=search_region_dialog_close,
                **opencv_args['ui']
            )
        except Exception as e:
            print(f"Error locating 'Close' button: {e}. Trying without grayscale.")
            temp_args_close = opencv_args['ui'].copy()
            temp_args_close.pop('grayscale', None)
            close_button_location = pyautogui.locateCenterOnScreen(
                image_paths['replace_dialog_close_button'], region=search_region_dialog_close, **temp_args_close)

        if not close_button_location:
            pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_{test_name}_close_button_not_found.png"))
            pytest.fail(f"Failed to find 'Close' button image in Replace dialog.")

        print(f"Clicking 'Close' button at {close_button_location}")
        pyautogui.click(close_button_location)
        time.sleep(1.0)

        print("Validating Replace dialog is closed...")
        dialog_found = False
        for title_part in possible_dialog_titles:
            open_dialogs = pyautogui.getWindowsWithTitle(title_part)
            if open_dialogs:
                for win in open_dialogs:
                     if any(expected_title.lower() in win.title.lower() for expected_title in possible_dialog_titles):
                          dialog_found = True
                          print(f"ERROR: Dialog window with title '{win.title}' still found after clicking Close.")
                          pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_{test_name}_dialog_not_closed.png"))
                          break
            if dialog_found:
                break

        assert not dialog_found, "The Replace dialog window was still found after the Close button was clicked."

        npp_window_title = notepad_is_ready.title
        active_window = pyautogui.getActiveWindow()
        assert active_window and active_window.title == npp_window_title, \
            f"Main Notepad++ window ('{npp_window_title}') is not active after closing Replace dialog. Active: {active_window.title if active_window else 'None'}"

        print(f"SUCCESS: Test '{test_name}' passed. Replace dialog closed successfully.")
        pyautogui.screenshot(screenshot_success_name) # Saves to SCREENSHOTS_DIR
        assert os.path.exists(screenshot_success_name), f"Screenshot was not created: {screenshot_success_name}"


    except pyautogui.FailSafeException:
        pytest.fail("PyAutoGUI fail-safe triggered")
    except Exception as e:
        pyautogui.screenshot(os.path.join(SCREENSHOTS_DIR, f"error_test_dialog_close_{test_name}_{type(e).__name__}.png"))
        print(f"Error during test '{test_name}': {e}")
        raise