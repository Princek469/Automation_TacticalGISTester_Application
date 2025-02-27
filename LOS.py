import pyautogui as auto
import subprocess
import xml.etree.ElementTree as ET
import time
import pyperclip
import pygetwindow as gw    
import os
import csv

def read_app_path_from_xml(xml_file_path):
    """Read the application path from the XML file."""
    tree = ET.parse(xml_file_path)
    root = tree.getroot()
    app_path_element = root.find('AppPath')
    if app_path_element is not None:
        return app_path_element.text
    else:
        raise ValueError("AppPath not found in XML.")

def read_workspace_file_from_xml(xml_file_path):
    """Read the workspace file path from the XML file."""
    tree = ET.parse(xml_file_path)
    root = tree.getroot()
    workspace_file_path_element = root.find('WorkspaceFilePath')
    if workspace_file_path_element is not None:
        return workspace_file_path_element.text
    else:
        raise ValueError("WorkspaceFilePath not found in XML.")

def read_csv_file_path_from_xml(xml_file_path, input_path_name):
    """Read the CSV file path from the XML file."""
    tree = ET.parse(xml_file_path)
    root = tree.getroot()
    
    # Determine the correct path based on the input_path_name
    if input_path_name == 'InputCSVFilePathLOS':
        csv_file_path_element = root.find('LOS/InputCSVFilePath')
    elif input_path_name == 'OutputCSVFilePathLOS':
        csv_file_path_element = root.find('LOS/OutputCSVFilePath')
    elif input_path_name == 'InputCSVFilePathCC':
        csv_file_path_element = root.find('CoordinateConversion/InputCSVFilePath')
    elif input_path_name == 'OutputCSVFilePathCC':
        csv_file_path_element = root.find('CoordinateConversion/OutputCSVFilePath')
    else:
        raise ValueError(f"Unknown input path name: {input_path_name}")
    
    if csv_file_path_element is not None:
        return csv_file_path_element.text
    else:
        raise ValueError(f"{input_path_name} not found in XML.")

def press_tabs(press):
    auto.press('tab', interval=0.2, presses=press)
    auto.press('enter')
    time.sleep(1)

def press_tabs_value(press, value):
    auto.press('tab', interval=0.2, presses=press)
    auto.press('backspace')
    auto.typewrite(value, interval=0.2)
    time.sleep(1) 

def open_application(app_path):
    """Open the application and wait for it to be ready."""
    subprocess.Popen(app_path)
    time.sleep(5)

def activate_window(window_title):
    """Activate and maximize the application window."""
    windows = gw.getWindowsWithTitle(window_title)
    if windows:
        app_window = windows[0]
        app_window.activate()  
        app_window.maximize() 
        print("Application window activated and maximized.")
    else:
        print("Application window not found!")
        return False
    time.sleep(1)  
    return True

def click_file_menu():
    """Click on the 'File' menu."""
    auto.press('alt')  
    time.sleep(0.5)
    auto.press('f') 
    print("Clicked on 'File' menu.")
    time.sleep(1)

def select_workspace_open_option():
    """Select the 'Workspace' and then 'Open' option from the File menu."""
    for _ in range(5):
        auto.press('down')
        time.sleep(0.5)
    auto.press('enter') 
    time.sleep(1)

    for _ in range(1):  
        auto.press('down')
        time.sleep(0.5)
    auto.press('enter') 
    print("Selected 'Open' option.")
    time.sleep(1) 

def select_workspace_file(file_path):
    """Type the workspace file path and select the file."""
    if os.path.exists(file_path):
        auto.typewrite(file_path, interval=0.1)
        time.sleep(1)  
        auto.press('enter') 
        print("Workspace file selected successfully.")
        time.sleep(2)
    else:
        print("Workspace file does not exist.")

def add_vector_layer(file_menu_position, add_button_position):
    """Add a vector layer to the workspace."""
    auto.moveTo(file_menu_position[0], file_menu_position[1], duration=0.5)
    print("Clicked on 'File' menu.")
    time.sleep(1)  

    press_tabs(3)

    for _ in range(4):
        auto.press('down')
        time.sleep(0.5)  
    auto.press('enter')
    time.sleep(1)

    auto.moveTo(add_button_position[0], add_button_position[1], duration=0.5)
    auto.click()
    print("Clicked on 'Add' button.")
    time.sleep(2)  # Wait for the layer to be added 

def select_item_and_apply():
    click_file_menu()
    for _ in range(5):
        auto.press('right')
        time.sleep(0.5)
    
    for _ in range(1):
        auto.press('down')
        time.sleep(0.5)
    auto.press('enter')

    auto.press('tab', interval=0.5, presses=4)

    for _ in range(11): 
        auto.press('down')  
        time.sleep(0.5) 
    auto.press('enter')  
    print("Selected 'LOS' option.")
    time.sleep(2)

def read_user_defined_input_from_csv(csv_file_path):
    """Read user-defined input values from a CSV file and return a list of tuples."""
    user_defined_list = []
    with open(csv_file_path, mode='r') as file:
        reader = csv.reader(file)
        next(reader) 
        for row in reader:
            if len(row) >= 4: 
                east, north, height, user_defined_value = row[0].strip(), row[1].strip(), row[2].strip(), row[3].strip()
                user_defined_list.append((east, north, height, user_defined_value))
    return user_defined_list

def read_coordinates_from_csv(csv_file_path):
    """Read coordinates from a CSV file and return a list of tuples."""
    coordinates_list = []
    with open(csv_file_path, mode='r') as file:
        reader = csv.reader(file)
        next(reader)  
        for row in reader:
            if len(row) >= 3: 
                east, north, height = row[0].strip(), row[1].strip(), row[2].strip()
                coordinates_list.append((east, north, height))
    return coordinates_list 

def input_user_defined_values(user_defined_value):
    """Input user-defined values into the specified fields."""
    press_tabs_value(2, user_defined_value)

def read_output_csv_file_path_from_xml(xml_file_path, output_path_name):
    """Read the output CSV file path from the XML file."""
    tree = ET.parse(xml_file_path)
    root = tree.getroot()
    
    # Determine the correct path based on the output_path_name
    if output_path_name == 'OutputCSVFilePathLOS':
        output_csv_file_path_element = root.find('LOS/OutputCSVFilePath')
    elif output_path_name == 'OutputCSVFilePathCC':
        output_csv_file_path_element = root.find('CoordinateConversion/OutputCSVFilePath')
    else:
        raise ValueError(f"Unknown output path name: {output_path_name}")
    
    if output_csv_file_path_element is not None:
        return output_csv_file_path_element.text
    else:
        raise ValueError(f"{output_path_name} not found in XML.")

def write_output_to_csv(user_defined_input):
    """Write the output values to a CSV file after copying from the specified input field."""
    auto.moveTo(user_defined_input[0], user_defined_input[1], duration=1) 
    auto.click()  
    time.sleep(1)  
    auto.hotkey('ctrl', 'a') 
    time.sleep(0.5)  
    auto.hotkey('ctrl', 'c')
    time.sleep(1) 
    # Get the copied values from the clipboard
    output_text = pyperclip.paste()  
    return output_text.strip() if output_text else None

def click_arrow_button(arrow_button_position):
    auto.moveTo(arrow_button_position[0], arrow_button_position[1], duration=0.5)
    auto.click()
    print("Clicked on the arrow button.")
    time.sleep(1)  

def drag_and_select(start_position, end_position):
    """Drag from start_position to end_position."""
    auto.moveTo(start_position[0], start_position[1], duration=1)
    auto.mouseDown() 
    print("Mouse down at starting position:", start_position)
    time.sleep(1)

    auto.moveTo(end_position[0], end_position[1], duration=1)
    print("Dragging to ending position:", end_position)
    time.sleep(1)  

    auto.mouseUp()  
    print("Mouse up at ending position:", end_position)
    time.sleep(1) 

def select_properties_option():
    """Click the 'Item' button and select ' Properties' option."""
    click_file_menu()

    for _ in range(5):  
        auto.press('right')
        time.sleep(0.5)
   
    for _ in range(6): 
        auto.press('down')
        time.sleep(0.5) 

    auto.press('enter')  
    print("Selected 'Properties' option.")
    time.sleep(2)  

def validate_output(input_user_defined_value, output_user_defined_value):
    """Validate output values against user input."""
    return (input_user_defined_value == output_user_defined_value)

def write_validation_to_csv(output_csv_file_path, user_defined_value, validation_result):
    """Write validation result to the output CSV file."""
    with open(output_csv_file_path, mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([user_defined_value, validation_result])

def read_output_value_from_csv(output_csv_file_path):
    """Read the last written output value from the output CSV file."""
    with open(output_csv_file_path, mode='r') as file:
        reader = csv.reader(file)
        last_row = list(reader)[-1]  
        return last_row[0] if last_row else None

def open_workspace(xml_file_path, window_title, file_menu_position, add_button_position, user_defined_input):
    """Open the application and load the workspace file."""
    try:
        print("Opening application...")
        app_path = read_app_path_from_xml(xml_file_path)
        open_application(app_path)
        
        if not activate_window(window_title):
            print("Failed to activate window.")
            return
        
        print("Clicking file menu...")
        click_file_menu()
        select_workspace_open_option()
        workspace_file_path = read_workspace_file_from_xml(xml_file_path)
        select_workspace_file(workspace_file_path)

        print("Adding vector layer...")
        add_vector_layer(file_menu_position, add_button_position)

        print("Selecting item and applying...")
        select_item_and_apply()

        print("Reading user-defined input from CSV...")
        csv_file_path = read_csv_file_path_from_xml(xml_file_path, 'InputCSVFilePathLOS')
        user_defined_list = read_user_defined_input_from_csv(csv_file_path)
        coordinates_list = read_coordinates_from_csv(csv_file_path)
        click_arrow_button(arrow_button_position) 
        drag_and_select(start_position, end_position)    
        select_properties_option()

        output_csv_file_path = read_output_csv_file_path_from_xml(xml_file_path, 'OutputCSVFilePathLOS')
        # Clear the output CSV file before writing results
        with open(output_csv_file_path, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['User _Defined', 'validation'])  # Write header

        for index, (east, north, height, user_defined_value) in enumerate(user_defined_list):
            print(f"Processing index {index}: {user_defined_value}")

            input_user_defined_values(user_defined_value)
            # Click the 'Vertices' button
            press_tabs(14)
            # click get button
            press_tabs(23)

            east, north, height = coordinates_list[index]  

            press_tabs_value(43, east) 
            press_tabs_value(1, north)
            press_tabs_value(1, height)

            # Click the update button
            press_tabs(18)
            # Click 'OK' in the coordinates dialog
            press_tabs(52)
            # Copy the user-defined value (if needed)
            output_user_defined_value = write_output_to_csv(user_defined_input) 
            # Click 'OK' in the vertices dialog
            press_tabs(48)
            # Validate the output
            is_valid = validate_output(user_defined_value, output_user_defined_value)
            # Write validation results to the output CSV file
            write_validation_to_csv(output_csv_file_path, user_defined_value, is_valid)
            
            if index == 0:
                select_properties_option()  
            elif index < len(user_defined_list) - 1: 
                select_properties_option()
                       
        print("Finished processing all user-defined values.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    auto.FAILSAFE = False
    window_title = "TacticalGis Test Application"
    file_menu_position = (18, 39)  
    add_button_position = (435, 96)    
    arrow_button_position = (284, 62) 
    start_position = (9, 156)  
    end_position = (1818, 1009)
    user_defined_input = (716, 250)
    close_tactical_tester = (1983, 9)
    xml_file_path = 'C:\\Users\\ec.pkjha\\Desktop\\AutoGUI\\AutoGUIFiles.xml'
    try:
        open_workspace(xml_file_path, window_title, file_menu_position, add_button_position, user_defined_input)
        time.sleep(2)
        auto.moveTo(close_tactical_tester[0], close_tactical_tester[1], duration=0.2)
        auto.click()
    except Exception as e:
         print(f"An error occurred: {e}")







































