import pyautogui as auto
from LOS import open_application, read_app_path_from_xml, read_csv_file_path_from_xml, read_output_csv_file_path_from_xml,  read_workspace_file_from_xml, activate_window, click_file_menu, select_workspace_open_option, select_workspace_file
import pyperclip
import time
import csv
import os

projection_mapping = {
    "Geolagacy": (2, 'down'),  
    "56K_LCC": (0, ''),        
    "Geo": (1, 'down'),       
    "igrs": (7, 'down'),       
    "Webmercator": (6, 'down'),
    "India": (4, 'down'),
    "IGRS_W...4Datum": (3, 'down'), 
    "UTM": (5, 'down')
}

def click_file_menu():
    """Click on the 'File' menu."""
    auto.press('alt')  
    time.sleep(0.5)
    print("Clicked on 'File' menu.")
    time.sleep(1)



def click_tools_button():
    """Click on the Tools button and then the Coordinate Calculator button."""
    click_file_menu()

    for _ in range(6):
        auto.press('right')
        time.sleep(0.5)

    for _ in range(3):
        auto.press('down')
        time.sleep(0.5)
    auto.press('enter')
    time.sleep(2)

def read_coordinates_from_csv(csv_file_path):
    """Read coordinates and projection values from a CSV file and return as lists."""
    coordinates = []

    with open(csv_file_path, mode='r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            try:
                coordinates.append({
                   'input_projection': row['input_projection'].strip(),
                    'output_projection': row['output_projection'].strip(),
                    'zone': row['zone'].strip(),
                    'lat': row['lat'].strip(),
                    'long': row['long'].strip(),
                    'tolerance' : row['tolerance'].strip(),
                    'East': row['East'].strip(),
                    'North': row['North'].strip()
                })
            except KeyError as e:
                print(f"Missing key in row: {row}. Error: {e}")
            except ValueError as e:
                print(f"Invalid value in row: {row}. Error: {e}")

    return coordinates

def input_coordinate(press, value):
    """Input a coordinate value into the specified field."""
    auto.press('tab', interval=0.1, presses=press)
    auto.press('backspace')
    auto.typewrite(value, interval=0.1)
    time.sleep(1) 
    
def move_down_in_menu(press_count):
    """Simulate pressing the down key a specified number of times."""
    for _ in range(press_count):
        auto.press('down')
        time.sleep(0.5)

def select_input_projection(input_projection, tab_presses):
    """Select the input projection based on the provided projection name."""
    print("SELECT INPUT PROJECTION")
    auto.press('tab', interval=0.2, presses=tab_presses)
    time.sleep(1)  
    auto.press('home')
    if input_projection in projection_mapping:
        presses, direction = projection_mapping[input_projection]
        if direction == 'down':
            print(f"Moving down {presses} times to select input projection: {input_projection}")
            move_down_in_menu(presses)
        elif presses == 0:
            print(f"Selecting {input_projection} directly.")
    else:
        print(f"Input projection '{input_projection}' not found in available projections.")
        return 

    time.sleep(1) 
    print(f"Input projection '{input_projection}' selected successfully.")

def select_output_projection(input_projection, output_projection):
    """Select the output projection based on the provided projection name."""
    print("SELECT OUTPUT PROJECTION")
    tab_press = 6 if input_projection == "Geolagacy" else 5  # Adjust based on input projection
    auto.press('tab', interval=0.2, presses=tab_press)
    time.sleep(0.2)

    auto.press('home')
    if output_projection in projection_mapping:
        presses, direction = projection_mapping[output_projection]
        if direction == 'down':
            print(f"Moving down {presses} times to select output projection: {output_projection}")
            move_down_in_menu(presses)
        elif presses == 0:
            print(f"Selecting {output_projection} directly.")
    else:
        print(f"Output projection '{output_projection}' not found in available projections.")
        return 

    time.sleep(1)  
    print(f"Output projection '{output_projection}' selected successfully.")


def press_tabs_copy_output_values(press):
    """Press Tab to navigate and copy the selected text."""
    auto.press('tab', interval=0.2, presses=press)       
    time.sleep(0.5) 
    auto.hotkey('ctrl', 'c')  
    time.sleep(1)

def copy_values_with_tabs(tabs):
    """Copy values from the output fields based on the number of tabs to press."""
    output_values = []
    
    press_tabs_copy_output_values(tabs)
    first_value = pyperclip.paste().strip()
    output_values.append(first_value)
    time.sleep(1)

    press_tabs_copy_output_values(1)
    second_value = pyperclip.paste().strip()
    output_values.append(second_value)
    time.sleep(1)

    press_tabs_copy_output_values(1)
    third_value = pyperclip.paste().strip()
    output_values.append(third_value)

    return output_values

def copy_output_values(input_projection, output_projection):
    """Copy the output values from the specified output field positions."""
    if input_projection == "Geolagacy":
        return copy_values_with_tabs(17) 
    elif output_projection == "Geolagacy":
        return copy_values_with_tabs(16)  
    else:
        return copy_values_with_tabs(16)

def write_output_to_csv(output_csv_file_path, zone, lat, long, validation):
    """Write the output values to a new CSV file in the specified format."""
    file_exists = os.path.isfile(output_csv_file_path)

    with open(output_csv_file_path, mode='a', newline='') as file:
        writer = csv.writer(file)

        if not file_exists:
            writer.writerow(['zone', 'lat', 'long', 'validation'])
        writer.writerow([zone, lat, long, validation])

def validate_output(output_lat, output_long, user_lat, user_long):
    """Validate output values against user input."""
    return (output_lat == user_lat) and (output_long == user_long)

def coordinate_value(coordinate, input_projection, output_projection):
    zone, latitude, longitude = coordinate
    
    if input_projection == "Geolagacy":
        input_coordinate(14, zone)  
        input_coordinate(1, latitude)
        input_coordinate(1, longitude)  
        auto.press('tab', interval=0.2, presses=18)  
        auto.press('enter') 
        time.sleep(0.2)
    elif output_projection == "Geolagacy":
        input_coordinate(15, zone)  
        input_coordinate(1, latitude)
        input_coordinate(1, longitude)  
        auto.press('tab', interval=0.2, presses=18)  
        auto.press('enter') 
        time.sleep(0.2)
    else:
        input_coordinate(14, zone)  
        input_coordinate(1, latitude)
        input_coordinate(1, longitude)  
        auto.press('tab', interval=0.2, presses=17)  
        auto.press('enter') 
        time.sleep(0.2)


def open_application_and_workspace(xml_file_path, window_title):
    """Open the application and load the workspace."""
    app_path = read_app_path_from_xml(xml_file_path)
    open_application(app_path)

    if activate_window(window_title):
        click_file_menu() 
        select_workspace_open_option() 
        workspace_file_path = read_workspace_file_from_xml(xml_file_path)
        select_workspace_file(workspace_file_path) 
        print("Application window is open")
    else:
        print("APPLICATION WINDOW NOT FOUND")


def process_coordinate(coordinate, iteration_count, total_coordinates):
    """Process a single coordinate."""
    input_projection = coordinate['input_projection']
    output_projection = coordinate['output_projection']
    print(f"Processing input projection: {input_projection}")
    select_input_projection(input_projection, 4)
    print(f"Processing output projection: {output_projection}")
    select_output_projection(input_projection, output_projection)  
    coordinate_value((coordinate['zone'], coordinate['lat'], coordinate['long']), input_projection, output_projection)
    output_values = copy_output_values(input_projection, output_projection)

    if len(output_values) >= 3:  
        output_zone = output_values[0]  
        output_lat = output_values[1]
        output_long = output_values[2]
        user_lat = coordinate['East']
        user_long = coordinate['North']
        validation_status = validate_output(output_lat, output_long, user_lat, user_long)
        output_csv_file_path = read_output_csv_file_path_from_xml(xml_file_path, 'OutputCSVFilePathCC')
        write_output_to_csv(output_csv_file_path, output_zone, output_lat, output_long, validation_status)
    else:
        print("Not enough output values to write to CSV.") 

    
    if(input_projection == "Geolagacy"):
        auto.press('tab', interval=0.2, presses=7)
        auto.press('enter')
        time.sleep(0.2)
    elif(output_projection == "Geolagacy"):
        auto.press('tab', interval=0.2, presses=8)
        auto.press('enter')
        time.sleep(0.2)
    else:
        auto.press('tab', interval=0.2, presses=7)
        auto.press('enter')
        time.sleep(0.2)

    # Check if there are more coordinates to process
    if iteration_count + 1 < total_coordinates:
        click_tools_button()
    else:
        print("All coordinates have been processed.")

    click_file_menu()


def calculate_coordinate_calculator(xml_file_path, window_title):
    """Main function to calculate coordinates using the application."""
    open_application_and_workspace(xml_file_path, window_title)
    click_tools_button()  
    csv_file_path = read_csv_file_path_from_xml(xml_file_path, 'InputCSVFilePathCC')
    coordinates = read_coordinates_from_csv(csv_file_path)

    total_coordinates = len(coordinates)

    for index, coordinate in enumerate(coordinates):
        process_coordinate(coordinate, iteration_count=index, total_coordinates=total_coordinates)

    print("Finished processing all coordinates.")
  

if __name__ == "__main__":
    auto.FAILSAFE = False
    window_title = "TacticalGis Test Application"
    close_button_position = (1219, 280)  # Path for the output CSV file
    xml_file_path = 'C:\\Users\\ec.pkjha\\Desktop\\AutoGUI\\AutoGUIFiles.xml'
    close_tactical_tester = (1983, 9)
    try:
        calculate_coordinate_calculator(xml_file_path, window_title)
        time.sleep(2)           
        auto.moveTo(close_tactical_tester[0], close_tactical_tester[1], duration=0.2)
        auto.click()
    except Exception as e:
         print(f"An error occurred: {e}")







    