# AudioMixerBegelesy
Audio Mixer for Windows by Mikelesi and JohnBronx01

# Audio Mixer 1.0
Audio Mixer 1.0 is a Python application that allows you to control the audio volume of different applications on your Windows system.\
It provides a graphical user interface (GUI) to manage audio sessions, adjust volumes, and manage favorite audio sessions.

## Features
- List all active audio sessions.
- Adjust the master volume of the system.
- Adjust the volume of individual audio sessions.
- Add and remove audio sessions from favorites.
- Automatically refresh the list of audio sessions and favorites.
- Display a timer indicating the time remaining until the next automatic refresh.
- Control volume and switch sessions using keyboard shortcuts, even when the application is in the background.

## Requirements
- Python 3.x
- pycaw library
- comtypes library
- pystray library
- Pillow library
- keyboard library

## Installation
1) Clone the repository or download the audiomixer1.0.py file.
2) Install the required libraries using pip:
```sh
pip install pycaw comtypes pystray pillow keyboard
```

## Usage
Run the audiomixer1.0.py file:
``` sh
python audiomixer1.0.py
```
The application window will open, displaying the list of active audio sessions and the master volume.

Use the following keyboard shortcuts to control the application:
- Volume Up: F18 (default, can be changed in the GUI)
- Volume Down: F17 (default, can be changed in the GUI)
- Cycle Sessions: Tab (default, can be changed in the GUI)
- Change Favorite Session: F19

You can add or remove audio sessions from favorites using the buttons provided in the GUI.

The application will automatically refresh the list of audio sessions and favorites every 60 seconds.\
The timer indicating the time remaining until the next refresh is displayed in the GUI.

## GUI Components
- Session Listbox: Displays the list of active audio sessions and the master volume.
- Volume Slider: Allows you to adjust the volume of the selected audio session or the master volume.
- Refresh Button: Manually refreshes the list of audio sessions.
- Favorite Listbox: Displays the list of favorite audio sessions.
- Add to Favorites Button: Adds the selected audio session to the favorites list.
- Remove from Favorites Button: Removes the selected audio session from the favorites list.
- Timer Label: Displays the time remaining until the next automatic refresh.
- Credits Label: Displays the credits for the application.

## Notes
Ensure that the background.jpg file is present in the same directory as the [AudioMixer1.0.py](AudioMixer1.0.py) file for the background image to be displayed correctly.\
The application uses the keyboard library to intercept keyboard events at the system level, allowing it to respond to keyboard shortcuts even when it is not the active window.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgements
- [pycaw](https://pypi.org/project/pycaw/) - Python Core Audio Windows Library
- [comtypes](https://pypi.org/project/comtypes/) - A pure Python COM package
- [pystray](https://pypi.org/project/pystray/) - A cross-platform system tray icon library
- [Pillow](https://pypi.org/project/pillow/) - Python Imaging Library (PIL) fork
- [keyboard](https://pypi.org/project/keyboard/) - Hook and simulate global keyboard events on Windows and Linux

```
This README provides an overview of the application's features, installation instructions, usage, and details about the GUI components. It also includes notes on the requirements and acknowledgements for the libraries used in the project. 
```
