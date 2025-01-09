import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import ttk
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, IAudioSessionManager2, IAudioSessionControl
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
import json
import os
import pystray
from PIL import Image, ImageDraw, ImageTk
import keyboard

# Function to get the list of audio sessions
def get_audio_sessions():
    sessions = AudioUtilities.GetAllSessions()
    return [session for session in sessions if session.Process and session.Process.name()]

# Function to get the master volume of Windows
def get_master_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    return volume

# Function to set the volume of a specific session
def set_session_volume(session, volume):
    volume_control = session.SimpleAudioVolume
    volume_control.SetMasterVolume(volume, None)

# Function to set the master volume of Windows
def set_master_volume(volume):
    master_volume.SetMasterVolumeLevelScalar(volume, None)

# Function to update the list of audio sessions
def refresh_sessions():
    global sessions
    sessions = get_audio_sessions()
    session_listbox.delete(0, tk.END)
    session_listbox.insert(tk.END, "Master Volume")
    for session in sessions:
        session_listbox.insert(tk.END, session.Process.name())

# Function to handle volume control for the selected session
def control_selected_session(event=None):
    selected_index = session_listbox.curselection()
    if selected_index:
        if selected_index[0] == 0:
            current_volume = master_volume.GetMasterVolumeLevelScalar()
            if event.keysym == volume_up_key.get():
                set_master_volume(min(1.0, current_volume + 0.05))
            elif event.keysym == volume_down_key.get():
                set_master_volume(max(0.0, current_volume - 0.05))
            # Update the slider value
            volume_slider.set(master_volume.GetMasterVolumeLevelScalar() * 100)
        else:
            selected_session = sessions[selected_index[0] - 1]
            current_volume = selected_session.SimpleAudioVolume.GetMasterVolume()
            if event.keysym == volume_up_key.get():
                set_session_volume(selected_session, min(1.0, current_volume + 0.05))
            elif event.keysym == volume_down_key.get():
                set_session_volume(selected_session, max(0.0, current_volume - 0.05))
            # Update the slider value
            volume_slider.set(selected_session.SimpleAudioVolume.GetMasterVolume() * 100)

# Function to update the volume based on the slider value
def update_volume(val):
    selected_index = session_listbox.curselection()
    if selected_index:
        if selected_index[0] == 0:
            set_master_volume(float(val) / 100)
        else:
            selected_session = sessions[selected_index[0] - 1]
            set_session_volume(selected_session, float(val) / 100)
    else:
        selected_index = favorites_listbox.curselection()
        if selected_index:
            favorite_name = favorites[selected_index[0]]
            if favorite_name == "Master Volume":
                set_master_volume(float(val) / 100)
            else:
                for session in sessions:
                    if session.Process.name() == favorite_name:
                        set_session_volume(session, float(val) / 100)
                        break

# Function to cycle through sessions
def cycle_sessions(event=None):
    current_selection = session_listbox.curselection()
    if current_selection:
        next_selection = (current_selection[0] + 1) % session_listbox.size()
        session_listbox.selection_clear(0, tk.END)
        session_listbox.selection_set(next_selection)
        session_listbox.activate(next_selection)

# Function to add a session to favorites
def add_to_favorites():
    selected_index = session_listbox.curselection()
    if selected_index:
        selected_session_name = "Master Volume" if selected_index[0] == 0 else sessions[selected_index[0] - 1].Process.name()
        if selected_session_name not in favorites:
            favorites.append(selected_session_name)
            save_favorites()
            refresh_favorites()

# Function to remove a session from favorites
def remove_from_favorites():
    selected_index = favorites_listbox.curselection()
    if selected_index:
        favorites.pop(selected_index[0])
        save_favorites()
        refresh_favorites()

# Function to save favorites to file
def save_favorites():
    with open('favorites.json', 'w') as f:
        json.dump(favorites, f)

# Function to save buttons to file
def save_buttons():
    buttons = {
            "volume_up": volume_up_key.get(),
            "volume_down": volume_down_key.get(),
            "cycle_session": cycle_key.get(),
            "cycle_favorites_session": cycle_favorites_key.get()
    }
    with open('favorites_buttons.json', 'w') as f:
        json.dump(buttons, f)

# Function to load favorites from file
def load_favorites():
    if os.path.exists('favorites.json'):
        with open('favorites.json', 'r') as f:
            return json.load(f)
    return []

# Function to load buttons from file
def load_buttons():
    buttons = {
            "volume_up": 'f18',
            "volume_down": 'f17',
            "cycle_session": 'f20',
            "cycle_favorites_session": 'f19'
    }
    if os.path.exists('favorites_buttons.json'):
        with open('favorites_buttons.json', 'r') as f:
            return json.load(f)
    return buttons

def default_buttons():
    volume_up_key.set('f18')
    volume_down_key.set('f17')
    cycle_key.set('f20')
    cycle_favorites_key.set('f19')
    save_buttons()

# Function to refresh the list of favorites
def refresh_favorites():
    favorites_listbox.delete(0, tk.END)
    for favorite in favorites:
        favorites_listbox.insert(tk.END, favorite)

# Function to update the timer
def update_timer(seconds_left):
    timer_label.config(text=f"Time remaining for refresh: {seconds_left}s")
    if seconds_left > 0:
        root.after(1000, update_timer, seconds_left - 1)
    else:
        periodic_refresh()

# Function for periodic refresh
def periodic_refresh():
    refresh_sessions()
    refresh_favorites()
    update_timer(60)  # Set the timer to 60 seconds
    # Select Master Volume from favourites
    favorites_listbox.selection_clear(0, tk.END)
    favorites_listbox.selection_set(0)
    favorites_listbox.activate(0)
    volume_slider.set(master_volume.GetMasterVolumeLevelScalar() * 100)

# Function to handle volume control for favorites
def control_favorite_volume(event=None):
    selected_index = favorites_listbox.curselection()
    if selected_index:
        favorite_name = favorites[selected_index[0]]
        if favorite_name == "Master Volume":
            current_volume = master_volume.GetMasterVolumeLevelScalar()
            if event.keysym == volume_up_key.get():
                set_master_volume(min(1.0, current_volume + 0.05))
            elif event.keysym == volume_down_key.get():
                set_master_volume(max(0.0, current_volume - 0.05))
            # Update the slider value
            volume_slider.set(master_volume.GetMasterVolumeLevelScalar() * 100)
        else:
            for session in sessions:
                if session.Process.name() == favorite_name:
                    current_volume = session.SimpleAudioVolume.GetMasterVolume()
                    if event.keysym == volume_up_key.get():
                        set_session_volume(session, min(1.0, current_volume + 0.05))
                    elif event.keysym == volume_down_key.get():
                        set_session_volume(session, max(0.0, current_volume - 0.05))
                    # Update the slider value
                    volume_slider.set(session.SimpleAudioVolume.GetMasterVolume() * 100)
                    break

# Function to create the system tray icon
def create_image():
    # Generate an image for the system tray icon
    width = 64
    height = 64
    image = Image.new('RGB', (width, height), (255, 255, 255))
    dc = ImageDraw.Draw(image)
    dc.rectangle(
        (width // 2 - 10, height // 2 - 10, width // 2 + 10, height // 2 + 10),
        fill=(0, 0, 0))
    return image

def turn_off():
    root.destroy()
    os._exit(0)

# GUI configuration
root = tk.Tk()
root.title("Windows Audio Mixer Control")

# Frame for volume control
frame = ttk.Frame(root, padding="10", style="Custom.TFrame")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Listbox to display audio sessions
session_listbox = tk.Listbox(frame, bg="#f0f0f0", fg="#000000", font=("Helvetica", 12))
session_listbox.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

# Populate the listbox with audio sessions and the master volume of Windows
master_volume = get_master_volume()
sessions = get_audio_sessions()
session_listbox.insert(tk.END, "Master Volume")
for session in sessions:
    session_listbox.insert(tk.END, session.Process.name())

# Volume slider
volume_slider = ttk.Scale(frame, from_=0, to=100, orient='horizontal', command=update_volume, style="Custom.Horizontal.TScale")
volume_slider.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

# Add a label for the timer
timer_label = ttk.Label(frame, text="Time remaining for refresh: 60s", font=("Helvetica", 10), anchor="e", justify="right", style="Custom.TLabel")
timer_label.grid(row=2, column=0, columnspan=2)

# Refresh button to update the list of audio sessions
refresh_button = ttk.Button(frame, text="Refresh Sessions", command=refresh_sessions, style="Custom.TButton")
refresh_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

# Input fields to select keys for volume up and down
buttons = load_buttons()
volume_up_key_label = ttk.Label(frame, text="Volume Up Key:", style="Custom.TLabel")
volume_up_key_label.grid(row=4, column=0, sticky=(tk.W))
volume_up_key = tk.StringVar(value=buttons["volume_up"])
volume_up_key_menu = ttk.Combobox(frame, textvariable=volume_up_key, values=[chr(i) for i in range(32, 127)], style="Custom.TCombobox")
volume_up_key_menu.grid(row=4, column=1, sticky=(tk.W, tk.E))

volume_down_key_label = ttk.Label(frame, text="Volume Down Key:", style="Custom.TLabel")
volume_down_key_label.grid(row=5, column=0, sticky=(tk.W))
volume_down_key = tk.StringVar(value=buttons["volume_down"])
volume_down_key_menu = ttk.Combobox(frame, textvariable=volume_down_key, values=[chr(i) for i in range(32, 127)], style="Custom.TCombobox")
volume_down_key_menu.grid(row=5, column=1, sticky=(tk.W, tk.E))

# Input field to select the key for cycling through sessions
cycle_key_label = ttk.Label(frame, text="Cycle Session Key:", style="Custom.TLabel")
cycle_key_label.grid(row=6, column=0, sticky=(tk.W))
cycle_key = tk.StringVar(value=buttons["cycle_session"])
cycle_key_menu = ttk.Combobox(frame, textvariable=cycle_key, values=[chr(i) for i in range(32, 127)], style="Custom.TCombobox")
cycle_key_menu.grid(row=6, column=1, sticky=(tk.W, tk.E))

# Input field to select the key for cycling through favorites sessions
cycle_favorites_key_label = ttk.Label(frame, text="Cycle Favorites Session Key:", style="Custom.TLabel")
cycle_favorites_key_label.grid(row=7, column=0, sticky=(tk.W))
cycle_favorites_key = tk.StringVar(value=buttons["cycle_favorites_session"])
cycle_favorites_key_menu = ttk.Combobox(frame, textvariable=cycle_favorites_key, values=[chr(i) for i in range(32, 127)], style="Custom.TCombobox")
cycle_favorites_key_menu.grid(row=7, column=1, sticky=(tk.W, tk.E))

save_buttons_button = ttk.Button(frame, text="Save Buttons", command=save_buttons, style="Custom.TButton")
save_buttons_button.grid(row=8, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
save_buttons()
default_buttons_button = ttk.Button(frame, text="Set Default Buttons", command=default_buttons, style="Custom.TButton")
default_buttons_button.grid(row=8, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

# Listbox to display favorites
favorites_listbox = tk.Listbox(frame, bg="#f0f0f0", fg="#000000", font=("Helvetica", 12))
favorites_listbox.grid(row=9, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

# Button to add to favorites
add_favorite_button = ttk.Button(frame, text="Add to Favorites", command=add_to_favorites, style="Custom.TButton")
add_favorite_button.grid(row=10, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

# Button to remove from favorites
remove_favorite_button = ttk.Button(frame, text="Remove from Favorites", command=remove_from_favorites, style="Custom.TButton")
remove_favorite_button.grid(row=10, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

# Add a label on the right side of the GUI
credits_label = ttk.Label(frame, text="Made with love by Mikelesi,\nBegony and Copilot,\nwith the vain attempt of ChatGPT", font=("Helvetica", 10), anchor="e", justify="right", style="Custom.TLabel")
credits_label.grid(row=11, column=0, columnspan=2)

# Load favorites from file
favorites = load_favorites()
refresh_favorites()

# Configure dynamic resizing
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
frame.columnconfigure(0, weight=1)
frame.columnconfigure(1, weight=1)
frame.rowconfigure(0, weight=1)
frame.rowconfigure(6, weight=1)

# Start the initial timer
update_timer(60)

# Start the event loop of the GUI
root.after(60000, periodic_refresh)  # Start the periodic refresh

# Start the system tray icon
#setup_tray_icon()
icon = pystray.Icon("audio_mixer")
icon.icon = create_image()
icon.title = "Windows Audio Mixer Control"
icon.run_detached()

def change_image(image):
    icon.icon = image
    icon.update_menu()

# Function to change the selected session (assume that a session is always selected)
def change_session(event, use_favorites=False):
    if use_favorites:
        selected_index = favorites_listbox.curselection()
        if selected_index:
            next_index = (selected_index[0] + 1) % favorites_listbox.size()
        else:
            next_index = 0
        favorites_listbox.selection_clear(0, tk.END)
        favorites_listbox.selection_set(next_index)
        favorites_listbox.activate(next_index)
        if favorites_listbox.get(next_index) == "Master Volume":
            volume_slider.set(master_volume.GetMasterVolumeLevelScalar() * 100)
        else:
            for session in sessions:
                if session.Process.name() == favorites_listbox.get(next_index):
                    volume_slider.set(session.SimpleAudioVolume.GetMasterVolume() * 100)
                    break

    else:
        selected_index = session_listbox.curselection()
        if selected_index:
            next_index = (selected_index[0] + 1) % session_listbox.size()
        else:
            next_index = 0
        session_listbox.selection_clear(0, tk.END)
        session_listbox.selection_set(next_index)
        session_listbox.activate(next_index)
        if next_index == 0:
            volume_slider.set(master_volume.GetMasterVolumeLevelScalar() * 100)
        else:
            selected_session = sessions[next_index - 1]
            volume_slider.set(selected_session.SimpleAudioVolume.GetMasterVolume() * 100)

def my_control_selected_session(sign):
    selected_index = session_listbox.curselection()
    if selected_index:
        selected_session = sessions[selected_index[0] - 1]
        current_volume = selected_session.SimpleAudioVolume.GetMasterVolume()
        if sign == '+':
            set_session_volume(selected_session, min(1.0, current_volume + 0.05))
        elif sign == '-':
            set_session_volume(selected_session, max(0.0, current_volume - 0.05))
        # Update the slider value
        volume_slider.set(selected_session.SimpleAudioVolume.GetMasterVolume() * 100)

def my_favorite_volume(sign):
    selected_index = favorites_listbox.curselection()
    if selected_index:
        favorite_name = favorites[selected_index[0]]
        if favorite_name == "Master Volume":
            current_volume = master_volume.GetMasterVolumeLevelScalar()
            if sign == '+':
                set_master_volume(min(1.0, current_volume + 0.05))
            elif sign == '-':
                set_master_volume(max(0.0, current_volume - 0.05))
            # Update the slider value
            volume_slider.set(master_volume.GetMasterVolumeLevelScalar() * 100)
        else:
            for session in sessions:
                if session.Process.name() == favorite_name:
                    current_volume = session.SimpleAudioVolume.GetMasterVolume()
                    if sign == '+':
                        set_session_volume(session, min(1.0, current_volume + 0.05))
                    elif sign == '-':
                        set_session_volume(session, max(0.0, current_volume - 0.05))
                    # Update the slider value
                    volume_slider.set(session.SimpleAudioVolume.GetMasterVolume() * 100)
                    break

def onKeyEvent(event):
    if event.event_type == keyboard.KEY_UP:
        return
    # Pass the pressed key to the application
    if event.name == volume_up_key.get():
        # Increase the volume of the selected session
        if session_listbox.curselection():
            my_control_selected_session('+')
        elif favorites_listbox.curselection():
            my_favorite_volume('+')
    elif event.name == volume_down_key.get():
        # Decrease the volume of the selected session
        if session_listbox.curselection():
            my_control_selected_session('-')
        elif favorites_listbox.curselection():
            my_favorite_volume('-')
    elif event.name == cycle_key.get():
        change_session(event, use_favorites=False)
    elif event.name == cycle_favorites_key.get():
        change_session(event, use_favorites=True)

keyboard.hook(onKeyEvent)

# Bind the events to the listboxes
session_listbox.bind("<ButtonRelease-1>", control_selected_session)
session_listbox.bind("<Key>", control_selected_session)
session_listbox.bind("<Return>", control_selected_session)
session_listbox.bind("<Down>", lambda event: change_session(event, use_favorites=False))
session_listbox.bind("<Up>", lambda event: change_session(event, use_favorites=False))
favorites_listbox.bind("<ButtonRelease-1>", control_favorite_volume)
favorites_listbox.bind("<Key>", control_favorite_volume)
favorites_listbox.bind("<Return>", control_favorite_volume)
favorites_listbox.bind("<Down>", lambda event: change_session(event, use_favorites=True))
favorites_listbox.bind("<Up>", lambda event: change_session(event, use_favorites=True))

favorites_listbox.selection_clear(0, tk.END)
favorites_listbox.selection_set(0)
favorites_listbox.activate(0)
volume_slider.set(master_volume.GetMasterVolumeLevelScalar() * 100)

root.protocol("WM_DELETE_WINDOW", turn_off)
root.mainloop()
