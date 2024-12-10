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

# Funzione per ottenere la lista delle sessioni audio
def get_audio_sessions():
    sessions = AudioUtilities.GetAllSessions()
    return [session for session in sessions if session.Process and session.Process.name()]

# Funzione per ottenere il volume generale di Windows
def get_master_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    return volume

# Funzione per impostare il volume di una specifica sessione
def set_session_volume(session, volume):
    volume_control = session.SimpleAudioVolume
    volume_control.SetMasterVolume(volume, None)

# Funzione per impostare il volume generale di Windows
def set_master_volume(volume):
    master_volume.SetMasterVolumeLevelScalar(volume, None)

# Funzione per aggiornare la lista delle sessioni audio
def refresh_sessions():
    global sessions
    sessions = get_audio_sessions()
    session_listbox.delete(0, tk.END)
    session_listbox.insert(tk.END, "Master Volume")
    for session in sessions:
        session_listbox.insert(tk.END, session.Process.name())

# Funzione per gestire il controllo del volume per la sessione selezionata
def control_selected_session(event=None):
    selected_index = session_listbox.curselection()
    if selected_index:
        if selected_index[0] == 0:
            current_volume = master_volume.GetMasterVolumeLevelScalar()
            if event.keysym == volume_up_key.get():
                set_master_volume(min(1.0, current_volume + 0.05))
            elif event.keysym == volume_down_key.get():
                set_master_volume(max(0.0, current_volume - 0.05))
            # Aggiorna il valore dello slider
            volume_slider.set(master_volume.GetMasterVolumeLevelScalar() * 100)
        else:
            selected_session = sessions[selected_index[0] - 1]
            current_volume = selected_session.SimpleAudioVolume.GetMasterVolume()
            if event.keysym == volume_up_key.get():
                set_session_volume(selected_session, min(1.0, current_volume + 0.05))
            elif event.keysym == volume_down_key.get():
                set_session_volume(selected_session, max(0.0, current_volume - 0.05))
            # Aggiorna il valore dello slider
            volume_slider.set(selected_session.SimpleAudioVolume.GetMasterVolume() * 100)

# Funzione per aggiornare il volume in base al valore dello slider
def update_volume(val):
    selected_index = session_listbox.curselection()
    if selected_index:
        if selected_index[0] == 0:
            set_master_volume(float(val) / 100)
        else:
            selected_session = sessions[selected_index[0] - 1]
            set_session_volume(selected_session, float(val) / 100)

# Funzione per cambiare la sessione in modo ciclico
def cycle_sessions(event=None):
    current_selection = session_listbox.curselection()
    if current_selection:
        next_selection = (current_selection[0] + 1) % session_listbox.size()
        session_listbox.selection_clear(0, tk.END)
        session_listbox.selection_set(next_selection)
        session_listbox.activate(next_selection)

# Funzione per aggiungere una sessione ai preferiti
def add_to_favorites():
    selected_index = session_listbox.curselection()
    if selected_index:
        selected_session_name = "Master Volume" if selected_index[0] == 0 else sessions[selected_index[0] - 1].Process.name()
        if selected_session_name not in favorites:
            favorites.append(selected_session_name)
            save_favorites()
            refresh_favorites()

# Funzione per rimuovere una sessione dai preferiti
def remove_from_favorites():
    selected_index = favorites_listbox.curselection()
    if selected_index:
        favorites.pop(selected_index[0])
        save_favorites()
        refresh_favorites()

# Funzione per salvare i preferiti su file
def save_favorites():
    with open('favorites.json', 'w') as f:
        json.dump(favorites, f)

# Funzione per caricare i preferiti da file
def load_favorites():
    if os.path.exists('favorites.json'):
        with open('favorites.json', 'r') as f:
            return json.load(f)
    return []

# Funzione per aggiornare la lista dei preferiti
def refresh_favorites():
    favorites_listbox.delete(0, tk.END)
    for favorite in favorites:
        favorites_listbox.insert(tk.END, favorite)

# Funzione per aggiornare il timer
def update_timer(seconds_left):
    timer_label.config(text=f"Tempo rimanente al refresh: {seconds_left}s")
    if seconds_left > 0:
        root.after(1000, update_timer, seconds_left - 1)
    else:
        periodic_refresh()

# Funzione per il refresh ciclico
def periodic_refresh():
    refresh_sessions()
    refresh_favorites()
    update_timer(60)  # Imposta il timer a 60 secondi

# Funzione per gestire il controllo del volume per i preferiti
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
            # Aggiorna il valore dello slider
            volume_slider.set(master_volume.GetMasterVolumeLevelScalar() * 100)
        else:
            for session in sessions:
                if session.Process.name() == favorite_name:
                    current_volume = session.SimpleAudioVolume.GetMasterVolume()
                    if event.keysym == volume_up_key.get():
                        set_session_volume(session, min(1.0, current_volume + 0.05))
                    elif event.keysym == volume_down_key.get():
                        set_session_volume(session, max(0.0, current_volume - 0.05))
                    # Aggiorna il valore dello slider
                    volume_slider.set(session.SimpleAudioVolume.GetMasterVolume() * 100)
                    break

# Funzione per creare l'icona di sistema
def create_image():
    # Genera un'immagine per l'icona di sistema
    width = 64
    height = 64
    image = Image.new('RGB', (width, height), (255, 255, 255))
    dc = ImageDraw.Draw(image)
    dc.rectangle(
        (width // 2 - 10, height // 2 - 10, width // 2 + 10, height // 2 + 10),
        fill=(0, 0, 0))
    return image

# Funzione per avviare l'icona di sistema
def setup_tray_icon():
    icon = pystray.Icon("audio_mixer")
    icon.icon = create_image()
    icon.title = "Controllo Mixer Audio di Windows"
    icon.run_detached()

# Configurazione della GUI
root = tk.Tk()
root.title("Controllo Mixer Audio di Windows")

# Aggiungi un'immagine di sfondo
background_image = Image.open("background.jpg")
background_photo = ImageTk.PhotoImage(background_image)
background_label = tk.Label(root, image=background_photo)
background_label.image = background_photo  # Mantieni un riferimento all'immagine
background_label.place(relwidth=1, relheight=1)

# Frame per il controllo del volume
frame = ttk.Frame(root, padding="10", style="Custom.TFrame")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Listbox per visualizzare le sessioni audio
session_listbox = tk.Listbox(frame, bg="#f0f0f0", fg="#000000", font=("Helvetica", 12))
session_listbox.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

# Popola la listbox con le sessioni audio e il volume generale di Windows
master_volume = get_master_volume()
sessions = get_audio_sessions()
session_listbox.insert(tk.END, "Master Volume")
for session in sessions:
    session_listbox.insert(tk.END, session.Process.name())

# Slider del volume
volume_slider = ttk.Scale(frame, from_=0, to=100, orient='horizontal', command=update_volume, style="Custom.Horizontal.TScale")
volume_slider.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

# Bottone di refresh per aggiornare la lista delle sessioni audio
refresh_button = ttk.Button(frame, text="Aggiorna Sessioni", command=refresh_sessions, style="Custom.TButton")
refresh_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))

# Campi di input per selezionare i tasti per volume up e down
volume_up_key_label = ttk.Label(frame, text="Tasto Volume Su:", style="Custom.TLabel")
volume_up_key_label.grid(row=3, column=0, sticky=(tk.W))
volume_up_key = tk.StringVar(value='f18')
volume_up_key_menu = ttk.Combobox(frame, textvariable=volume_up_key, values=[chr(i) for i in range(32, 127)], style="Custom.TCombobox")
volume_up_key_menu.grid(row=3, column=1, sticky=(tk.W, tk.E))

volume_down_key_label = ttk.Label(frame, text="Tasto Volume Giù:", style="Custom.TLabel")
volume_down_key_label.grid(row=4, column=0, sticky=(tk.W))
volume_down_key = tk.StringVar(value='f17')
volume_down_key_menu = ttk.Combobox(frame, textvariable=volume_down_key, values=[chr(i) for i in range(32, 127)], style="Custom.TCombobox")
volume_down_key_menu.grid(row=4, column=1, sticky=(tk.W, tk.E))

# Campi di input per selezionare il tasto per cambiare le sessioni in modo ciclico
cycle_key_label = ttk.Label(frame, text="Tasto Cambio Sessione:", style="Custom.TLabel")
cycle_key_label.grid(row=5, column=0, sticky=(tk.W))
cycle_key = tk.StringVar(value='Tab')
cycle_key_menu = ttk.Combobox(frame, textvariable=cycle_key, values=[chr(i) for i in range(32, 127)], style="Custom.TCombobox")
cycle_key_menu.grid(row=5, column=1, sticky=(tk.W, tk.E))

# Listbox per visualizzare i preferiti
favorites_listbox = tk.Listbox(frame, bg="#f0f0f0", fg="#000000", font=("Helvetica", 12))
favorites_listbox.grid(row=6, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

# Bottone per aggiungere ai preferiti
add_favorite_button = ttk.Button(frame, text="Aggiungi ai Preferiti", command=add_to_favorites, style="Custom.TButton")
add_favorite_button.grid(row=7, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

# Bottone per rimuovere dai preferiti
remove_favorite_button = ttk.Button(frame, text="Rimuovi dai Preferiti", command=remove_from_favorites, style="Custom.TButton")
remove_favorite_button.grid(row=7, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

# Aggiungi una scritta sul lato destro della GUI
credits_label = ttk.Label(root, text="Realizzato con amore da Mikelesi,\nBegony e Copilot,\ncon il vano tentativo di ChatGPT", font=("Helvetica", 10), anchor="e", justify="right", style="Custom.TLabel")
credits_label.place(relx=0.95, rely=0.1, anchor='ne')

# Aggiungi un'etichetta per il timer
timer_label = ttk.Label(root, text="Tempo rimanente al refresh: 60s", font=("Helvetica", 10), anchor="e", justify="right", style="Custom.TLabel")
timer_label.place(relx=0.95, rely=0.2, anchor='ne')

# Carica i preferiti da file
favorites = load_favorites()
refresh_favorites()

# Associa la rotella dedicata della tastiera alle funzioni di controllo del volume per la sessione selezionata
root.bind('<KeyPress>', control_selected_session)
root.bind(f'<KeyPress-{cycle_key.get()}>', cycle_sessions)
favorites_listbox.bind('<KeyPress>', control_favorite_volume)

# Configura il ridimensionamento dinamico
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
frame.columnconfigure(0, weight=1)
frame.columnconfigure(1, weight=1)
frame.rowconfigure(0, weight=1)
frame.rowconfigure(6, weight=1)

# Avvia il timer iniziale
update_timer(60)

# Avvia il loop degli eventi della GUI
root.after(60000, periodic_refresh)  # Avvia il refresh ciclico

# Avvia l'icona di sistema
setup_tray_icon()

# Funzione per cambiare la sessione selezionata
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
    else:
        selected_index = session_listbox.curselection()
        if selected_index:
            next_index = (selected_index[0] + 1) % session_listbox.size()
        else:
            next_index = 0
        session_listbox.selection_clear(0, tk.END)
        session_listbox.selection_set(next_index)
        session_listbox.activate(next_index)

# Associa il tasto cambio sessione alla funzione change_session per i preferiti
root.bind('<KeyPress-F19>', lambda event: change_session(event, use_favorites=True))

# Associa il tasto cambio sessione alla funzione change_session per tutte le applicazioni
# root.bind('<KeyPress-5>', lambda event: change_session(event, use_favorites=False))

def my_control_selected_session(segno):
    selected_index = session_listbox.curselection()
    if selected_index:
        selected_session = sessions[selected_index[0] - 1]  # Corretto l'indice
        current_volume = selected_session.SimpleAudioVolume.GetMasterVolume()
        if segno == '+':
            set_session_volume(selected_session, min(1.0, current_volume + 0.05))
        elif segno == '-':
            set_session_volume(selected_session, max(0.0, current_volume - 0.05))
        # Aggiorna il valore dello slider
        volume_slider.set(selected_session.SimpleAudioVolume.GetMasterVolume() * 100)

def my_favorite_volume(segno):
    selected_index = favorites_listbox.curselection()
    if selected_index:
        favorite_name = favorites[selected_index[0]]
        if favorite_name == "Master Volume":
            current_volume = master_volume.GetMasterVolumeLevelScalar()
            if segno == '+':
                set_master_volume(min(1.0, current_volume + 0.05))
            elif segno == '-':
                set_master_volume(max(0.0, current_volume - 0.05))
            # Aggiorna il valore dello slider
            volume_slider.set(master_volume.GetMasterVolumeLevelScalar() * 100)
        else:
            for session in sessions:
                if session.Process.name() == favorite_name:
                    current_volume = session.SimpleAudioVolume.GetMasterVolume()
                    if segno == '+':
                        set_session_volume(session, min(1.0, current_volume + 0.05))
                    elif segno == '-':
                        set_session_volume(session, max(0.0, current_volume - 0.05))
                    # Aggiorna il valore dello slider
                    volume_slider.set(session.SimpleAudioVolume.GetMasterVolume() * 100)
                    break

def onKeyEvent(event):
    if event.event_type == keyboard.KEY_UP:
        return
    # devo passare il tasto premuto all'applicazione
    if event.name == volume_up_key.get():
        # Alza il volume della sessione selezionata
        if session_listbox.curselection():
            my_control_selected_session('+')
        elif favorites_listbox.curselection():
            my_favorite_volume('+')
    elif event.name == volume_down_key.get():
        # Abbassa il volume della sessione selezionata
        if session_listbox.curselection():
            my_control_selected_session('-')
        elif favorites_listbox.curselection():
            my_favorite_volume('-')
    elif event.name == 'f19':
        change_session(event, use_favorites=True)

keyboard.hook(onKeyEvent)

root.mainloop()